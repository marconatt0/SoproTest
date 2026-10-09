import difflib
import io
import re
import unicodedata

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import openpyxl

st.set_page_config(page_title="Painel de Controle - Estações", layout="wide")

# ---------------------------------------------------------------------------
# Classificação dos status do portal
# ---------------------------------------------------------------------------
STATUS_RESOLVIDO = ["concluído", "concluido", "fechado"]
STATUS_CANCELADO = ["cancelado"]
# Status que não entram em nenhum cálculo nem relatório
STATUS_EXCLUIDOS = ["aguardando número do chamado", "cancelado pela claro", "fechado pelo mso", "pendente claro"]

CATEGORIA_CORES = {"Resolvido": "#2ECC71", "Aberto": "#FFA500", "Cancelado": "#95A5A6"}
SLA_CORES = {"NO PRAZO": "#2ECC71", "FORA DO PRAZO": "#E74C3C", "SEM SLA": "#BDC3C7"}


def normalizar_texto(texto):
    """Minúsculas, sem acentos e com qualquer pontuação/espaço (inclusive invisível) virando um espaço só."""
    texto = unicodedata.normalize('NFKD', str(texto))
    texto = ''.join(c for c in texto if not unicodedata.combining(c)).lower()
    return re.sub(r'[^a-z0-9]+', ' ', texto).strip()


STATUS_EXCLUIDOS_NORM = [normalizar_texto(s) for s in STATUS_EXCLUIDOS]


def status_excluido(status):
    s = normalizar_texto(status)
    return any(k in s for k in STATUS_EXCLUIDOS_NORM)


def categorizar_status(status):
    s = str(status).lower()
    if any(k in s for k in STATUS_RESOLVIDO):
        return "Resolvido"
    if any(k in s for k in STATUS_CANCELADO):
        return "Cancelado"
    return "Aberto"


def cor_sla(cell):
    cor_hex = None
    if cell.font and cell.font.color and cell.font.color.type == 'rgb':
        cor_hex = str(cell.font.color.rgb).upper()
    elif cell.fill and cell.fill.fgColor and cell.fill.fgColor.type == 'rgb':
        cor_hex = str(cell.fill.fgColor.rgb).upper()

    if cor_hex:
        # Se a cor em HEX (RGB) tiver código de Verde (usado pelo Excel)
        if "00B050" in cor_hex or "00FF00" in cor_hex or "2ECC71" in cor_hex or "008000" in cor_hex:
            return "NO PRAZO"
        # Se a cor tiver código de Vermelho
        if "FF0000" in cor_hex or "C00000" in cor_hex or "FF4500" in cor_hex or "E74C3C" in cor_hex:
            return "FORA DO PRAZO"
    return "SEM SLA"


@st.cache_data(show_spinner="Lendo planilha...")
def carregar_dados(conteudo):
    # 1. O pandas lê apenas o texto (onde aparece o ●)
    df = pd.read_excel(io.BytesIO(conteudo), header=1)

    # 2. O openpyxl lê a cor da célula de SLA
    wb = openpyxl.load_workbook(io.BytesIO(conteudo), data_only=True)
    ws = wb.active

    sla_col_idx = None
    for idx, cell in enumerate(ws[2], 1):
        if cell.value and str(cell.value).strip().upper() == 'SLA':
            sla_col_idx = idx
            break

    status_sla_cores = []
    for row_idx in range(3, ws.max_row + 1):
        if sla_col_idx:
            status_sla_cores.append(cor_sla(ws.cell(row=row_idx, column=sla_col_idx)))
        else:
            status_sla_cores.append("SEM SLA")

    status_sla_cores = status_sla_cores[:len(df)]
    if len(status_sla_cores) < len(df):
        status_sla_cores += ["SEM SLA"] * (len(df) - len(status_sla_cores))

    df['SLA'] = status_sla_cores

    # A planilha tem duas colunas "Data": abertura e acionamento (a 2ª vira "Data.1")
    df = df.rename(columns={'Data.1': 'Data Acionamento'})

    df = df.dropna(subset=['Estação'])
    # Colunas com números e textos misturados (ex.: "1213/ 1214") viram texto para exibição
    for col in ['ID Atividade', 'ID Fornecedor', 'Chamado']:
        if col in df.columns:
            df[col] = df[col].apply(lambda v: '' if pd.isna(v) else str(int(v)) if isinstance(v, float) and v.is_integer() else str(v))
    df['Status'] = df['Status'].astype(str).str.strip()
    df = df[~df['Status'].apply(status_excluido)]
    df['Empresa'] = df['Empresa'].fillna('NÃO INFORMADA').astype(str).str.strip()
    if 'Tipo' in df.columns:
        df['Tipo'] = df['Tipo'].fillna('NÃO INFORMADO').astype(str).str.strip()

    for col in ['Data', 'Data Acionamento', 'Data Fechamento']:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    df['Categoria'] = df['Status'].apply(categorizar_status)

    if 'Data' in df.columns:
        df['Mês Abertura'] = df['Data'].dt.to_period('M').dt.to_timestamp()

    # Tempos em dias (valores negativos são inconsistências de digitação e são descartados)
    if 'Data' in df.columns and 'Data Fechamento' in df.columns:
        dias = (df['Data Fechamento'] - df['Data']).dt.days
        df['Dias p/ Resolver'] = dias.where((dias >= 0) & (df['Categoria'] == 'Resolvido'))
        df['Mês Fechamento'] = df['Data Fechamento'].dt.to_period('M').dt.to_timestamp()
    if 'Data' in df.columns and 'Data Acionamento' in df.columns:
        dias = (df['Data Acionamento'] - df['Data']).dt.days
        df['Dias p/ Acionar'] = dias.where(dias >= 0)

    return df


def faixa_aging(dias):
    if pd.isna(dias):
        return "Sem data"
    if dias <= 30:
        return "0-30 dias"
    if dias <= 60:
        return "31-60 dias"
    if dias <= 90:
        return "61-90 dias"
    if dias <= 180:
        return "91-180 dias"
    return "+180 dias"


FAIXAS_AGING = ["0-30 dias", "31-60 dias", "61-90 dias", "91-180 dias", "+180 dias", "Sem data"]


def pct(parte, total):
    return (parte / total * 100) if total else 0.0


def resumo_por_grupo(df, grupo, data_ref):
    """Indicadores de desempenho agrupados por uma coluna (Empresa, Estação, Tipo...)."""
    g = df.groupby(grupo)
    r = pd.DataFrame({
        'Total': g.size(),
        'Resolvidos': g['Categoria'].apply(lambda x: (x == 'Resolvido').sum()),
        'Abertos': g['Categoria'].apply(lambda x: (x == 'Aberto').sum()),
        'Cancelados': g['Categoria'].apply(lambda x: (x == 'Cancelado').sum()),
        'SLA No Prazo': g['SLA'].apply(lambda x: (x == 'NO PRAZO').sum()),
        'SLA Fora do Prazo': g['SLA'].apply(lambda x: (x == 'FORA DO PRAZO').sum()),
    })
    validos = r['Total'] - r['Cancelados']
    r['% Resolução'] = (r['Resolvidos'] / validos.where(validos > 0) * 100).round(1)
    sla_medido = r['SLA No Prazo'] + r['SLA Fora do Prazo']
    r['% SLA No Prazo'] = (r['SLA No Prazo'] / sla_medido.where(sla_medido > 0) * 100).round(1)
    r['% Cancelados'] = (r['Cancelados'] / r['Total'] * 100).round(1)

    if 'Dias p/ Resolver' in df.columns:
        r['Média Dias Resolução'] = g['Dias p/ Resolver'].mean().round(1)
        r['Mediana Dias Resolução'] = g['Dias p/ Resolver'].median().round(1)
        r['P90 Dias Resolução'] = g['Dias p/ Resolver'].quantile(0.9).round(1)
    if 'Dias p/ Acionar' in df.columns:
        r['Mediana Dias Acionamento'] = g['Dias p/ Acionar'].median().round(1)

    abertos = df[df['Categoria'] == 'Aberto']
    if 'Data' in df.columns and not abertos.empty:
        idade = (data_ref - abertos['Data']).dt.days
        ab = abertos.assign(Idade=idade).groupby(grupo)['Idade']
        r['Idade Média Backlog (dias)'] = ab.mean().round(1)
        r['Backlog +90 dias'] = ab.apply(lambda x: (x > 90).sum())
    else:
        r['Idade Média Backlog (dias)'] = pd.NA
        r['Backlog +90 dias'] = 0
    r['Backlog +90 dias'] = r['Backlog +90 dias'].fillna(0).astype(int)

    return r.reset_index()


def calcular_score(r):
    """Índice de eficiência 0-100: 40% resolução, 30% SLA, 30% rapidez (mediana de dias, relativa)."""
    resol = r['% Resolução'].fillna(0) / 100
    sla = r['% SLA No Prazo'].fillna(0) / 100
    if 'Mediana Dias Resolução' in r.columns and r['Mediana Dias Resolução'].notna().any():
        med = r['Mediana Dias Resolução']
        mn, mx = med.min(), med.max()
        rapidez = (1 - (med - mn) / (mx - mn)) if mx > mn else pd.Series(1.0, index=r.index)
        rapidez = rapidez.fillna(0)
    else:
        rapidez = pd.Series(0.0, index=r.index)
    return ((0.4 * resol + 0.3 * sla + 0.3 * rapidez) * 100).round(1)


def historico_mensal(df, inicio=None, fim=None):
    """Abertos x Resolvidos x Cancelados por mês + backlog acumulado e tempo de resolução.

    inicio/fim (primeiro dia do mês) limitam o eixo aos meses do período escolhido na barra lateral.
    """
    abertos = df.groupby('Mês Abertura').size().rename('Abertos')
    res = df[df['Categoria'] == 'Resolvido']
    resolvidos = res.groupby('Mês Fechamento').size().rename('Resolvidos')
    canc = df[df['Categoria'] == 'Cancelado']
    cancelados = canc.groupby('Mês Fechamento').size().rename('Cancelados')
    mediana = res.groupby('Mês Fechamento')['Dias p/ Resolver'].median().rename('Mediana Dias Resolução')
    media = res.groupby('Mês Fechamento')['Dias p/ Resolver'].mean().rename('Média Dias Resolução')

    h = pd.concat([abertos, resolvidos, cancelados, mediana, media], axis=1).sort_index()
    h.index.name = 'Mês'
    if h.empty:
        return h.reset_index()
    inicio = h.index.min() if inicio is None else inicio
    fim = h.index.max() if fim is None else fim
    h = h.reindex(pd.date_range(inicio, fim, freq='MS'))
    h.index.name = 'Mês'
    for c in ['Abertos', 'Resolvidos', 'Cancelados']:
        h[c] = h[c].fillna(0).astype(int)
    h['Saldo do Mês'] = h['Abertos'] - h['Resolvidos'] - h['Cancelados']
    h['Backlog Acumulado'] = h['Saldo do Mês'].cumsum()
    h['% Resolvidos s/ Abertos'] = (h['Resolvidos'] / h['Abertos'].where(h['Abertos'] > 0) * 100).round(1)
    h['Mediana Dias Resolução'] = h['Mediana Dias Resolução'].round(1)
    h['Média Dias Resolução'] = h['Média Dias Resolução'].round(1)
    return h.reset_index()


def filtrar_status(df, status_selecionado):
    if status_selecionado != "TODOS":
        return df[df['Status'] == status_selecionado]
    return df


def gerar_excel(abas):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for nome, tabela in abas.items():
            tabela.to_excel(writer, sheet_name=nome[:31], index=False)
            ws = writer.sheets[nome[:31]]
            ws.freeze_panes = 'A2'
            for col in ws.columns:
                largura = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(largura + 2, 10), 50)
    return buffer.getvalue()


def ranking_ocorrencias(df, filtro, titulo_periodo, escala, colunas_detalhe, opcoes_periodo, chave):
    if 'Data' not in df.columns or not pd.api.types.is_datetime64_any_dtype(df['Data']):
        st.error("Não foi possível ler as datas da planilha para gerar este relatório.")
        return

    periodo = st.radio("Selecione o período de análise:", list(opcoes_periodo), horizontal=True, key=chave)
    data_base = df['Data'].max()
    data_corte = data_base - pd.DateOffset(months=opcoes_periodo[periodo])

    df_sel = df[(df['Data'] >= data_corte) & filtro(df)]
    if df_sel.empty:
        st.success(f"Nenhum chamado de {titulo_periodo} encontrado no período selecionado.")
        return

    ranking = df_sel.groupby('Estação').agg(
        **{'Casos Registrados': ('Estação', 'size'),
           'Em Aberto': ('Categoria', lambda x: (x == 'Aberto').sum()),
           'Resolvidos': ('Categoria', lambda x: (x == 'Resolvido').sum()),
           'Último Registro': ('Data', 'max')}
    ).reset_index().sort_values('Casos Registrados', ascending=False)
    ranking['Último Registro'] = ranking['Último Registro'].dt.strftime('%d/%m/%Y')

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Chamados no período", len(df_sel))
    k2.metric("Estações afetadas", ranking['Estação'].nunique())
    k3.metric("Reincidentes (2+ chamados)", int((ranking['Casos Registrados'] >= 2).sum()))
    k4.metric("Em aberto", int((df_sel['Categoria'] == 'Aberto').sum()))

    colA, colB = st.columns([1, 2])
    with colA:
        st.write(f"**Ranking (Desde {data_corte.strftime('%d/%m/%Y')} até {data_base.strftime('%d/%m/%Y')})**")
        st.dataframe(ranking, width='stretch', hide_index=True)
    with colB:
        top = ranking.head(30)
        fig = px.bar(top, x='Estação', y='Casos Registrados', title=f'Top 30 estações ({periodo})',
                     color='Casos Registrados', color_continuous_scale=escala)
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, width='stretch')

    tend = df_sel[df_sel['Categoria'] == 'Aberto'].groupby('Mês Abertura').size().reset_index(name='Chamados')
    fig_t = px.bar(tend, x='Mês Abertura', y='Chamados', text='Chamados', title='Chamados abertos por mês',
                   color_discrete_sequence=[CATEGORIA_CORES['Aberto']])
    fig_t.update_xaxes(dtick='M1', tickformat='%m/%Y')
    st.plotly_chart(fig_t, width='stretch')

    with st.expander("🔎 Ver lista completa de chamados neste período"):
        cols = [c for c in colunas_detalhe if c in df_sel.columns]
        st.dataframe(df_sel[cols], width='stretch')


# ---------------------------------------------------------------------------
# Assistente de consulta por estação
# ---------------------------------------------------------------------------
COLUNAS_CONSULTA = ['ID Sopro', 'Data', 'Tipo', 'Descrição', 'Empresa', 'Status', 'Dias em aberto']
MAX_ESTACOES_RESPOSTA = 10

MENSAGEM_AJUDA = (
    "Digite o código de uma ou mais estações (ex.: **RSBGE02**) ou o ID de um chamado (ex.: **SPR00072339**). "
    "Também dá para digitar o início do código (ex.: **RSBGE**) para ver todas as estações daquele grupo."
)


def identificar_estacoes(pergunta, estacoes):
    """Retorna (estações encontradas, sugestões) a partir do texto digitado pelo técnico."""
    mapa = {str(e).upper(): e for e in estacoes}
    tokens = [t.upper() for t in re.findall(r'[A-Za-z0-9]+', pergunta) if len(t) >= 4]

    encontradas = [mapa[t] for t in tokens if t in mapa]
    if encontradas:
        return list(dict.fromkeys(encontradas)), []

    # Início do código: "RSBGE" -> RSBGE02, RSBGE05...
    por_prefixo = [e for chave, e in sorted(mapa.items()) for t in tokens if chave.startswith(t)]
    if por_prefixo:
        return list(dict.fromkeys(por_prefixo)), []

    sugestoes = []
    for t in tokens:
        sugestoes += difflib.get_close_matches(t, list(mapa), n=5, cutoff=0.75)
    return [], [mapa[s] for s in dict.fromkeys(sugestoes)]


def tabela_chamados(chamados, data_ref):
    t = chamados.sort_values('Data', ascending=False).copy()
    t['Dias em aberto'] = (data_ref - t['Data']).dt.days
    t['Data'] = t['Data'].dt.strftime('%d/%m/%Y')
    return t[[c for c in COLUNAS_CONSULTA if c in t.columns]].reset_index(drop=True)


def responder_consulta(pergunta, df, data_ref):
    """Monta a resposta do assistente como uma lista de blocos {'texto': ..., 'tabela': DataFrame opcional}."""
    blocos = []

    ids = [i.upper() for i in re.findall(r'SPR\d+', pergunta, flags=re.IGNORECASE)]
    if ids and 'ID Sopro' in df.columns:
        achados = df[df['ID Sopro'].astype(str).str.upper().isin(ids)]
        if achados.empty:
            blocos.append({'texto': f"Não encontrei o chamado **{', '.join(ids)}** nesta planilha."})
        for _, c in achados.iterrows():
            fechamento = c.get('Data Fechamento')
            fech_txt = f" · fechado em {fechamento.strftime('%d/%m/%Y')}" if pd.notna(fechamento) else ""
            blocos.append({'texto': f"**{c['ID Sopro']}** · estação **{c['Estação']}** · {c.get('Tipo', '')} — "
                                    f"{c.get('Descrição', '')} · aberto em {c['Data'].strftime('%d/%m/%Y')} · "
                                    f"status **{c['Status']}** · {c.get('Empresa', '')}{fech_txt}"})
        if not re.sub(r'SPR\d+', '', pergunta, flags=re.IGNORECASE).strip():
            return blocos

    estacoes, sugestoes = identificar_estacoes(pergunta, df['Estação'].dropna().unique())
    if not estacoes:
        if blocos:
            return blocos
        texto = "Não reconheci nenhuma estação na sua mensagem."
        if sugestoes:
            texto += " Você quis dizer: " + ", ".join(f"**{s}**" for s in sugestoes) + "?"
        return [{'texto': texto + "\n\n" + MENSAGEM_AJUDA}]

    if len(estacoes) > MAX_ESTACOES_RESPOSTA:
        abertos = df[df['Estação'].isin(estacoes) & (df['Categoria'] == 'Aberto')]
        resumo = abertos.groupby('Estação').agg(**{
            'Chamados em aberto': ('Estação', 'size'),
            'Mais antigo': ('Data', 'min'),
            'Tipos': ('Tipo', lambda x: ', '.join(sorted(x.unique()))),
        }).reset_index().sort_values('Chamados em aberto', ascending=False)
        resumo['Mais antigo'] = resumo['Mais antigo'].dt.strftime('%d/%m/%Y')
        blocos.append({'texto': f"Encontrei **{len(estacoes)}** estações com esse início de código; "
                                f"**{len(resumo)}** delas têm chamado em aberto. Digite o código completo "
                                f"para ver os detalhes de uma estação.",
                       'tabela': resumo if not resumo.empty else None})
        return blocos

    for est in estacoes:
        chamados = df[df['Estação'] == est]
        abertos = chamados[chamados['Categoria'] == 'Aberto']
        if not abertos.empty:
            tipos = ', '.join(sorted(abertos['Tipo'].unique())) if 'Tipo' in abertos else ''
            blocos.append({'texto': f"🔴 A estação **{est}** tem **{len(abertos)}** chamado(s) em aberto "
                                    f"({tipos}):",
                           'tabela': tabela_chamados(abertos, data_ref)})
        else:
            texto = f"🟢 A estação **{est}** não tem chamado em aberto."
            if not chamados.empty:
                ultimo = chamados.sort_values('Data').iloc[-1]
                texto += (f" Último chamado: **{ultimo['ID Sopro']}** ({ultimo.get('Tipo', '')} — "
                          f"{ultimo.get('Descrição', '')}), aberto em {ultimo['Data'].strftime('%d/%m/%Y')}, "
                          f"status {ultimo['Status']}.")
            blocos.append({'texto': texto})
    return blocos


@st.fragment
def aba_consulta(df, data_ref):
    # Fragmento: cada pergunta atualiza só o chat, sem recalcular os gráficos das outras abas
    st.subheader("💬 Assistente de Consulta por Estação")
    st.write("Pergunte se uma estação já tem chamado aberto antes de abrir um novo. A consulta usa a planilha "
             "inteira, sem os filtros da barra lateral.")

    if 'chat_consulta' not in st.session_state:
        st.session_state.chat_consulta = [{'papel': 'assistant', 'blocos': [{'texto': "Olá! " + MENSAGEM_AJUDA}]}]

    historico = st.container()
    if st.button("🗑️ Limpar conversa"):
        del st.session_state.chat_consulta
        st.rerun(scope='fragment')

    pergunta = st.chat_input("Digite a estação ou o ID do chamado (ex.: RSBGE02)")
    if pergunta:
        st.session_state.chat_consulta.append({'papel': 'user', 'blocos': [{'texto': pergunta}]})
        st.session_state.chat_consulta.append({'papel': 'assistant',
                                               'blocos': responder_consulta(pergunta, df, data_ref)})

    with historico:
        for msg in st.session_state.chat_consulta:
            with st.chat_message(msg['papel']):
                for bloco in msg['blocos']:
                    st.markdown(bloco['texto'])
                    if bloco.get('tabela') is not None:
                        st.dataframe(bloco['tabela'], width='stretch', hide_index=True)


# ---------------------------------------------------------------------------
# Seletor de mês/ano em rolagem (wheel picker)
# ---------------------------------------------------------------------------
MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho",
         "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]

CSS = """
.wp { font-family: var(--st-font, sans-serif); color: var(--st-text-color, #31333F); }
.wp-titulo { font-size: 14px; margin-bottom: 4px; }
.wp-corpo { position: relative; display: flex; gap: 4px; height: 180px; border-radius: 12px;
  background: var(--st-secondary-background-color, #f0f2f6); overflow: hidden; user-select: none; }
.wp-faixa { position: absolute; left: 6px; right: 6px; top: 72px; height: 36px; border-radius: 8px;
  background: var(--st-background-color, #fff); opacity: .9; pointer-events: none; }
.wp-col { position: relative; flex: 1; height: 180px; overflow-y: scroll; scroll-snap-type: y mandatory;
  scrollbar-width: none; -webkit-mask-image: linear-gradient(transparent, #000 35%, #000 65%, transparent);
  mask-image: linear-gradient(transparent, #000 35%, #000 65%, transparent); }
.wp-col::-webkit-scrollbar { display: none; }
.wp-item { height: 36px; line-height: 36px; text-align: center; scroll-snap-align: center;
  font-size: 16px; opacity: .45; cursor: pointer; transition: opacity .15s, font-size .15s; }
.wp-item.sel { opacity: 1; font-size: 18px; font-weight: 600; }
.wp-pad { height: 72px; }
"""

JS = """
export default function(component) {
  const { data, parentElement, setStateValue } = component;
  const ALT = 36;
  let raiz = parentElement.querySelector('.wp');
  if (!raiz) {
    raiz = document.createElement('div');
    raiz.className = 'wp';
    parentElement.appendChild(raiz);
  }
  const assinatura = JSON.stringify([data.titulo, data.meses, data.anos]);
  const valorInicial = data.valor;
  if (raiz.dataset.assinatura === assinatura && raiz.dataset.valorAtual === valorInicial) {
    return;  // a roda já mostra esse valor: mantém a posição atual
  }
  raiz.dataset.assinatura = assinatura;
  raiz.dataset.valorAtual = valorInicial;
  raiz.innerHTML = '';

  const titulo = document.createElement('div');
  titulo.className = 'wp-titulo';
  titulo.textContent = data.titulo;
  raiz.appendChild(titulo);

  const corpo = document.createElement('div');
  corpo.className = 'wp-corpo';
  corpo.innerHTML = '<div class="wp-faixa"></div>';
  raiz.appendChild(corpo);

  let [anoSel, mesSel] = valorInicial.split('-').map(Number);
  let ultimoEnviado = valorInicial;

  function enviar() {
    const v = anoSel + '-' + String(mesSel).padStart(2, '0');
    if (v !== ultimoEnviado) { ultimoEnviado = v; raiz.dataset.valorAtual = v; setStateValue('valor', v); }
  }

  function criarColuna(rotulos, indiceInicial, aoMudar) {
    const col = document.createElement('div');
    col.className = 'wp-col';
    col.tabIndex = 0;
    col.innerHTML = '<div class="wp-pad"></div>' +
      rotulos.map((r, i) => '<div class="wp-item" data-i="' + i + '">' + r + '</div>').join('') +
      '<div class="wp-pad"></div>';
    corpo.appendChild(col);
    const itens = col.querySelectorAll('.wp-item');
    let atual = indiceInicial;
    const marcar = (i) => itens.forEach((el, k) => el.classList.toggle('sel', k === i));
    marcar(atual);
    requestAnimationFrame(() => { col.scrollTop = atual * ALT; });
    let timer = null;
    col.addEventListener('scroll', () => {
      const i = Math.max(0, Math.min(rotulos.length - 1, Math.round(col.scrollTop / ALT)));
      marcar(i);
      clearTimeout(timer);
      timer = setTimeout(() => { if (i !== atual) { atual = i; aoMudar(i); } }, 250);
    });
    itens.forEach((el) => el.addEventListener('click', () => {
      col.scrollTo({ top: Number(el.dataset.i) * ALT, behavior: 'smooth' });
    }));
    col.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        const i = Math.max(0, Math.min(rotulos.length - 1, atual + (e.key === 'ArrowDown' ? 1 : -1)));
        col.scrollTo({ top: i * ALT, behavior: 'smooth' });
      }
    });
  }

  criarColuna(data.meses, mesSel - 1, (i) => { mesSel = i + 1; enviar(); });
  criarColuna(data.anos.map(String), Math.max(0, data.anos.indexOf(anoSel)), (i) => { anoSel = data.anos[i]; enviar(); });
}
"""

_roda_mes_ano = st.components.v2.component("seletor_mes_ano", css=CSS, js=JS)


def seletor_mes_ano(titulo, key, valor_inicial, ano_min, ano_max):
    """Duas rodas (mês e ano) no estilo do seletor do celular. Retorna o primeiro dia do mês escolhido."""
    padrao = valor_inicial.strftime('%Y-%m')
    # Estado já atualizado pelo navegador nesta execução (evita devolver à roda um valor antigo)
    estado = st.session_state.get(key)
    atual = (estado.get('valor') if estado else None) or padrao
    res = _roda_mes_ano(key=key, data={'titulo': titulo, 'meses': MESES,
                                       'anos': list(range(ano_min, ano_max + 1)), 'valor': atual},
                        default={'valor': padrao}, on_valor_change=lambda: None)
    return pd.Timestamp((res.valor or padrao) + '-01')


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------
st.title("📊 Afunilador de Chamados e Alertas")
st.write("Converta o arquivo extraído do portal para .xlsx e suba aqui")

arquivo_upload = st.file_uploader("Suba o arquivo Excel aqui", type=["xlsx", "xls"])

if arquivo_upload:
    df = carregar_dados(arquivo_upload.getvalue())
    data_ref = df['Data'].max() if 'Data' in df.columns else pd.Timestamp.today()

    st.sidebar.header("🎯 Funil de Filtros")

    periodo_ini = periodo_fim = None
    if 'Data' in df.columns and df['Data'].notna().any():
        mes_min = df['Data'].min().to_period('M').to_timestamp()
        mes_max = df['Data'].max().to_period('M').to_timestamp()
        versao = st.session_state.get('periodo_versao', 0)

        st.sidebar.markdown("**0. Período de abertura:**")
        with st.sidebar:
            periodo_ini = seletor_mes_ano("De", f'periodo_ini_{versao}', mes_min, mes_min.year, mes_max.year)
            periodo_fim = seletor_mes_ano("Até", f'periodo_fim_{versao}', mes_max, mes_min.year, mes_max.year)
        if st.sidebar.button("↺ Todo o período"):
            st.session_state['periodo_versao'] = versao + 1
            st.rerun()

        if periodo_ini > periodo_fim:
            periodo_ini, periodo_fim = periodo_fim, periodo_ini
            st.sidebar.caption("O mês inicial era posterior ao final; os dois foram invertidos.")
        periodo_ini, periodo_fim = max(periodo_ini, mes_min), min(periodo_fim, mes_max)
        st.sidebar.caption(f"Considerando de {periodo_ini.strftime('%m/%Y')} a {periodo_fim.strftime('%m/%Y')}.")

        df_f0 = df[(df['Data'] >= periodo_ini) & (df['Data'] < periodo_fim + pd.DateOffset(months=1))]
    else:
        df_f0 = df

    empresas_disponiveis = sorted(df_f0['Empresa'].dropna().astype(str).unique().tolist())
    empresa_selecionada = st.sidebar.selectbox("1. Empresa Prestadora:", options=["TODAS"] + empresas_disponiveis)

    if empresa_selecionada != "TODAS":
        df_f1 = df_f0[df_f0['Empresa'] == empresa_selecionada]
    else:
        df_f1 = df_f0

    status_selecionado = st.sidebar.selectbox(
        "2. Status do Chamado:",
        options=["TODOS"] + sorted(df_f1['Status'].unique().tolist())
    )

    df_f2 = filtrar_status(df_f1, status_selecionado)

    sla_disponiveis = sorted(df_f2['SLA'].dropna().astype(str).unique().tolist())
    sla_selecionado = st.sidebar.selectbox("3. SLA:", options=["TODOS"] + sla_disponiveis)

    if sla_selecionado != "TODOS":
        df_f3 = df_f2[df_f2['SLA'].astype(str) == sla_selecionado]
    else:
        df_f3 = df_f2

    if 'Tipo' in df_f3.columns:
        tipos_disponiveis = sorted(df['Tipo'].unique().tolist())
        tipos_selecionados = st.sidebar.multiselect("4. Tipo de Chamado:", options=tipos_disponiveis,
                                                    default=tipos_disponiveis)
        df_f4 = df_f3[df_f3['Tipo'].isin(tipos_selecionados)]
    else:
        df_f4 = df_f3

    estacoes_disponiveis = sorted(df_f4['Estação'].dropna().astype(str).unique().tolist())
    estacoes_selecionadas = st.sidebar.multiselect(
        "5. Estações (Sites):", options=estacoes_disponiveis,
        help="Deixe vazio para considerar todas as estações."
    )

    df_final = df_f4[df_f4['Estação'].isin(estacoes_selecionadas)] if estacoes_selecionadas else df_f4

    if df_final.empty:
        st.warning("Nenhum chamado encontrado com os filtros selecionados.")
        st.stop()

    st.caption(f"Data de referência da planilha (último chamado aberto): **{data_ref.strftime('%d/%m/%Y')}** · "
               f"{len(df_final)} chamados no filtro atual")

    aba1, aba2, aba3, aba4, aba5, aba6, aba7 = st.tabs([
        "📊 Visão Geral", "📈 Histórico de Resolução", "🏢 Eficiência por Empresa",
        "🧾 Resumo Executivo", "🐦 Ninhos na EV", "🌿 Zeladorias", "💬 Consultar Estação"
    ])

    # ----------------------------------------------------------------- Visão Geral
    with aba1:
        st.subheader("📈 Resumo do Filtro Atual")
        total_chamados = len(df_final)
        cat = df_final['Categoria'].value_counts()
        qtd_aberto = int(cat.get('Aberto', 0))
        qtd_resolvido = int(cat.get('Resolvido', 0))
        qtd_cancelado = int(cat.get('Cancelado', 0))

        qtd_no_prazo = int((df_final['SLA'] == 'NO PRAZO').sum())
        qtd_sla_medido = int(df_final['SLA'].isin(['NO PRAZO', 'FORA DO PRAZO']).sum())
        mediana_res = df_final['Dias p/ Resolver'].median() if 'Dias p/ Resolver' in df_final else None

        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("Total de Chamados", total_chamados)
        m2.metric("% Em Aberto", f"{pct(qtd_aberto, total_chamados):.1f}%", f"{qtd_aberto} chamados", delta_color="off")
        m3.metric("% Resolvidos", f"{pct(qtd_resolvido, total_chamados):.1f}%", f"{qtd_resolvido} chamados", delta_color="off")
        m4.metric("% Cancelados", f"{pct(qtd_cancelado, total_chamados):.1f}%", f"{qtd_cancelado} chamados", delta_color="off")
        m5.metric("% SLA (No Prazo)", f"{pct(qtd_no_prazo, qtd_sla_medido):.1f}%",
                  help="Chamados com bolinha verde ÷ chamados com bolinha verde ou vermelha.")
        m6.metric("Mediana p/ Resolver", f"{mediana_res:.0f} dias" if pd.notna(mediana_res) else "—",
                  help="Dias entre a abertura e a data de fechamento dos chamados concluídos/fechados.")

        st.divider()

        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Status dos chamados")
            df_status = df_final.groupby(['Categoria', 'Status']).size().reset_index(name='Quantidade')
            fig_sun = px.sunburst(df_status, path=['Categoria', 'Status'], values='Quantidade',
                                  color='Categoria', color_discrete_map=CATEGORIA_CORES)
            fig_sun.update_traces(textinfo='label+percent root')
            st.plotly_chart(fig_sun, width='stretch')
        with c2:
            st.subheader("SLA por tipo de chamado")
            if 'Tipo' in df_final.columns:
                df_sla_tipo = df_final.groupby(['Tipo', 'SLA']).size().reset_index(name='Quantidade')
                fig_sla = px.bar(df_sla_tipo, x='Tipo', y='Quantidade', color='SLA', barmode='stack',
                                 color_discrete_map=SLA_CORES)
                st.plotly_chart(fig_sla, width='stretch')

        st.subheader("📊 Distribuição de Status por Estação (Top 30 em volume)")
        top_estacoes = df_final['Estação'].value_counts().head(30).index
        df_grafico = df_final[df_final['Estação'].isin(top_estacoes)].groupby(['Estação', 'Categoria']).size().reset_index(name='Quantidade')
        df_grafico['Total_Estacao'] = df_grafico.groupby('Estação')['Quantidade'].transform('sum')
        df_grafico['Porcentagem'] = (df_grafico['Quantidade'] / df_grafico['Total_Estacao'] * 100).round(1)

        fig = px.bar(
            df_grafico, x='Estação', y='Quantidade', color='Categoria',
            text='Porcentagem', barmode='stack', color_discrete_map=CATEGORIA_CORES,
            category_orders={'Estação': list(top_estacoes)}
        )
        fig.update_traces(texttemplate='%{text}%', textposition='inside')
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, width='stretch')

        st.divider()

        st.subheader("🏢 Relatório por Estação")
        relatorio_estacao = resumo_por_grupo(df_final, 'Estação', data_ref).sort_values('Total', ascending=False)
        st.dataframe(relatorio_estacao, width='stretch', hide_index=True)

        csv_relatorio = relatorio_estacao.to_csv(index=False, sep=';').encode('utf-8-sig')
        st.download_button(label="📥 Baixar Relatório (CSV)", data=csv_relatorio, file_name='relatorio.csv', mime='text/csv')

        with st.expander("🔎 Ver dados detalhados deste filtro"):
            st.dataframe(df_final, width='stretch')

            csv_detalhado = df_final.to_csv(index=False, sep=';').encode('utf-8-sig')
            st.download_button(
                label=f"Baixar Dados Detalhados - {empresa_selecionada} (CSV)",
                data=csv_detalhado,
                file_name=f'dados_detalhados_{empresa_selecionada}.csv',
                mime='text/csv',
            )

    # ------------------------------------------------------- Histórico de Resolução
    with aba2:
        st.subheader("📈 Histórico de Resolução")
        st.write("Chamados abertos por mês (data de abertura) comparados com os resolvidos e cancelados "
                 "(data de fechamento). O backlog acumulado mostra se a fila está crescendo ou diminuindo.")

        hist = historico_mensal(df_final, periodo_ini, periodo_fim)
        if hist.empty:
            st.info("Sem dados suficientes para montar o histórico.")
        else:
            fig_h = make_subplots(specs=[[{"secondary_y": True}]])
            fig_h.add_bar(x=hist['Mês'], y=hist['Abertos'], name='Abertos', marker_color='#3498DB')
            fig_h.add_bar(x=hist['Mês'], y=hist['Resolvidos'], name='Resolvidos', marker_color=CATEGORIA_CORES['Resolvido'])
            fig_h.add_bar(x=hist['Mês'], y=hist['Cancelados'], name='Cancelados', marker_color=CATEGORIA_CORES['Cancelado'])
            fig_h.add_trace(go.Scatter(x=hist['Mês'], y=hist['Backlog Acumulado'], name='Backlog acumulado',
                                       mode='lines+markers', line=dict(color='#E74C3C', width=3)), secondary_y=True)
            fig_h.update_layout(barmode='group', title='Abertos x Resolvidos x Cancelados por mês',
                                legend=dict(orientation='h', y=-0.2), hovermode='x unified')
            fig_h.update_yaxes(title_text='Chamados no mês', secondary_y=False)
            fig_h.update_yaxes(title_text='Backlog acumulado', secondary_y=True)
            st.plotly_chart(fig_h, width='stretch')

            c1, c2 = st.columns(2)
            with c1:
                fig_tempo = go.Figure()
                fig_tempo.add_trace(go.Scatter(x=hist['Mês'], y=hist['Mediana Dias Resolução'], name='Mediana',
                                               mode='lines+markers', line=dict(color='#8E44AD', width=3)))
                fig_tempo.add_trace(go.Scatter(x=hist['Mês'], y=hist['Média Dias Resolução'], name='Média',
                                               mode='lines', line=dict(color='#8E44AD', dash='dot')))
                fig_tempo.update_layout(title='Tempo de resolução dos chamados fechados no mês (dias)',
                                        hovermode='x unified', legend=dict(orientation='h', y=-0.2))
                st.plotly_chart(fig_tempo, width='stretch')
            with c2:
                fig_taxa = px.line(hist, x='Mês', y='% Resolvidos s/ Abertos', markers=True,
                                   title='Capacidade de vazão: resolvidos ÷ abertos no mês (%)')
                fig_taxa.add_hline(y=100, line_dash='dash', line_color='gray',
                                   annotation_text='100% = fila estável')
                st.plotly_chart(fig_taxa, width='stretch')

            # Tempo de resolução por coorte de abertura e tipo
            res = df_final[df_final['Categoria'] == 'Resolvido'].dropna(subset=['Dias p/ Resolver'])
            if not res.empty and 'Tipo' in res.columns:
                c3, c4 = st.columns(2)
                with c3:
                    fig_box = px.box(res, x='Tipo', y='Dias p/ Resolver', color='Tipo', points=False,
                                     title='Distribuição do tempo de resolução por tipo')
                    fig_box.update_layout(showlegend=False)
                    st.plotly_chart(fig_box, width='stretch')
                with c4:
                    curva = res['Dias p/ Resolver'].sort_values().reset_index(drop=True)
                    acum = pd.DataFrame({'Dias': curva, '% Resolvidos': (curva.index + 1) / len(curva) * 100})
                    fig_cdf = px.line(acum, x='Dias', y='% Resolvidos',
                                      title='Curva acumulada: % dos chamados resolvidos até X dias')
                    for d in [30, 60, 90]:
                        p = (curva <= d).mean() * 100
                        fig_cdf.add_vline(x=d, line_dash='dot', line_color='gray',
                                          annotation_text=f'{d}d: {p:.0f}%')
                    st.plotly_chart(fig_cdf, width='stretch')

            # Backlog atual por idade
            abertos = df_final[df_final['Categoria'] == 'Aberto'].copy()
            if not abertos.empty:
                st.subheader("⏳ Idade do backlog em aberto")
                abertos['Idade (dias)'] = (data_ref - abertos['Data']).dt.days
                abertos['Faixa'] = abertos['Idade (dias)'].apply(faixa_aging)
                aging = abertos.groupby(['Faixa', 'Tipo' if 'Tipo' in abertos else 'Status']).size().reset_index(name='Chamados')
                fig_aging = px.bar(aging, x='Faixa', y='Chamados', color=aging.columns[1], barmode='stack',
                                   category_orders={'Faixa': FAIXAS_AGING},
                                   title=f'{len(abertos)} chamados em aberto por tempo desde a abertura')
                st.plotly_chart(fig_aging, width='stretch')

            with st.expander("📋 Tabela do histórico mensal"):
                hist_tab = hist.copy()
                hist_tab['Mês'] = hist_tab['Mês'].dt.strftime('%m/%Y')
                st.dataframe(hist_tab, width='stretch', hide_index=True)

    # ------------------------------------------------------- Eficiência por Empresa
    with aba3:
        st.subheader("🏢 Eficiência das Empresas Prestadoras")
        st.write("O **índice de eficiência** (0-100) combina: 40% taxa de resolução (resolvidos ÷ chamados não "
                 "cancelados), 30% SLA no prazo e 30% rapidez (mediana de dias para resolver, comparada entre as "
                 "empresas exibidas).")

        # A comparação usa período, tipo e estações da barra lateral, mas ignora os filtros de
        # empresa, status e SLA, que distorceriam as taxas de resolução e de SLA.
        df_emp_base = df_f0
        if 'Tipo' in df_emp_base.columns:
            df_emp_base = df_emp_base[df_emp_base['Tipo'].isin(tipos_selecionados)]
        if estacoes_selecionadas:
            df_emp_base = df_emp_base[df_emp_base['Estação'].isin(estacoes_selecionadas)]
        if empresa_selecionada != "TODAS" or status_selecionado != "TODOS" or sla_selecionado != "TODOS":
            st.info("Os filtros de empresa, status e SLA não se aplicam a esta aba, para que todas as "
                    "empresas sejam comparadas sobre a mesma base de chamados.")

        vol_min = st.slider("Volume mínimo de chamados para entrar no ranking:", 1, 200, 20, step=1)
        emp = resumo_por_grupo(df_emp_base, 'Empresa', data_ref)
        emp = emp[emp['Total'] >= vol_min].copy()

        if emp.empty:
            st.info("Nenhuma empresa atinge o volume mínimo selecionado.")
        else:
            emp['Índice de Eficiência'] = calcular_score(emp)
            emp = emp.sort_values('Índice de Eficiência', ascending=False)
            ordem = emp['Empresa'].tolist()

            melhor, pior = emp.iloc[0], emp.iloc[-1]
            k1, k2, k3 = st.columns(3)
            k1.metric("Empresas no ranking", len(emp))
            k2.metric("🥇 Mais eficiente", melhor['Empresa'], f"índice {melhor['Índice de Eficiência']:.1f}",
                      delta_color="off")
            k3.metric("⚠️ Menos eficiente", pior['Empresa'], f"índice {pior['Índice de Eficiência']:.1f}",
                      delta_color="off")

            fig_score = px.bar(emp, x='Índice de Eficiência', y='Empresa', orientation='h',
                               color='Índice de Eficiência', color_continuous_scale='RdYlGn', range_color=[0, 100],
                               text='Índice de Eficiência', title='Índice de eficiência por empresa')
            fig_score.update_layout(yaxis={'categoryorder': 'total ascending'}, height=max(350, 35 * len(emp)))
            st.plotly_chart(fig_score, width='stretch')

            c1, c2 = st.columns(2)
            with c1:
                comp = emp.melt(id_vars='Empresa', value_vars=['% Resolução', '% SLA No Prazo'],
                                var_name='Indicador', value_name='%')
                fig_comp = px.bar(comp, x='Empresa', y='%', color='Indicador', barmode='group',
                                  category_orders={'Empresa': ordem}, title='Taxa de resolução x SLA no prazo',
                                  color_discrete_map={'% Resolução': '#2ECC71', '% SLA No Prazo': '#3498DB'})
                fig_comp.update_layout(xaxis_tickangle=-45, legend=dict(orientation='h', y=-0.45))
                st.plotly_chart(fig_comp, width='stretch')
            with c2:
                fig_disp = px.scatter(emp, x='Mediana Dias Resolução', y='% Resolução', size='Total',
                                      color='% SLA No Prazo', color_continuous_scale='RdYlGn', hover_name='Empresa',
                                      title='Volume x rapidez x resolução (bolha = volume)', size_max=60)
                st.plotly_chart(fig_disp, width='stretch')

            c3, c4 = st.columns(2)
            with c3:
                mix = df_emp_base[df_emp_base['Empresa'].isin(ordem)].groupby(['Empresa', 'Categoria']).size().reset_index(name='Chamados')
                fig_mix = px.bar(mix, x='Empresa', y='Chamados', color='Categoria', barmode='stack',
                                 color_discrete_map=CATEGORIA_CORES, category_orders={'Empresa': ordem},
                                 title='Volume e situação dos chamados por empresa')
                fig_mix.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_mix, width='stretch')
            with c4:
                res_emp = df_emp_base[(df_emp_base['Empresa'].isin(ordem)) & df_emp_base['Dias p/ Resolver'].notna()]
                fig_box_emp = px.box(res_emp, x='Empresa', y='Dias p/ Resolver', points=False,
                                     category_orders={'Empresa': ordem},
                                     title='Distribuição dos dias para resolver por empresa')
                fig_box_emp.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_box_emp, width='stretch')

            st.subheader("📉 Evolução mensal por empresa")
            empresas_hist = st.multiselect("Empresas para comparar:", options=ordem, default=ordem[:5])
            if empresas_hist:
                base_h = df_emp_base[df_emp_base['Empresa'].isin(empresas_hist)]
                res_h = base_h[base_h['Categoria'] == 'Resolvido']
                if periodo_ini is not None:
                    res_h = res_h[res_h['Mês Fechamento'].between(periodo_ini, periodo_fim)]
                c5, c6 = st.columns(2)
                with c5:
                    serie = res_h.groupby(['Mês Fechamento', 'Empresa']).size().reset_index(name='Resolvidos')
                    fig_r = px.line(serie, x='Mês Fechamento', y='Resolvidos', color='Empresa', markers=True,
                                    title='Chamados resolvidos por mês')
                    fig_r.update_layout(legend=dict(orientation='h', y=-0.3))
                    st.plotly_chart(fig_r, width='stretch')
                with c6:
                    serie_t = res_h.groupby(['Mês Fechamento', 'Empresa'])['Dias p/ Resolver'].median().reset_index()
                    fig_t = px.line(serie_t, x='Mês Fechamento', y='Dias p/ Resolver', color='Empresa', markers=True,
                                    title='Mediana de dias para resolver (por mês de fechamento)')
                    fig_t.update_layout(legend=dict(orientation='h', y=-0.3))
                    st.plotly_chart(fig_t, width='stretch')

                if 'Tipo' in base_h.columns:
                    calor = resumo_por_grupo(base_h, ['Empresa', 'Tipo'], data_ref)
                    calor = calor.pivot(index='Empresa', columns='Tipo', values='% Resolução')
                    fig_heat = px.imshow(calor, text_auto='.0f', color_continuous_scale='RdYlGn', zmin=0, zmax=100,
                                         aspect='auto', title='% de resolução por empresa e tipo de chamado')
                    st.plotly_chart(fig_heat, width='stretch')

            st.subheader("📋 Tabela de indicadores por empresa")
            st.dataframe(emp, width='stretch', hide_index=True)
            st.download_button("📥 Baixar indicadores por empresa (CSV)",
                               data=emp.to_csv(index=False, sep=';').encode('utf-8-sig'),
                               file_name='eficiencia_empresas.csv', mime='text/csv')

    # ------------------------------------------------------------ Resumo Executivo
    with aba4:
        st.subheader("🧾 Resumo Executivo")

        hist = historico_mensal(df_final, periodo_ini, periodo_fim)
        emp_all = resumo_por_grupo(df_final, 'Empresa', data_ref)
        emp_all = emp_all[emp_all['Total'] >= 20].copy()
        if not emp_all.empty:
            emp_all['Índice de Eficiência'] = calcular_score(emp_all)
        tipo_res = resumo_por_grupo(df_final, 'Tipo', data_ref) if 'Tipo' in df_final.columns else pd.DataFrame()
        est_res = resumo_por_grupo(df_final, 'Estação', data_ref)

        abertos = df_final[df_final['Categoria'] == 'Aberto'].copy()
        abertos['Idade (dias)'] = (data_ref - abertos['Data']).dt.days

        # Comparação dos últimos 3 meses fechados com os 3 anteriores
        destaques = []
        if len(hist) >= 6:
            ult, ant = hist.tail(3), hist.iloc[-6:-3]
            ab_u, ab_a = ult['Abertos'].sum(), ant['Abertos'].sum()
            rs_u, rs_a = ult['Resolvidos'].sum(), ant['Resolvidos'].sum()
            var_ab = pct(ab_u - ab_a, ab_a)
            var_rs = pct(rs_u - rs_a, rs_a)
            destaques.append(f"Nos últimos 3 meses foram abertos **{ab_u}** chamados "
                             f"({var_ab:+.1f}% vs. trimestre anterior) e resolvidos **{rs_u}** ({var_rs:+.1f}%).")
            med_u, med_a = ult['Mediana Dias Resolução'].mean(), ant['Mediana Dias Resolução'].mean()
            if pd.notna(med_u) and pd.notna(med_a):
                tend = "melhorou" if med_u < med_a else "piorou"
                destaques.append(f"O tempo mediano de resolução **{tend}**: {med_a:.0f} → {med_u:.0f} dias.")
            saldo = ult['Saldo do Mês'].sum()
            destaques.append(f"O backlog {'cresceu' if saldo > 0 else 'diminuiu'} **{abs(saldo)}** chamados no último trimestre.")

        if not abertos.empty:
            mais90 = int((abertos['Idade (dias)'] > 90).sum())
            destaques.append(f"Existem **{len(abertos)}** chamados em aberto; **{mais90}** "
                             f"({pct(mais90, len(abertos)):.0f}%) têm mais de 90 dias.")
        if not emp_all.empty and len(emp_all) > 1:
            e = emp_all.sort_values('Índice de Eficiência', ascending=False)
            destaques.append(f"Empresa mais eficiente (≥20 chamados): **{e.iloc[0]['Empresa']}** "
                             f"(índice {e.iloc[0]['Índice de Eficiência']:.1f}); menos eficiente: "
                             f"**{e.iloc[-1]['Empresa']}** (índice {e.iloc[-1]['Índice de Eficiência']:.1f}).")
            maior_backlog = emp_all.sort_values('Abertos', ascending=False).iloc[0]
            destaques.append(f"Maior backlog: **{maior_backlog['Empresa']}** com {maior_backlog['Abertos']} "
                             f"chamados em aberto ({maior_backlog['Backlog +90 dias']} com +90 dias).")
        if not tipo_res.empty and tipo_res['Mediana Dias Resolução'].notna().any():
            lento = tipo_res.sort_values('Mediana Dias Resolução', ascending=False).iloc[0]
            destaques.append(f"Tipo mais lento para resolver: **{lento['Tipo']}** "
                             f"(mediana {lento['Mediana Dias Resolução']:.0f} dias).")
        reinc = est_res[est_res['Total'] >= 5]
        destaques.append(f"**{len(reinc)}** estações tiveram 5 ou mais chamados no período (reincidência).")
        if 'Dias p/ Acionar' in df_final.columns and df_final['Dias p/ Acionar'].notna().any():
            destaques.append(f"Tempo mediano entre abertura e acionamento da empresa: "
                             f"**{df_final['Dias p/ Acionar'].median():.0f} dia(s)**.")

        st.markdown("\n".join(f"- {d}" for d in destaques))
        st.divider()

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**Resumo por tipo de chamado**")
            if not tipo_res.empty:
                st.dataframe(tipo_res[['Tipo', 'Total', 'Resolvidos', 'Abertos', 'Cancelados', '% Resolução',
                                       '% SLA No Prazo', 'Mediana Dias Resolução']],
                             width='stretch', hide_index=True)
        with c2:
            st.markdown("**Top 15 problemas mais recorrentes**")
            if 'Descrição' in df_final.columns:
                desc = resumo_por_grupo(df_final, 'Descrição', data_ref).sort_values('Total', ascending=False).head(15)
                st.dataframe(desc[['Descrição', 'Total', 'Abertos', '% Resolução', 'Mediana Dias Resolução']],
                             width='stretch', hide_index=True)

        if 'Descrição' in df_final.columns and 'Tipo' in df_final.columns:
            tree = df_final.groupby(['Tipo', 'Descrição']).size().reset_index(name='Chamados')
            fig_tree = px.treemap(tree, path=[px.Constant('Todos'), 'Tipo', 'Descrição'], values='Chamados',
                                  title='Mapa de problemas: tipo → descrição')
            st.plotly_chart(fig_tree, width='stretch')

        c3, c4 = st.columns(2)
        with c3:
            st.markdown("**Estações mais reincidentes**")
            top_est = est_res.sort_values('Total', ascending=False).head(20)
            st.dataframe(top_est[['Estação', 'Total', 'Abertos', 'Resolvidos', '% Resolução',
                                  'Idade Média Backlog (dias)']], width='stretch', hide_index=True)
        with c4:
            st.markdown("**Chamados abertos há mais tempo**")
            cols_crit = [c for c in ['ID Sopro', 'Data', 'Estação', 'Tipo', 'Descrição', 'Empresa', 'Status',
                                     'Idade (dias)'] if c in abertos.columns]
            criticos = abertos.sort_values('Idade (dias)', ascending=False)[cols_crit].head(20)
            st.dataframe(criticos, width='stretch', hide_index=True)

        st.divider()
        resumo_kpis = pd.DataFrame({
            'Indicador': ['Período inicial', 'Período final', 'Total de chamados', 'Em aberto', 'Resolvidos',
                          'Cancelados', '% SLA no prazo', 'Mediana dias p/ resolver', 'Backlog com +90 dias'],
            'Valor': [df_final['Data'].min().strftime('%d/%m/%Y'), df_final['Data'].max().strftime('%d/%m/%Y'),
                      len(df_final), len(abertos), int((df_final['Categoria'] == 'Resolvido').sum()),
                      int((df_final['Categoria'] == 'Cancelado').sum()),
                      f"{pct((df_final['SLA'] == 'NO PRAZO').sum(), df_final['SLA'].isin(['NO PRAZO', 'FORA DO PRAZO']).sum()):.1f}%",
                      df_final['Dias p/ Resolver'].median(), int((abertos['Idade (dias)'] > 90).sum())],
        })
        resumo_kpis['Valor'] = resumo_kpis['Valor'].astype(str)
        destaques_df = pd.DataFrame({'Destaques': [d.replace('**', '') for d in destaques]})
        hist_xlsx = hist.copy()
        if not hist_xlsx.empty:
            hist_xlsx['Mês'] = hist_xlsx['Mês'].dt.strftime('%m/%Y')

        excel = gerar_excel({
            'Resumo': resumo_kpis,
            'Destaques': destaques_df,
            'Histórico Mensal': hist_xlsx,
            'Por Empresa': emp_all.sort_values('Índice de Eficiência', ascending=False) if not emp_all.empty else emp_all,
            'Por Tipo': tipo_res,
            'Por Estação': est_res.sort_values('Total', ascending=False),
            'Backlog em Aberto': abertos.sort_values('Idade (dias)', ascending=False),
            'Dados Filtrados': df_final,
        })
        st.download_button("📥 Baixar Relatório Completo (Excel)", data=excel,
                           file_name=f"relatorio_sopro_{data_ref.strftime('%Y%m%d')}.xlsx",
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # ------------------------------------------------------------- Ninhos / Zeladoria
    colunas_det = ['ID Sopro', 'Data', 'Estação', 'Tipo', 'Descrição', 'Empresa', 'Status', 'Data Fechamento']

    with aba5:
        st.subheader("🚨 Sites com Maior Incidência de Ninhos")
        st.write("Sites que mais geram chamado de ninhos")
        ranking_ocorrencias(
            df_f1, lambda d: d['Descrição'].astype(str).str.contains('Insetos Ou Ninhos', case=False, na=False),
            "Ninhos", 'Reds', colunas_det, {"Últimos 3 meses": 3, "Últimos 6 meses": 6, "Últimos 12 meses": 12},
            'periodo_ninhos'
        )

    with aba6:
        st.subheader("🌿 Sites com Maior Incidência de Problemas de Zeladoria")
        st.write("Monitore as estações que mais demandaram serviços de zeladoria no histórico recente.")
        ranking_ocorrencias(
            df_f1, lambda d: d['Tipo'].astype(str).str.contains('Zeladoria', case=False, na=False),
            "Zeladoria", 'Greens', colunas_det, {"Últimos 3 meses": 3, "Últimos 6 meses": 6, "Últimos 12 meses": 12},
            'periodo_zel'
        )

    with aba7:
        aba_consulta(df, data_ref)
