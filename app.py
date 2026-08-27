import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl

# Configuração da página
st.set_page_config(page_title="Painel de Controle - Estações", layout="wide")

st.title("📊 Afunilador de Chamados e Alertas")
st.write("Converta o arquivo extraído do portal para .xlxs e suba aqui")

arquivo_upload = st.file_uploader("Suba o arquivo Excel aqui", type=["xlsx", "xls"])

if arquivo_upload:
    # 1. O pandas lê apenas o texto (onde aparece o ●)
    df = pd.read_excel(arquivo_upload, header=1)
    
    # 2. Reseta o arquivo na memória para podermos ler a formatação
    arquivo_upload.seek(0)
    
    # 3. Usa o openpyxl para "olhar" as cores da planilha
    wb = openpyxl.load_workbook(arquivo_upload, data_only=True)
    ws = wb.active
    
    # Descobre qual é a coluna do SLA (na linha 2 do Excel, já que o header=1)
    sla_col_idx = None
    for idx, cell in enumerate(ws[2], 1):
        if cell.value and str(cell.value).strip().upper() == 'SLA':
            sla_col_idx = idx
            break
    
    # Analisa a cor de cada célula da coluna SLA
    status_sla_cores = []
    for row_idx in range(3, ws.max_row + 1):
        status = "SEM SLA"
        if sla_col_idx:
            cell = ws.cell(row=row_idx, column=sla_col_idx)
            
            # Tenta pegar a cor da fonte (texto) ou do fundo da célula
            cor_hex = None
            if cell.font and cell.font.color and cell.font.color.type == 'rgb':
                cor_hex = str(cell.font.color.rgb).upper()
            elif cell.fill and cell.fill.fgColor and cell.fill.fgColor.type == 'rgb':
                cor_hex = str(cell.fill.fgColor.rgb).upper()
            
            if cor_hex:
                # Se a cor em HEX (RGB) tiver código de Verde (usado pelo Excel)
                if "00B050" in cor_hex or "00FF00" in cor_hex or "2ECC71" in cor_hex or "008000" in cor_hex:
                    status = "NO PRAZO"
                # Se a cor tiver código de Vermelho
                elif "FF0000" in cor_hex or "C00000" in cor_hex or "FF4500" in cor_hex or "E74C3C" in cor_hex:
                    status = "FORA DO PRAZO"
        
        status_sla_cores.append(status)
        
    # Como o Excel pode ler linhas vazias invisíveis a mais no final, ajustamos o tamanho
    status_sla_cores = status_sla_cores[:len(df)]
    if len(status_sla_cores) < len(df):
        status_sla_cores += ["SEM SLA"] * (len(df) - len(status_sla_cores))
        
    # Substitui os "●" pelo status real lido pelas cores!
    df['SLA'] = status_sla_cores
    
    # Limpeza básica e continua o app normal...
    df = df.dropna(subset=['Estação'])
    df['Status'] = df['Status'].astype(str).str.strip()
    df['Empresa'] = df['Empresa'].astype(str).str.strip()
    
    # Convertendo a coluna 'Data' para o formato de data do Python
    if 'Data' in df.columns:
        df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
   
    st.sidebar.header("🎯 Funil de Filtros")
    
    # 1. FILTRO: Empresa Prestadora
    empresas_disponiveis = sorted(df['Empresa'].dropna().astype(str).unique().tolist())
    empresa_selecionada = st.sidebar.selectbox("1. Empresa Prestadora:", options=["TODAS"] + empresas_disponiveis)
    
    if empresa_selecionada != "TODAS":
        df_f1 = df[df['Empresa'] == empresa_selecionada]
    else:
        df_f1 = df

    # 2. FILTRO: Status do Chamado
    status_selecionado = st.sidebar.selectbox("2. Status do Chamado:", options=["TODOS", "EM ANDAMENTO", "CONCLUÍDOS / FECHADOS"])
    
    if status_selecionado == "EM ANDAMENTO":
        df_f2 = df_f1[df_f1['Status'].str.lower() == 'em andamento']
    elif status_selecionado == "CONCLUÍDOS / FECHADOS":
        df_f2 = df_f1[df_f1['Status'].str.lower().str.contains('concl|fech')]
    else:
        df_f2 = df_f1 

    # 3. FILTRO: SLA
    sla_disponiveis = sorted(df_f2['SLA'].dropna().astype(str).unique().tolist())
    sla_selecionado = st.sidebar.selectbox("3. SLA:", options=["TODOS"] + sla_disponiveis)
    
    if sla_selecionado != "TODOS":
        df_f3 = df_f2[df_f2['SLA'].astype(str) == sla_selecionado]
    else:
        df_f3 = df_f2

    # 4. FILTRO: Estações (Sites)
    estacoes_disponiveis = sorted(df_f3['Estação'].dropna().astype(str).unique().tolist())
    estacoes_selecionadas = st.sidebar.multiselect("4. Estações (Sites):", options=estacoes_disponiveis, default=estacoes_disponiveis)
    
    if estacoes_selecionadas:
        df_final = df_f3[df_f3['Estação'].isin(estacoes_selecionadas)]
        
        # --- CRIANDO AS ABAS NA TELA ---
        aba1, aba2, aba3 = st.tabs(["📊 Visão Geral e Relatórios", "🐦 Ninhos na EV", "🌿 Zeladorias"])
        
        # ==========================================
        # ABA 1: VISÃO GERAL (O QUE JÁ TÍNHAMOS)
        # ==========================================
        with aba1:
            st.subheader("📈 Resumo do Filtro Atual")
            total_chamados = len(df_final)
            
            # Contagens
            status_counts = df_final['Status'].value_counts()
            qtd_andamento = status_counts.get("Em andamento", 0)
            qtd_concluido = sum(v for k, v in status_counts.items() if "concl" in k.lower() or "fech" in k.lower())
            
            # Cálculo de Porcentagens
            pct_andamento = (qtd_andamento / total_chamados * 100) if total_chamados > 0 else 0
            pct_concluido = (qtd_concluido / total_chamados * 100) if total_chamados > 0 else 0
            
            # Lógica do SLA (Aqui ele busca palavras que indicam que o SLA está bom)
            # Se na sua planilha estiver escrito "Dentro do Prazo", ele vai achar pela palavra "prazo".
            qtd_no_prazo = df_final['SLA'].astype(str).str.lower().str.contains('prazo|cumprido|dentro|ok').sum()
            pct_sla = (qtd_no_prazo / total_chamados * 100) if total_chamados > 0 else 0
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total de Chamados", total_chamados)
            m2.metric("% Em Andamento", f"{pct_andamento:.1f}%")
            m3.metric("% Concluídos", f"{pct_concluido:.1f}%")
            m4.metric("% SLA (No Prazo)", f"{pct_sla:.1f}%") # Exibe o SLA
            
            st.divider()
            
            # Gráfico de Barras com Porcentagem
            st.subheader("📊 Distribuição de Status por Estação")
            if not df_final.empty:
                df_grafico = df_final.groupby(['Estação', 'Status']).size().reset_index(name='Quantidade')
                df_grafico['Total_Estacao'] = df_grafico.groupby('Estação')['Quantidade'].transform('sum')
                df_grafico['Porcentagem'] = (df_grafico['Quantidade'] / df_grafico['Total_Estacao'] * 100).round(1)
                
                fig = px.bar(
                    df_grafico, x='Estação', y='Quantidade', color='Status',
                    text='Porcentagem', barmode='stack',
                    color_discrete_map={'Em andamento': '#FFA500', 'Concluído': '#2ECC71', 'Fechado': '#27AE60'}
                )
                fig.update_traces(texttemplate='%{text}%', textposition='inside')
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig, use_container_width=True)
                
            st.divider()
            
            # Tabela de Relatório
            st.subheader(f"🏢 Relatório por Estação")
            relatorio_estacao = df_final.groupby('Estação').agg(
                Total_Chamados=('ID Sopro', 'count'),
                Em_Andamento=('Status', lambda x: (x == 'Em andamento').sum()),
                Concluidos_Fechados=('Status', lambda x: x.str.lower().str.contains('concl|fech').sum())
            ).reset_index()
            
            st.dataframe(relatorio_estacao, use_container_width=True)
            
            csv_relatorio = relatorio_estacao.to_csv(index=False, sep=';').encode('utf-8-sig')
            st.download_button(label="📥 Baixar Relatório (CSV)", data=csv_relatorio, file_name='relatorio.csv', mime='text/csv')

            # Permite visualizar a planilha aberta e filtrada no final se necessário
        with st.expander("🔎 Ver dados detalhados deste filtro"):
            st.dataframe(df_final, use_container_width=True)

            # ---NOVO: BOTÃO DE DOWNLOAD DOS DADOS BRUTOS---
            csv_detalhado = df_final.to_csv(index=False, sep=';').encode('utf-8-sig')
            st.download_button(
                label=f"Baixar Dados Detalhados - {empresa_selecionada} (CSV)",
                data=csv_detalhado,
                file_name=f'dados_detalhados_{empresa_selecionada}.csv',
                mime='text/csv',
            )

        # ==========================================
        # ABA 2: ALERTAS DE NINHOS
        # ==========================================
        with aba2:
            st.subheader("🚨 Sites com Maior Incidência de Ninhos")
            st.write("Sites que mais geram chamado de ninhos")
            
            # Só roda se a coluna de Data existir e foi convertida corretamente
            if 'Data' in df.columns and pd.api.types.is_datetime64_any_dtype(df['Data']):
                
                # Seleção do Período na tela
                periodo = st.radio("Selecione o período de análise:", ["Últimos 3 meses", "Últimos 6 meses"], horizontal=True)
                
                # Pega a data mais recente da planilha como base (ideal caso suba planilhas antigas)
                data_base = df['Data'].max()
                
                # Subtrai 3 ou 6 meses da data base usando o Pandas DateOffset
                if periodo == "Últimos 3 meses":
                    data_corte = data_base - pd.DateOffset(months=3)
                else:
                    data_corte = data_base - pd.DateOffset(months=6)
                
                # Aplica o filtro de Data e busca pela palavra-chave na coluna Descrição
                df_insetos = df[(df['Data'] >= data_corte) & (df['Descrição'].astype(str).str.contains('Insetos Ou Ninhos', case=False, na=False))]
                
                if not df_insetos.empty:
                    # Agrupa para ver quais estações tem mais chamados desse tipo
                    ranking_insetos = df_insetos.groupby('Estação').size().reset_index(name='Casos Registrados')
                    ranking_insetos = ranking_insetos.sort_values(by='Casos Registrados', ascending=False)
                    
                    colA, colB = st.columns([1, 2])
                    
                    with colA:
                        st.write(f"**Ranking (Desde {data_corte.strftime('%d/%m/%Y')} até {data_base.strftime('%d/%m/%Y')})**")
                        st.dataframe(ranking_insetos, use_container_width=True, hide_index=True)
                    
                    with colB:
                        # Gráfico visual do Ranking
                        fig_insetos = px.bar(
                            ranking_insetos, x='Estação', y='Casos Registrados',
                            title=f'Volume de Ocorrências ({periodo})',
                            color='Casos Registrados', color_continuous_scale='Reds' # Deixa as barras mais vermelhas quanto mais casos
                        )
                        st.plotly_chart(fig_insetos, use_container_width=True)
                        
                    # Detalhamento
                    with st.expander("🔎 Ver lista completa de chamados de Insetos/Ninhos neste período"):
                        st.dataframe(df_insetos[['ID Sopro', 'Data', 'Estação', 'Descrição', 'Empresa', 'Status']], use_container_width=True)
                else:
                    st.success(f"Ótima notícia! Nenhum chamado de Ninhos encontrado no período de {periodo}.")
            else:
                st.error("Não foi possível ler as datas da planilha para gerar este relatório.")

        # ==========================================
        # ABA 3: ALERTAS DE ZELADORIA
        # ==========================================
        with aba3:
            st.subheader("🌿 Sites com Maior Incidência de Problemas de Zeladoria")
            st.write("Monitore as estações que mais demandaram serviços de zeladoria no histórico recente.")
            
            # Só roda se a coluna de Data existir
            if 'Data' in df.columns and pd.api.types.is_datetime64_any_dtype(df['Data']):
                
                # Seleção do Período na tela (Agora com a opção de 12 meses!)
                periodo_zel = st.radio(
                    "Selecione o período de análise (Zeladoria):", 
                    ["Últimos 3 meses", "Últimos 6 meses", "Últimos 12 meses"], 
                    horizontal=True
                )
                
                data_base = df['Data'].max()
                
                # Calcula a data de corte baseada na escolha
                if periodo_zel == "Últimos 3 meses":
                    data_corte_zel = data_base - pd.DateOffset(months=3)
                elif periodo_zel == "Últimos 6 meses":
                    data_corte_zel = data_base - pd.DateOffset(months=6)
                else:
                    data_corte_zel = data_base - pd.DateOffset(months=12)
                
                # Filtra pela palavra 'Zeladoria' na coluna 'Tipo' 
                df_zeladoria = df[(df['Data'] >= data_corte_zel) & (df['Tipo'].astype(str).str.contains('Zeladoria', case=False, na=False))]
                
                if not df_zeladoria.empty:
                    # Ranking de estações
                    ranking_zel = df_zeladoria.groupby('Estação').size().reset_index(name='Casos Registrados')
                    ranking_zel = ranking_zel.sort_values(by='Casos Registrados', ascending=False)
                    
                    colA, colB = st.columns([1, 2])
                    
                    with colA:
                        st.write(f"**Ranking ({periodo_zel})**")
                        st.dataframe(ranking_zel, use_container_width=True, hide_index=True)
                    
                    with colB:
                        # Gráfico em tons de verde para combinar com Zeladoria
                        fig_zel = px.bar(
                            ranking_zel, x='Estação', y='Casos Registrados',
                            title=f'Volume de Ocorrências de Zeladoria',
                            color='Casos Registrados', color_continuous_scale='Greens' 
                        )
                        st.plotly_chart(fig_zel, use_container_width=True)
                        
                    # Tabela expandida para detalhes
                    with st.expander("🔎 Ver lista completa de chamados de Zeladoria neste período"):
                        st.dataframe(df_zeladoria[['ID Sopro', 'Data', 'Estação', 'Tipo', 'Descrição', 'Empresa', 'Status']], use_container_width=True)
                else:
                    st.success(f"Nenhum chamado de Zeladoria encontrado no período selecionado.")
            else:
                st.error("Não foi possível ler as datas da planilha para gerar este relatório.")
            
    else:
        st.warning("Por favor, selecione ao menos uma Estação na barra lateral.")
