import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Painel de Controle - Site", layout="wide")

st.title("📊 Afunilador de Chamados - Sharings & Site")
st.write("Suba o arquivo extraído do portal para cruzar dados de Sharing e site.")

# Componente para subir o arquivo (Aceita .xlsx e .xls) ''
arquivo_upload = st.file_uploader("Suba o arquivo Excel aqui", type=["xlsx", "xls"])

if arquivo_upload:
    # O header=1 ignora a primeira linha vazia do portal
    df = pd.read_excel(arquivo_upload, header=1)
    
    # Limpeza básica de linhas sem Site ou sem status
    df = df.dropna(subset=['Estação']) 
    df['Status'] = df['Status'].astype(str).str.strip()
    df['Sharing'] = df['Empresa'].astype(str).str.strip()

    # --- BARRA LATERAL (FUNIL DE FILTROS) ---
    st.sidebar.header("🎯 Funil de Filtros")
    
    # 1. NOVO FILTRO: Filtrar por Sharing Prestadora
    empresas_disponiveis = sorted(df['Empresa']dropna().astype(str).unique().tolist())
    empresa_selecionada = st.sidebar.selectbox(
        "1. Selecione a Sharing:",
        options=["TODAS"] + empresas_disponiveis
    )
    
    # Aplicando o primeiro nível do funil (Sharing)
    if empresa_selecionada != "TODAS":
        df_filtrado_emp = df[df['Sharing'] == empresa_selecionada]
    else:
        df_filtrado_emp = df

    # Filtro de Status
    status_selecionado = st.sidebar.selectbox(
        "2. Filtre pelo Status do Chamado:",
        options=["TODOS", "EM ANDAMENTO", "CONCLUÍDOS"]
    )

    if status_selecionado == "EM ANDAMENTO":
        df_filtrado_status = df_filtrado_emp[df_filtrado_emp['Status'].str.lower() == 'em andamento']
    elif status_selecionado == "CONCLUÍDOS":
        df_filtrado_status = df_filtrado_emp[df_filtrado_emp['Status'].str.lower() == 'concluído']
    else:
        df_filtrado_status = df_filtrado_emp

    # 2. FILTRO: Filtrar pelas Estações daquela Sharing
    estacoes_disponiveis = sorted(df_filtrado_status['Estação'].unique().tolist())
    estacoes_selecionadas = st.sidebar.multiselect(
        "2. Filtre pelas Estações (Sites):", 
        options=estacoes_disponiveis,
        default=estacoes_disponiveis
    )
    
    if estacoes_selecionadas:
        # Dataframe final com todos os filtros aplicados
        df_final = df_filtrado_status[df_filtrado_status['Estação'].isin(estacoes_selecionadas)]
        
        # --- SEÇÃO 1: METRICAS GERAIS GLOBAIS ---
        st.subheader("📈 Resumo do Filtro Atual")
        total_chamados = len(df_final)
        
        # Contagem de Status
        status_counts = df_final['Status'].value_counts()
        qtd_andamento = status_counts.get("Em andamento", 0)
        
        # Identifica dinamicamente variações de concluído/fechado
        qtd_concluido = sum(v for k, v in status_counts.items() if "concl" in k.lower() or "fech" in k.lower())
        qtd_outros = total_chamados - (qtd_andamento + qtd_concluido)
        
        pct_andamento = (qtd_andamento / total_chamados * 100) if total_chamados > 0 else 0
        pct_concluido = (qtd_concluido / total_chamados * 100) if total_chamados > 0 else 0
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Total de Chamados no Filtro", total_chamados)
        m2.metric("% Em Andamento", f"{pct_andamento:.1f}%")
        m3.metric("% Concluídos", f"{pct_concluido:.1f}%")
        
        st.divider()
        
        # --- SEÇÃO 2: RELATÓRIO DA EMPRESA X ESTAÇÃO ---
        st.subheader(f"🏢 Relatório de Chamados por Site (Sharing: {empresa_selecionada})")
        
        # Criando a tabela agrupada por Site   
        relatorio_estacao = df_final.groupby('Estação').agg(
            Total_Chamados=('ID Sopro', 'count'),
            Em_Andamento=('Status', lambda x: (x == 'Em andamento').sum()),
            Concluidos_Fechados=('Status', lambda x: x.str.lower().str.contains('concl|fech').sum())
        ).reset_index()
        
        st.dataframe(relatorio_estacao, use_container_width=True)

        csv_relatorio = relatorio_estacao.to_csv(index=False, sep=';').encode('utf-8-sig')
        st.download_button(
            label=f"Baixar Relatório da Sharing - {empresa_selecionada} (CSV)",
            data=csv_relatorio,
            file_name=f'relatorio_sharing_{empresa_selecionada}.csv',
            mime='text/csv',
        )
        
        st.divider()
        
        # --- SEÇÃO 3: GRÁFICO DE BARRAS EMPILHADAS COM PORCENTAGEM ---
        st.subheader("📊 Distribuição de Rótulos e Status por Site")
        
        if not df_final.empty:
            # Agrupa os dados para o formato que o gráfico precisa
            df_grafico = df_final.groupby(['Estação', 'Status']).size().reset_index(name='Quantidade')
            
            # Calcula o total da Site   para descobrir a porcentagem individual de cada pedaço da barra
            df_grafico['Total_Estacao'] = df_grafico.groupby('Estação')['Quantidade'].transform('sum')
            df_grafico['Porcentagem'] = (df_grafico['Quantidade'] / df_grafico['Total_Estacao'] * 100).round(1)
            
            # Monta o gráfico utilizando o Plotly
            fig = px.bar(
                df_grafico, 
                x='Estação', 
                y='Quantidade', 
                color='Status',
                title=f'Volume Total por Site e % de Status Status ({empresa_selecionada})',
                text='Porcentagem', # Aplica a porcentagem calculada dentro do bloco da barra
                labels={'Quantidade': 'Total de Chamados', 'Site': 'Site / Site'},
                barmode='stack',
                color_discrete_map={'Em andamento': '#FFA500', 'Concluído': '#2ECC71', 'Fechado': '#27AE60'} # Cores amigáveis
            )
            
            # Formata a exibição do texto para incluir o símbolo de % dentro das barras
            fig.update_traces(texttemplate='%{text}%', textposition='inside')
            fig.update_layout(xaxis_tickangle=-45, legend_title_text='Situação do Chamado')
            
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Nenhum dado encontrado para os filtros selecionados.")
            
        st.divider()
        
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
            
    else:
        st.warning("Por favor, selecione ao menos uma Site na barra lateral para renderizar os dados.")
