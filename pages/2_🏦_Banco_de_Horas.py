"""
Página: Banco de Horas
Descrição: Visualização e gestão do banco de horas dos colaboradores
Recebe: Arquivo XLSX com banco de horas + CSV base de colaboradores
"""

import streamlit as st
import pandas as pd
import datetime
import io
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Banco de Horas", layout="wide")

st.title("🏦 Banco de Horas")
st.write("Visualização e gestão do banco de horas dos colaboradores")

st.divider()

# Upload dos arquivos
col1, col2 = st.columns(2)

with col1:
    st.subheader("📥 Arquivo de Banco de Horas")
    file_banco_horas = st.file_uploader(
        "Selecione o arquivo XLSX com banco de horas",
        type=["xlsx"],
        key="banco_horas"
    )

with col2:
    st.subheader("👥 Base de Colaboradores")
    file_colaboradores = st.file_uploader(
        "Selecione o CSV com base de colaboradores",
        type=["csv", "xlsx"],
        key="colaboradores_bh"
    )

st.divider()

# Processa os arquivos se ambos forem carregados
if file_banco_horas and file_colaboradores:
    try:
        # Carrega banco de horas
        df_banco_horas = pd.read_excel(file_banco_horas)
        
        # Carrega base de colaboradores
        if file_colaboradores.name.endswith('.csv'):
            df_colaboradores = pd.read_csv(file_colaboradores)
        else:
            df_colaboradores = pd.read_excel(file_colaboradores)
        
        st.success("✅ Arquivos carregados com sucesso!")
        
        st.divider()
        
        # Mostra as abas de informações
        tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "📋 Detalhes", "📈 Análises"])
        
        with tab1:
            st.subheader("Dashboard de Banco de Horas")
            st.info("Dashboard será implementado em breve...")
            
            # Preview dos dados carregados
            with st.expander("Ver dados do Banco de Horas (primeiras linhas)"):
                st.dataframe(df_banco_horas.head(10), use_container_width=True)
            
            with st.expander("Ver base de Colaboradores (primeiras linhas)"):
                st.dataframe(df_colaboradores.head(10), use_container_width=True)
        
        with tab2:
            st.subheader("Detalhes por Colaborador")
            st.info("Visualização detalhada será implementada em breve...")
        
        with tab3:
            st.subheader("Análises e Relatórios")
            st.info("Gráficos e análises serão implementados em breve...")
    
    except Exception as e:
        st.error(f"❌ Erro ao processar arquivos: {str(e)}")
else:
    st.warning("⚠️ Por favor, carregue ambos os arquivos para continuar.")
    
    st.info("""
    ### Como usar esta página:
    
    1. 📤 **Upload do XLSX de Banco de Horas**: Arquivo contendo dados de horas trabalhadas, extras, etc.
    2. 📤 **Upload do CSV de Colaboradores**: Base com informações dos funcionários
    3. 📊 **Visualize** os dados consolidados nas abas
    4. 📈 **Analise** o saldo de horas por colaborador, gestor, etc.
    """)

