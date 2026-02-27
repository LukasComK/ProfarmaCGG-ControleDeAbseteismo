import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ABS por Setor", layout="wide")

st.title("📊 ABS por Setor")
st.markdown("""
Esta página processa o arquivo Excel (XLSX) para gerar uma tabela de **Justificativas** (ABS) por dia.

**Regras de Processamento:**
1. **Filtro de Cargo** (Coluna I): Considera apenas `AUXILIAR DEPOSITO I`, `AUXILIAR DEPOSITO II`, `AUXILIAR DEPOSITO III`.
2. **Matriz**:
   - **Linhas**: Nomes (Coluna H)
   - **Colunas**: Datas (Coluna J)
   - **Célula**: Justificativa (Coluna P)
""")

uploaded_file = st.file_uploader("Carregue o arquivo Excel de Absenteísmo", type=["xlsx"])

if uploaded_file is not None:
    # 1. Leitura do Arquivo
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel: {e}")
        st.stop()

    st.success(f"Arquivo carregado! {len(df)} linhas encontradas.")

    if len(df.columns) < 16:
        st.error("O arquivo parece não ter as colunas H..P necessárias ou o cabeçalho não foi lido corretamente.")
        st.write("Colunas encontradas:", list(df.columns))
    else:
        # Colunas Chave pelo índice (0-based)
        # H -> 7 (Nome)
        # I -> 8 (Cargo)
        # J -> 9 (Data)
        # P -> 15 (Justificativa)
        
        col_nome_idx = 7
        col_cargo_idx = 8
        col_data_idx = 9
        col_justif_idx = 15
        
        col_nome = df.columns[col_nome_idx]
        col_cargo = df.columns[col_cargo_idx]
        col_data = df.columns[col_data_idx]
        col_justif = df.columns[col_justif_idx]
        
        # 2. Filtragem de Cargo
        cargos_permitidos = [
            "AUXILIAR DEPOSITO I",
            "AUXILIAR DEPOSITO II",
            "AUXILIAR DEPOSITO III"
        ]
        
        # Cria coluna normalizada para evitar erros de espaços/case
        df['Cargo_Norm'] = df[col_cargo].astype(str).str.strip().str.upper()
        
        # Filtra
        # Usamos regex para pegar variações ou lista exata? O pedido foi exato.
        # "AUXILIAR DEPOSITO I"
        
        # Vamos garantir que pegue variações com espaços extras
        mask = df['Cargo_Norm'].isin([c.strip().upper() for c in cargos_permitidos])
        
        # Se nao achar nada exato, tenta contains?
        if mask.sum() == 0:
             st.warning("Nenhum cargo exato encontrado. Tentando busca parcial 'AUXILIAR DEPOSITO'...")
             mask = df['Cargo_Norm'].str.contains("AUXILIAR DEPOSITO", case=False, na=False)
        
        df_filtered = df[mask].copy()
        
        st.info(f"Linhas após filtro de cargos: {len(df_filtered)}")
        
        if len(df_filtered) > 0:
            # 3. Pivot Table (Crosstab)
            
            # Formata Data para garantir ordem cronológica nas colunas
            # Tenta converter para datetime
            try:
                df_filtered['Data_Det'] = pd.to_datetime(df_filtered[col_data], dayfirst=True, errors='coerce')
                # Remove NaT se data inválida
                df_filtered = df_filtered.dropna(subset=['Data_Det'])
                # Formata string DD/MM/YYYY
                df_filtered['Data_Str'] = df_filtered['Data_Det'].dt.strftime('%d/%m/%Y')
                # Ordena pelo datetime para pivot respeitar
                df_filtered = df_filtered.sort_values('Data_Det')
            except Exception as e:
                st.warning(f"Erro ao converter datas, usando texto original: {e}")
                df_filtered['Data_Str'] = df_filtered[col_data].astype(str)
            
            # Pivot
            # Index: Nome (col H)
            # Columns: Data (col J formatada)
            # Values: Justificativa (col P)
            
            # Função de agregação: Se tiver 2 justificativas no mesmo dia, concatena.
            agg_func = lambda x: " | ".join([str(v) for v in x if pd.notna(v) and str(v).strip() != ''])
            
            pivot_df = df_filtered.pivot_table(
                index=col_nome, 
                columns='Data_Str', 
                values=col_justif, 
                aggfunc=agg_func
            ).fillna('') # Preenche vazios com string vazia
            
            # Ordenar colunas de data (o pivot pode bagunçar se for string)
            # Se conseguimos converter para datetime, podemos reordenar as colunas
            if 'Data_Det' in df_filtered.columns:
                 datas_unicas = df_filtered[['Data_Det', 'Data_Str']].drop_duplicates().sort_values('Data_Det')
                 cols_ordenadas = datas_unicas['Data_Str'].tolist()
                 # Garante que só usa colunas que existem no pivot (caso algum filtro tenha removido)
                 cols_finais = [c for c in cols_ordenadas if c in pivot_df.columns]
                 pivot_df = pivot_df[cols_finais]

            st.write("### Resultado da Matriz de Justificativas")
            st.dataframe(pivot_df, use_container_width=True)
            
            # 4. Exportação Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                pivot_df.to_excel(writer, sheet_name='Justificativas')
                
                # Ajustes visuais
                workbook = writer.book
                worksheet = writer.sheets['Justificativas']
                
                # Formato Header
                header_fmt = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1
                })
                
                # Escreve cabeçalho com formato
                for col_num, value in enumerate(pivot_df.columns.values):
                    worksheet.write(0, col_num + 1, value, header_fmt)
                worksheet.write(0, 0, "NOME", header_fmt)
                
                # Ajusta largura
                worksheet.set_column(0, 0, 40) # Nome largo
                worksheet.set_column(1, len(pivot_df.columns), 15) # Datas
                
            st.download_button(
                label="📥 Baixar Planilha Excel",
                data=buffer.getvalue(),
                file_name="Check_Justificativas_Aux_Deposito.xlsx",
                mime="application/vnd.ms-excel"
            )
            
        else:
            st.warning("Nenhum dado encontrado para os cargos filtrados.")
