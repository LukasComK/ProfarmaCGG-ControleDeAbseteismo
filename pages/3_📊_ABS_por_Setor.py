import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="ABS por Setor", layout="wide")

st.title("📊 ABS por Setor")
st.markdown("""
Esta página processa o arquivo Excel (XLSX) para gerar uma tabela de **Justificativas** (ABS) por dia.

**Regras de Processamento:**
1. **Filtro de Cargo** (Coluna I): Considera apenas `AUXILIAR DEPOSITO I`, `AUXILIAR DEPOSITO II`, `AUXILIAR DEPOSITO III`.
   - **Colunas Extras**: Cargo (Coluna I) e Gestor (via cruzamento com CSV).
   - **Matriz**:
     - **Linhas**: Nomes, Cargos e Gestores
     - **Colunas**: Datas (Coluna J)
     - **Célula**: Justificativa (Coluna P)
""")

uploaded_file = st.file_uploader("1. Carregue o arquivo Excel de Absenteísmo", type=["xlsx"])
uploaded_csv_gestores = st.file_uploader("2. Carregue o arquivo CSV de Gestores (Base Ativos)", type=["csv"])

if uploaded_file is not None and uploaded_csv_gestores is not None:
    # 1. Leitura do Arquivo Excel (Absenteísmo)
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel: {e}")
        st.stop()

    # 2. Leitura do Arquivo CSV (Gestores)
    try:
        # Tenta ; ISO-8859-1 que é comum
        df_gestores = pd.read_csv(uploaded_csv_gestores, sep=';', encoding='latin-1', engine='python')
    except:
        uploaded_csv_gestores.seek(0)
        try:
            # Tenta , UTF-8
            df_gestores = pd.read_csv(uploaded_csv_gestores, sep=',', encoding='utf-8', engine='python')
        except:
            uploaded_csv_gestores.seek(0)
            # Tenta tab
            df_gestores = pd.read_csv(uploaded_csv_gestores, sep='\t', encoding='utf-8', engine='python')

    # Verifica se CSV tem colunas D (idx 3) e Z (idx 25)
    # Mas como as letras podem não bater com indices se houver leitura errada, vamos tentar pegar pelo indice mesmo se tiver colunas suficientes.
    # D -> 3 (0,1,2,3)
    # Z -> 25
    
    col_nome_gestor_idx = 3 # Col D
    col_gestor_idx = 25     # Col Z
    
    mapa_gestores = {}
    
    if len(df_gestores.columns) > 25:
        # Cria dicionário Nome -> Gestor
        # Normaliza nomes para cruzar
        col_nome_csv = df_gestores.columns[col_nome_gestor_idx]
        col_gestor_csv = df_gestores.columns[col_gestor_idx]
        
        # Limpa e cria dicionario
        temp_df = df_gestores[[col_nome_csv, col_gestor_csv]].dropna().astype(str)
        temp_df['Chave'] = temp_df[col_nome_csv].str.strip().str.upper()
        temp_df['Valor'] = temp_df[col_gestor_csv].str.strip().str.upper()
        
        # Remove duplicados mantendo o ultimo ou primeiro? Vamos assumir que nome é unico
        temp_df = temp_df.drop_duplicates(subset=['Chave'])
        
        mapa_gestores = dict(zip(temp_df['Chave'], temp_df['Valor']))
        st.success(f"Arquivo de Gestores carregado! {len(mapa_gestores)} mapeamentos encontrados.")
    else:
        st.warning(f"Arquivo CSV de gestores parece não ter colunas suficientes (esperado > 25, encontrado {len(df_gestores.columns)}). Verifique separador.")

    st.success(f"Arquivo Absenteísmo carregado! {len(df)} linhas encontradas.")

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
        # Vamos garantir que pegue variações com espaços extras
        mask = df['Cargo_Norm'].isin([c.strip().upper() for c in cargos_permitidos])
        
        # Se nao achar nada exato, tenta contains?
        if mask.sum() == 0:
             st.warning("Nenhum cargo exato encontrado. Tentando busca parcial 'AUXILIAR DEPOSITO'...")
             mask = df['Cargo_Norm'].str.contains("AUXILIAR DEPOSITO", case=False, na=False)
        
        df_filtered = df[mask].copy()
        
        st.info(f"Linhas após filtro de cargos: {len(df_filtered)}")
        
        if len(df_filtered) > 0:
            # 3. Pivot Table (Crosstab) incluindo Cargo e Gestor
            
            # Adiciona Coluna Gestor no DataFrame Filtrado

            def get_gestor(nome_val):
                n = str(nome_val).strip().upper()
                return mapa_gestores.get(n, "NÃO ENCONTRADO")
            
            df_filtered['GESTOR'] = df_filtered[col_nome].apply(get_gestor)
            
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
            # Index: Nome (col H), Cargo (col I), Gestor (novo)
            # Columns: Data (col J formatada)
            # Values: Justificativa (col P)
            
            # Função de agregação: Se tiver 2 justificativas no mesmo dia, concatena.
            agg_func = lambda x: " | ".join([str(v) for v in x if pd.notna(v) and str(v).strip() != ''])
            
            # Renomeia colunas para ficar bonito no índice da Pivot
            df_filtered = df_filtered.rename(columns={
                col_nome: 'NOME',
                col_cargo: 'CARGO'
            })
            
            pivot_df = df_filtered.pivot_table(
                index=['NOME', 'CARGO', 'GESTOR'], 
                columns='Data_Str', 
                values=col_justif, 
                aggfunc=agg_func
            ).fillna('') # Preenche vazios com string vazia
            
            # Reset index para NOME, CARGO, GESTOR virarem colunas normais e facilitar export
            pivot_df = pivot_df.reset_index()
            
            # Ordenar colunas de data (o pivot pode bagunçar se for string)
            # Se conseguimos converter para datetime, podemos reordenar as colunas
            if 'Data_Det' in df_filtered.columns:
                 datas_unicas = df_filtered[['Data_Det', 'Data_Str']].drop_duplicates().sort_values('Data_Det')
                 cols_datas = datas_unicas['Data_Str'].tolist()
                 # Garante que só usa colunas que existem no pivot
                 cols_datas_finais = [c for c in cols_datas if c in pivot_df.columns]
                 
                 # Colunas fixas + colunas de datas
                 cols_finais = ['NOME', 'CARGO', 'GESTOR'] + cols_datas_finais
                 pivot_df = pivot_df[cols_finais]

            st.write("### Resultado da Matriz de Justificativas")
            st.dataframe(pivot_df, use_container_width=True)
            
            # 4. Exportação Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                pivot_df.to_excel(writer, sheet_name='Justificativas', index=False)
                
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
                    worksheet.write(0, col_num, value, header_fmt) # sem +1 pois index=False
                
                # Ajusta largura
                worksheet.set_column(0, 0, 40) # Nome
                worksheet.set_column(1, 1, 25) # Cargo
                worksheet.set_column(2, 2, 25) # Gestor
                worksheet.set_column(3, len(pivot_df.columns), 15) # Datas
                
            st.download_button(
                label="📥 Baixar Planilha Excel",
                data=buffer.getvalue(),
                file_name="Check_Justificativas_Aux_Deposito.xlsx",
                mime="application/vnd.ms-excel"
            )
            
        else:
            st.warning("Nenhum dado encontrado para os cargos filtrados.")
