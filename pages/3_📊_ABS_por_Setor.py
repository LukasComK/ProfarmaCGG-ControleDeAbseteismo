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
        # Tenta ler ignorando a primeira linha (que costuma ser título "Colaboradores")
        df_gestores = pd.read_csv(uploaded_csv_gestores, sep=';', encoding='latin-1', skiprows=1, engine='python')
        
        # Se mesmo assim tiver poucas colunas, tenta ler normal (caso o arquivo não tenha o título)
        if len(df_gestores.columns) < 5:
             uploaded_csv_gestores.seek(0)
             df_gestores = pd.read_csv(uploaded_csv_gestores, sep=';', encoding='latin-1', engine='python')

    except:
        uploaded_csv_gestores.seek(0)
        try:
             # Tenta skiprows=1 com ; e utf-8
             df_gestores = pd.read_csv(uploaded_csv_gestores, sep=';', encoding='utf-8', skiprows=1, engine='python')
        except:
             try:
                # Tenta normal , utf-8
                uploaded_csv_gestores.seek(0)
                df_gestores = pd.read_csv(uploaded_csv_gestores, sep=',', encoding='utf-8', engine='python')
             except:
                 pass

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
            
            # 1. Obtém o GESTOR direto
            df_filtered['GESTOR'] = df_filtered[col_nome].apply(get_gestor)
            
            # 2. Obtém o SUPERVISOR (Gestor do Gestor)
            # Reaproveita a mesma lógica: procura quem é o gestor do nome que está na coluna 'GESTOR'
            df_filtered['SUPERVISOR'] = df_filtered['GESTOR'].apply(lambda x: get_gestor(x) if x != "NÃO ENCONTRADO" else "NÃO ENCONTRADO")
            
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
            # Index: Nome (col H), Cargo (col I), Gestor (novo), Supervisor (novo)
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
                index=['NOME', 'CARGO', 'GESTOR', 'SUPERVISOR'], 
                columns='Data_Str', 
                values=col_justif, 
                aggfunc=agg_func
            ).fillna('') # Preenche vazios com string vazia
            
            # Reset index para NOME, CARGO, GESTOR, SUPERVISOR virarem colunas normais e facilitar export
            pivot_df = pivot_df.reset_index()
            
            # Ordenar colunas de data (o pivot pode bagunçar se for string)
            # Se conseguimos converter para datetime, podemos reordenar as colunas
            if 'Data_Det' in df_filtered.columns:
                 datas_unicas = df_filtered[['Data_Det', 'Data_Str']].drop_duplicates().sort_values('Data_Det')
                 cols_datas = datas_unicas['Data_Str'].tolist()
                 # Garante que só usa colunas que existem no pivot
                 cols_datas_finais = [c for c in cols_datas if c in pivot_df.columns]
                 
                 # Colunas fixas + colunas de datas
                 cols_finais = ['NOME', 'CARGO', 'GESTOR', 'SUPERVISOR'] + cols_datas_finais
                 pivot_df = pivot_df[cols_finais]

            st.write("### Resultado da Matriz de Justificativas")
            st.dataframe(pivot_df, use_container_width=True)
            
            # 4. Exportação Excel
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                # Aba 1: Matriz Detalhada
                pivot_df.to_excel(writer, sheet_name='Justificativas', index=False)
                
                # Ajustes visuais Aba Justificativas
                workbook = writer.book
                worksheet = writer.sheets['Justificativas']
                
                # Formatos
                header_fmt = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1
                })
                
                # Escreve cabeçalho com formato
                for col_num, value in enumerate(pivot_df.columns.values):
                    worksheet.write(0, col_num, value, header_fmt)
                
                # Ajusta largura
                worksheet.set_column(0, 0, 40) # Nome
                worksheet.set_column(1, 1, 25) # Cargo
                worksheet.set_column(2, 2, 25) # Gestor
                worksheet.set_column(3, 3, 25) # Supervisor
                worksheet.set_column(4, len(pivot_df.columns), 15) # Datas
                
                # Aba 2: Resumo Hierárquico (Tipo Tabela Dinâmica)
                # Vamos preparar os dados para esse relatório
                # Precisamos de NOME, CARGO, GESTOR, SUPERVISOR, DATA, JUSTIFICATIVA
                # Vamos explodir a pivot ou usar o df_filtered original?
                # O df_filtered original tem tudo o que precisamos.
                
                # Filtra apenas quem tem justificativa preenchida (falta/atraso/etc)
                # A coluna P (Justificativa) não pode ser vazia/nula
                df_resumo = df_filtered[df_filtered[col_justif].notna() & (df_filtered[col_justif].astype(str).str.strip() != '')].copy()
                
                if not df_resumo.empty:
                    worksheet_resumo = workbook.add_worksheet('Resumo_Faltas')
                    
                    # Formatos para o Resumo
                    fmt_supervisor = workbook.add_format({'bold': True, 'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1})
                    fmt_gestor = workbook.add_format({'bold': True, 'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'border': 1, 'indent': 1})
                    fmt_funcionario = workbook.add_format({'bold': True, 'indent': 2, 'border': 1})
                    fmt_data = workbook.add_format({'indent': 4, 'font_color': '#555555'})
                    fmt_total = workbook.add_format({'bold': True, 'align': 'center', 'border': 1})
                    
                    # Cabeçalho
                    worksheet_resumo.write(0, 0, "HIERARQUIA / DATA", header_fmt)
                    worksheet_resumo.write(0, 1, "QTD FALTAS", header_fmt)
                    worksheet_resumo.set_column(0, 0, 60)
                    worksheet_resumo.set_column(1, 1, 15)
                    
                    row = 1
                    
                    # Agrupa por Supervisor
                    grupos_supervisor = df_resumo.groupby('SUPERVISOR')
                    
                    for supervisor, df_sup in grupos_supervisor:
                        qtd_sup = len(df_sup)
                        worksheet_resumo.write(row, 0, f"📂 {supervisor}", fmt_supervisor)
                        worksheet_resumo.write(row, 1, qtd_sup, fmt_supervisor)
                        # Agrupamento Excel (Level 1)
                        # worksheet_resumo.set_row(row, None, None, {'level': 1}) 
                        row += 1
                        
                        # Agrupa por Gestor dentro do Supervisor
                        grupos_gestor = df_sup.groupby('GESTOR')
                        for gestor, df_ges in grupos_gestor:
                            qtd_ges = len(df_ges)
                            worksheet_resumo.write(row, 0, f"👤 {gestor}", fmt_gestor)
                            worksheet_resumo.write(row, 1, qtd_ges, fmt_gestor)
                            # worksheet_resumo.set_row(row, None, None, {'level': 2})
                            row += 1
                            
                            # Agrupa por Funcionário dentro do Gestor
                            grupos_func = df_ges.groupby('NOME')
                            for func, df_func in grupos_func:
                                qtd_func = len(df_func)
                                worksheet_resumo.write(row, 0, f"🔹 {func}", fmt_funcionario)
                                worksheet_resumo.write(row, 1, qtd_func, fmt_total)
                                # worksheet_resumo.set_row(row, None, None, {'level': 3})
                                row += 1
                                
                                # Lista as datas e justificativas
                                for _, reg in df_func.iterrows():
                                    data_txt = str(reg['Data_Str'])
                                    just_txt = str(reg[col_justif])
                                    texto = f"{data_txt} - {just_txt}"
                                    worksheet_resumo.write(row, 0, texto, fmt_data)
                                    worksheet_resumo.write(row, 1, 1, fmt_data) # Conta 1
                                    # worksheet_resumo.set_row(row, None, None, {'level': 4})
                                    row += 1
            
            st.download_button(
                label="📥 Baixar Planilha Excel",
                data=buffer.getvalue(),
                file_name="Check_Justificativas_Aux_Deposito.xlsx",
                mime="application/vnd.ms-excel"
            )
            
        else:
            st.warning("Nenhum dado encontrado para os cargos filtrados.")
