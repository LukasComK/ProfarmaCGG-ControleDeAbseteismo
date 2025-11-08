# Ferramenta de Lançamento de Absenteísmo com Busca LIKE
import streamlit as st
import pandas as pd
from unidecode import unidecode
import io
import datetime
import re
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from difflib import SequenceMatcher

def eh_fim_de_semana(data):
    """Retorna True se é sábado (5) ou domingo (6)"""
    return data.weekday() in [5, 6]

def calcular_similaridade(s1, s2):
    """Calcula similaridade entre duas strings (0 a 1)"""
    return SequenceMatcher(None, s1, s2).ratio()

def limpar_nome(nome):
    if isinstance(nome, str):
        return unidecode(nome).upper().strip()
    return ""

def extrair_dia_do_cabecalho(label_dia, mes, ano):
    if pd.isna(label_dia):
        return None
    
    label_str = str(label_dia).lower()
    
    try:
        data = pd.to_datetime(label_str, dayfirst=True)
        if data.month == mes and data.year == ano:
            return data.date()
    except:
        pass

    mapa_mes_curto = {'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6, 
                      'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12}
    
    dia_num_str = None
    mes_encontrado = None

    for nome_mes, num_mes in mapa_mes_curto.items():
        if nome_mes in label_str:
            if num_mes == mes:
                mes_encontrado = num_mes
                dia_num_str = label_str.split(nome_mes)[0]
                break
    
    if mes_encontrado is None:
        partes = re.split(r'[/.-]', label_str)
        if len(partes) >= 1:
            dia_num_str = partes[0]
        if len(partes) >= 2:
            try:
                if int(partes[1]) == mes:
                    mes_encontrado = int(partes[1])
            except:
                pass 

    if mes_encontrado is None and re.match(r'^\d+$', label_str.strip()):
        dia_num_str = label_str
        mes_encontrado = mes
        
    if dia_num_str and mes_encontrado == mes:
        try:
            dia_limpo = int("".join(filter(str.isdigit, dia_num_str)))
            if 1 <= dia_limpo <= 31:
                return datetime.date(ano, mes, dia_limpo)
        except:
            pass 

    return None

st.set_page_config(layout="wide", initial_sidebar_state="collapsed")

# CSS para expandir containers em full width
st.markdown("""
<style>
    .main > .block-container {
        padding: 1rem;
        max-width: 100%;
    }
    .stContainer {
        max-width: 100%;
    }
    /* Fix para columns ocuparem full width */
    [data-testid="column"] {
        flex: 1 1 calc(100% - 1rem) !important;
    }
    .st-emotion-cache-9edo8l.e1wguzas2 {
        flex: 1 1 calc(100% - 1rem) !important;
    }
</style>
""", unsafe_allow_html=True)

# Inicializa session_state globalmente
if 'idx_arquivo_nav' not in st.session_state:
    st.session_state.idx_arquivo_nav = 0
if 'config_arquivos' not in st.session_state:
    st.session_state.config_arquivos = {}

st.title("🤖 Lançamento de Absenteísmo")
st.write("Com busca LIKE (aproximada) para nomes")

MAPA_CODIGOS = {1: 'P', 2: 'FI', 4: 'FA', 3: 'FÉRIAS-BH', 5: 'DESLIGADO'}

MAPA_CORES = {
    'P': 'FF90EE90',      # Verde claro
    'FI': 'FFFF9999',     # Vermelho suave (rosa claro)
    'FA': 'FFFFFF99',     # Amarelo suave (bege claro)
    'FÉRIAS-BH': 'FF000000',    # Preto (com texto branco)
    'DESLIGADO': 'FF800080',   # Roxo
    'DESCANSO': 'FFC0C0C0'  # Cinza
}

col1, col2 = st.columns(2)

with col1:
    st.header("Upload")
    file_mestra = st.file_uploader("Planilha MESTRA", type=["xlsx"])
    files_encarregado = st.file_uploader("Planilhas ENCARREGADO (múltiplas permitidas)", type=["xlsx"], accept_multiple_files=True)

with col2:
    st.header("Config")
    ano = st.number_input("Ano", 2020, 2050, datetime.date.today().year)
    mes = st.number_input("Mês", 1, 12, datetime.date.today().month)

if files_encarregado:
    st.divider()
    st.header("Pré-Visualização")
    
    # Se há apenas 1 arquivo, processa normalmente
    # Se há múltiplos, mostra navegação
    if len(files_encarregado) == 1:
        file_encarregado = files_encarregado[0]
        idx_arquivo_atual = 0
    else:
        col_prev, col_info, col_next = st.columns([1, 3, 1])
        
        with col_prev:
            if st.button("⬅️ Anterior", key="btn_prev_arquivo"):
                st.session_state.idx_arquivo_nav = max(0, st.session_state.idx_arquivo_nav - 1)
                st.rerun()
        
        with col_info:
            nomes_arquivos = [f.name for f in files_encarregado]
            idx_arq = st.session_state.idx_arquivo_nav
            # Mostra se está configurado
            status = "✅" if nomes_arquivos[idx_arq] in st.session_state.config_arquivos else "⚠️"
            st.info(f"{status} {nomes_arquivos[idx_arq]} ({idx_arq + 1}/{len(files_encarregado)})")
        
        with col_next:
            if st.button("Próximo ➡️", key="btn_next_arquivo"):
                st.session_state.idx_arquivo_nav = min(len(files_encarregado) - 1, st.session_state.idx_arquivo_nav + 1)
                st.rerun()
        
        idx_arquivo_atual = st.session_state.idx_arquivo_nav
        file_encarregado = files_encarregado[idx_arquivo_atual]
    
    # Detecta as guias (sheets) disponíveis no arquivo
    guias_disponiveis = pd.ExcelFile(io.BytesIO(file_encarregado.getvalue())).sheet_names
    
    # Carrega guia salva anteriormente se existir
    nome_arquivo = file_encarregado.name
    if nome_arquivo in st.session_state.config_arquivos and 'guia' in st.session_state.config_arquivos[nome_arquivo]:
        default_guia = st.session_state.config_arquivos[nome_arquivo]['guia']
    else:
        default_guia = guias_disponiveis[0]
    
    # Se há múltiplas guias, deixa o usuário escolher
    if len(guias_disponiveis) > 1:
        guia_selecionada = st.selectbox("Selecione a guia:", guias_disponiveis, index=guias_disponiveis.index(default_guia), key=f"guia_{idx_arquivo_atual}")
    else:
        guia_selecionada = guias_disponiveis[0]
    
    # Guarda a guia selecionada no session_state
    st.session_state.guia_selecionada = guia_selecionada
    buf = io.BytesIO(file_encarregado.getvalue())
    df_raw = pd.read_excel(buf, sheet_name=guia_selecionada, header=None, dtype=str)  # Especifica dtype direto
    
    st.write(f"**Linhas detectadas:** {len(df_raw)} | **Colunas:** {len(df_raw.columns)}")
    
    # Cria nomes em formato Excel (A, B, C...)
    letras_disponíveis = []
    for i in range(len(df_raw.columns)):
        if i < 26:
            letras_disponíveis.append(chr(65 + i))  # A-Z
        else:
            letras_disponíveis.append(f"{chr(65 + i//26 - 1)}{chr(65 + i%26)}")  # AA, AB, etc
    
    # DETECTA AUTOMATICAMENTE qual é a coluna com nomes (testando múltiplas linhas)
    keywords_nomes = ['NOME', 'NOMES', 'COLABORADOR', 'COLABORADORES', 'FUNCIONARIO', 'FUNCIONARIOS', 'EMPLOYEE', 'EMPLOYEES', 'PESSOAL', 'PERSON', 'STAFF']
    col_detectada_auto = None
    idx_col_detectada_auto = None
    
    # Testa as primeiras 10 linhas procurando por keywords
    for linha_teste in range(min(10, len(df_raw))):
        for i in range(len(df_raw.columns) - 1, -1, -1):  # De trás para frente
            header = str(df_raw.iloc[linha_teste, i]).upper().strip()
            for keyword in keywords_nomes:
                if keyword in header:
                    col_detectada_auto = letras_disponíveis[i]
                    idx_col_detectada_auto = i
                    break
            if col_detectada_auto:
                break
        if col_detectada_auto:
            break
    
    # Se não encontrou pela keyword, detecta por conteúdo (muitas letras)
    if col_detectada_auto is None:
        for i in range(len(df_raw.columns) - 1, -1, -1):
            valores = df_raw.iloc[:, i].astype(str).str.strip()
            tem_letras = valores.apply(lambda x: any(c.isalpha() for c in x)).sum() > len(valores) * 0.7
            if tem_letras:
                col_detectada_auto = letras_disponíveis[i]
                idx_col_detectada_auto = i
                break
    
    # Detecta automaticamente qual é a linha com os dias
    # Procura pela primeira linha que tem números em sequência (1, 2, 3, 4, 5...)
    linha_detectada = None
    for tentativa_linha in range(min(20, len(df_raw))):
        valores_linha = [str(df_raw.iloc[tentativa_linha, i]).strip() for i in range(len(df_raw.columns))]
        numeros_encontrados = [v for v in valores_linha if v.isdigit() or v in ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31']]
        if len(numeros_encontrados) >= 15:  # Se tem pelo menos 15 números (dias do mês)
            linha_detectada = tentativa_linha
            break
    
    linhas = [f"Linha {i+1}" for i in range(min(20, len(df_raw)))]
    
    # Carrega configuração salva do arquivo se existir, senão usa default
    nome_arquivo = file_encarregado.name
    if nome_arquivo in st.session_state.config_arquivos:
        config = st.session_state.config_arquivos[nome_arquivo]
        default_linha = f"Linha {config['linha_idx'] + 1}"
        default_col = config['col_idx']
        default_nome_encarregado = config.get('nome_encarregado', '')
    else:
        default_linha = linhas[0]
        default_col = 0
        default_nome_encarregado = ''
    
    # Inicializa selectbox state se não existir (usa configuração salva)
    if f'l_{idx_arquivo_atual}' not in st.session_state:
        st.session_state[f'l_{idx_arquivo_atual}'] = default_linha
    if f'c_{idx_arquivo_atual}' not in st.session_state:
        st.session_state[f'c_{idx_arquivo_atual}'] = letras_disponíveis[default_col]
    if f'encarregado_{idx_arquivo_atual}' not in st.session_state:
        st.session_state[f'encarregado_{idx_arquivo_atual}'] = default_nome_encarregado
    
    c1, c2 = st.columns(2)
    with c1:
        linha_sel = st.selectbox("Linha com DATAS:", linhas, key=f"l_{idx_arquivo_atual}")
        idx_linha = int(linha_sel.split()[1]) - 1  # -1 para voltar ao índice 0 do pandas
    
    with c2:
        col_sel = st.selectbox("Coluna NOMES:", letras_disponíveis, key=f"c_{idx_arquivo_atual}")
        idx_col = letras_disponíveis.index(col_sel)  # Pega o índice baseado na letra

    # Mostra dicas se há diferenças entre detecção e seleção
    tem_dica_linha = linha_detectada is not None and idx_linha != linha_detectada
    tem_dica_coluna = col_detectada_auto is not None and idx_col_detectada_auto != idx_col
    
    if tem_dica_linha:
        st.info(f"💡 **Dica:** Detectei que a linha {linha_detectada + 1} tem os DIAS em sequência. Você selecionou a linha {idx_linha + 1}.")  # +1 para mostrar como Excel
    
    # Mostra dica da coluna logo após a dica da linha
    if tem_dica_coluna:
        st.info(f"💡 **Dica:** Detectei que a coluna **{col_detectada_auto}** tem nomes. Você selecionou a coluna **{col_sel}**.")
    
    # Botão "Aderir Dica" logo após as dicas - só mostra se há dicas
    if tem_dica_linha or tem_dica_coluna:
        col_dica_btn, col_dica_space = st.columns([1, 4])
        with col_dica_btn:
            def aderir_dica():
                if tem_dica_linha:
                    st.session_state[f'l_{idx_arquivo_atual}'] = f"Linha {linha_detectada + 1}"
                if tem_dica_coluna:
                    st.session_state[f'c_{idx_arquivo_atual}'] = col_detectada_auto
            
            st.button("✅ Aderir Dica", key=f"btn_aderir_{idx_arquivo_atual}", on_click=aderir_dica)
    
    # Caixa de texto para o nome do encarregado
    st.write("**👤 Informações do Encarregado:**")
    nome_encarregado = st.text_input("Nome do Encarregado:", placeholder="Digite o nome do encarregado", key=f"encarregado_{idx_arquivo_atual}")
    st.session_state.nome_encarregado = nome_encarregado

    # Salva configuração deste arquivo
    nome_arquivo = file_encarregado.name
    st.session_state.config_arquivos[nome_arquivo] = {
        'linha_idx': idx_linha,
        'col_idx': idx_col,
        'guia': guia_selecionada,
        'nome_encarregado': nome_encarregado
    }
    
    st.session_state.linha_idx = idx_linha
    st.session_state.col_idx = idx_col
    st.session_state.df_raw = df_raw
    
    st.success(f"✅ Linha {idx_linha + 1} + Coluna {col_sel} - Configurado!")  # +1 para mostrar como Excel
    
    # Se a coluna mudou, recarrega o preview com auto-fit
    if 'col_idx_anterior' not in st.session_state or st.session_state.col_idx_anterior != idx_col:
        st.session_state.col_idx_anterior = idx_col
        st.rerun()
    
    # Mostra TODAS as linhas de dados (não só 10)
    # PULA a primeira linha (idx_linha) porque é a linha de cabeçalho
    # INCLUI: a coluna de nomes E TODAS as colunas DEPOIS dela
    colunas_para_manter = [i for i in range(idx_col, len(df_raw.columns))]  # Inclui idx_col também!
    df_prev = df_raw.iloc[idx_linha+1:, colunas_para_manter].copy()
    
    # Cria índice começando em idx_linha+2 (próxima linha após o cabeçalho, em formato Excel)
    df_prev.index = range(idx_linha + 2, idx_linha + 2 + len(df_prev))
    
    # Renomeia colunas para letras (A, B, C, D...) como no Excel APÓS remover
    letras = []
    for i in range(len(df_prev.columns)):
        if i < 26:
            letras.append(chr(65 + i))  # A-Z
        else:
            letras.append(f"{chr(65 + i//26 - 1)}{chr(65 + i%26)}")  # AA, AB, etc
    
    df_prev.columns = letras
    
    # Remove "nan", "None" e NaN do preview - substitui por vazio
    df_prev = df_prev.replace(['nan', 'None', 'NaN', '<NA>'], '')
    df_prev = df_prev.fillna('')
    
    # Remove decimais desnecessários (1.0 -> 1, 4.0 -> 4)
    def remove_decimais(x):
        try:
            if isinstance(x, str) and '.' in x and x.replace('.', '').replace('-', '').isdigit():
                return str(int(float(x)))
        except:
            pass
        return x
    
    for col in df_prev.columns:
        df_prev[col] = df_prev[col].apply(remove_decimais)
    
    # Exibe com st.dataframe normal - key dinâmica força rerender
    st.dataframe(df_prev, width='stretch', height=600, key=f"preview_{idx_col}")

# Botão de processamento com validação
col_btn_processar, col_status = st.columns([1, 3])

with col_btn_processar:
    # Verifica se todos os arquivos foram configurados
    nomes_arquivos_upload = [f.name for f in files_encarregado] if files_encarregado else []
    configs_salvas = list(st.session_state.get('config_arquivos', {}).keys())
    todos_configurados = len(nomes_arquivos_upload) > 0 and all(nome in configs_salvas for nome in nomes_arquivos_upload)
    
    if st.button("🚀 Processar TODOS os Arquivos", disabled=not todos_configurados):
        if file_mestra and files_encarregado and todos_configurados:
            try:
                # Carrega a planilha mestra UMA VEZ
                df_mest = pd.read_excel(file_mestra, header=0)
                
                if 'NOME' not in df_mest.columns:
                    st.error("Coluna NOME não encontrada!")
                    st.stop()
                
                df_mest['NOME_LIMPO'] = df_mest['NOME'].apply(limpar_nome)
                
                mapa_datas = {}
                for col in df_mest.columns:
                    if isinstance(col, (datetime.datetime, datetime.date)):
                        mapa_datas[col.date()] = col
                
                # Pré-preenche TODOS os sábados e domingos com "D" (Descanso)
                st.info("🗓️ Pré-preenchendo todos os fins de semana com 'D'...")
                for col_data_obj in mapa_datas.values():
                    data = col_data_obj if isinstance(col_data_obj, datetime.date) else col_data_obj.date()
                    
                    if eh_fim_de_semana(data):
                        for idx in df_mest.index:
                            df_mest.at[idx, col_data_obj] = 'D'
                
                # Processa CADA arquivo de encarregado
                total_sucesso = 0
                total_erros = []
                total_nomes_unicos = set()
                total_linhas_processadas = set()
                
                with st.spinner('Processando todos os arquivos...'):
                    for idx_arquivo, file_enc in enumerate(files_encarregado):
                        # Recupera a configuração salva deste arquivo
                        config = st.session_state.config_arquivos.get(file_enc.name)
                        if not config:
                            st.warning(f"⚠️ Arquivo {file_enc.name} não foi configurado, pulando...")
                            continue
                        
                        idx_linha = config['linha_idx']
                        idx_col = config['col_idx']
                        guia_usar = config['guia']
                        nome_encarregado = config['nome_encarregado']
                        
                        st.write(f"📄 Processando: **{file_enc.name}**")
                        
                        buf = io.BytesIO(file_enc.getvalue())
                        df_enc = pd.read_excel(buf, sheet_name=guia_usar, header=None, dtype=str)
                        
                        cols_nomes = [str(df_enc.iloc[idx_linha, i]) for i in range(len(df_enc.columns))]
                        df_enc = df_enc.iloc[idx_linha+1:].copy()
                        df_enc.columns = cols_nomes
                        df_enc.reset_index(drop=True, inplace=True)
                        
                        col_nome = cols_nomes[idx_col]
                        cols_datas = cols_nomes[idx_col + 1:]
                        
                        df_enc = df_enc.dropna(how='all')
                        # Usa iloc para pegar a coluna por índice para evitar problema com nomes duplicados
                        df_enc = df_enc[df_enc.iloc[:, idx_col].astype(str).str.strip() != '']
                        df_enc.reset_index(drop=True, inplace=True)
                        
                        # Renomeia a coluna de nomes para algo único para evitar problemas com colunas duplicadas
                        df_enc_temp = df_enc.iloc[:, idx_col:].copy()
                        df_enc_temp.columns = ['___NOME___'] + list(df_enc_temp.columns[1:])
                        
                        df_long = df_enc_temp.melt(
                            id_vars=['___NOME___'],
                            value_vars=cols_datas,
                            var_name='DIA', value_name='COD'
                        )
                        
                        df_long.rename(columns={'___NOME___': 'NOME'}, inplace=True)
                        df_long['NOME_LIMPO'] = df_long['NOME'].apply(limpar_nome)
                        df_long['CODIGO'] = pd.to_numeric(df_long['COD'], errors='coerce')
                        df_long['DATA'] = df_long['DIA'].apply(lambda x: extrair_dia_do_cabecalho(x, mes, ano))
                        df_long = df_long[df_long['NOME_LIMPO'].astype(str).str.strip() != '']
                        df_long = df_long.dropna(subset=['DATA', 'NOME_LIMPO'])
                        
                        sucesso = 0
                        erros = []
                        linhas_processadas = set()
                        nomes_unicos = set()
                        
                        for _, row in df_long.iterrows():
                            nome = row['NOME_LIMPO']
                            cod = row['CODIGO']
                            data = row['DATA']
                            
                            nomes_unicos.add(nome)
                            total_nomes_unicos.add(nome)
                            
                            if pd.isna(cod) or cod not in MAPA_CODIGOS or data not in mapa_datas:
                                continue
                            
                            col_data = mapa_datas[data]
                            
                            # BUSCA EXATA
                            match = df_mest['NOME_LIMPO'] == nome
                            
                            # BUSCA FUZZY (por similaridade)
                            if match.sum() == 0:
                                similaridades = df_mest['NOME_LIMPO'].apply(lambda x: calcular_similaridade(nome, x))
                                match = similaridades >= 0.85
                            
                            # BUSCA POR PALAVRAS-CHAVE
                            if match.sum() == 0:
                                palavras_nome = nome.split()[:3]
                                
                                def contem_palavras_iniciais(nome_mestra):
                                    palavras_mestra = nome_mestra.split()
                                    return all(p in palavras_mestra for p in palavras_nome[:2])
                                
                                match = df_mest['NOME_LIMPO'].apply(contem_palavras_iniciais)
                            
                            if match.sum() > 0:
                                indices_match = df_mest[match].index.tolist()
                                
                                for idx in indices_match:
                                    if df_mest[col_data].dtype != 'object':
                                        df_mest[col_data] = df_mest[col_data].astype('object')
                                    
                                    df_mest.at[idx, col_data] = MAPA_CODIGOS[cod]
                                    linhas_processadas.add(idx)
                                    total_linhas_processadas.add(idx)
                                
                                sucesso += 1
                            else:
                                erros.append(nome)
                        
                        # Atualiza GESTOR para este arquivo (usa o nome_encarregado da configuração)
                        if nome_encarregado and nome_encarregado.strip() != '':
                            if 'GESTOR' in df_mest.columns:
                                for idx in linhas_processadas:
                                    df_mest.at[idx, 'GESTOR'] = nome_encarregado
                        
                        st.success(f"  ✅ {sucesso} lançamentos | 👥 {len(nomes_unicos)} colaboradores únicos")
                        total_sucesso += sucesso
                
                st.divider()
                st.success(f"🎉 Total: ✅ {total_sucesso} lançamentos | 👥 {len(total_nomes_unicos)} colaboradores processados")
                
                if total_erros:
                    with st.expander(f"⚠️ {len(set(total_erros))} não encontrados (de todos os arquivos)"):
                        for e in list(set(total_erros))[:15]:
                            st.write(f"- {e}")
                
                # ===== GERADOR DE RELATÓRIO =====
                st.divider()
                st.header("📊 Relatório Detalhado")
                
                # Seção 1: Colaboradores não processados
                st.subheader("❌ Colaboradores não encontrados")
                colaboradores_nao_encontrados = set(total_erros)
                
                if colaboradores_nao_encontrados:
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        st.write(f"**Total:** {len(colaboradores_nao_encontrados)} colaboradores")
                    with col2:
                        st.write(f"**Motivo:** Não encontrados na Planilha Mestra")
                    
                    with st.expander(f"📋 Ver lista completa ({len(colaboradores_nao_encontrados)} nomes)"):
                        cols_display = st.columns(2)
                        for idx, nome in enumerate(sorted(colaboradores_nao_encontrados)):
                            with cols_display[idx % 2]:
                                st.write(f"• {nome}")
                else:
                    st.success("✅ Todos os colaboradores foram encontrados e processados!")
                
                st.divider()
                
                # Seção 2: Resumo de FI, FA e FÉRIAS-BH por dia (de TODA a planilha mestra)
                st.subheader("📅 Resumo de FI, FA e FÉRIAS-BH por Dia")
                
                # Prepara dados de faltas - vai contar TODA a planilha mestra
                resumo_faltas = []
                
                # Mapa de dias da semana para português
                dias_semana_pt = {
                    'MON': 'SEG', 'TUE': 'TER', 'WED': 'QUA', 'THU': 'QUI',
                    'FRI': 'SEX', 'SAT': 'SÁB', 'SUN': 'DOM'
                }
                
                for data_obj in sorted(mapa_datas.keys()):
                    col_data = mapa_datas[data_obj]
                    if col_data in df_mest.columns:
                        total_fi = (df_mest[col_data] == 'FI').sum()
                        total_fa = (df_mest[col_data] == 'FA').sum()
                        total_ferias = (df_mest[col_data] == 'FÉRIAS-BH').sum()
                        total_lancamentos = total_fi + total_fa + total_ferias
                        
                        if total_lancamentos > 0:  # Só mostra dias com registros
                            data_formatada = data_obj.strftime('%d/%m/%Y') if isinstance(data_obj, datetime.date) else str(data_obj)
                            dia_en = data_obj.strftime('%a').upper() if isinstance(data_obj, datetime.date) else '???'
                            dia_semana = dias_semana_pt.get(dia_en, dia_en)
                            
                            resumo_faltas.append({
                                'Data': data_formatada,
                                'Dia': dia_semana,
                                'FI': total_fi,
                                'FA': total_fa,
                                'FÉRIAS-BH': total_ferias,
                                'Total': total_lancamentos
                            })
                
                if resumo_faltas:
                    df_resumo = pd.DataFrame(resumo_faltas)
                    
                    # Exibe a tabela com formatação
                    st.dataframe(
                        df_resumo,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            'Data': st.column_config.TextColumn(width=120),
                            'Dia': st.column_config.TextColumn(width=70),
                            'FI': st.column_config.NumberColumn(width=80),
                            'FA': st.column_config.NumberColumn(width=80),
                            'FÉRIAS-BH': st.column_config.NumberColumn(width=100),
                            'Total': st.column_config.NumberColumn(width=80)
                        }
                    )
                    
                    # Totalizador geral - 2 linhas de métricas
                    st.divider()
                    col1a, col1b, col1c = st.columns(3)
                    with col1a:
                        st.metric("📌 FI (Justificadas)", df_resumo['FI'].sum())
                    with col1b:
                        st.metric("🚫 FA (Não Justificadas)", df_resumo['FA'].sum())
                    with col1c:
                        st.metric("🏖️ FÉRIAS-BH", df_resumo['FÉRIAS-BH'].sum())
                    
                    col2a, col2b, col2c = st.columns([1, 1, 1])
                    with col2b:
                        st.metric("📊 TOTAL GERAL", df_resumo['Total'].sum())
                else:
                    st.info("✅ Nenhuma falta registrada!")
                
                st.divider()
                out = io.BytesIO()
                df_mest_final = df_mest.drop(columns=['NOME_LIMPO'])
                
                with pd.ExcelWriter(out, engine='openpyxl') as w:
                    df_mest_final.to_excel(w, index=False, sheet_name='Dados')
                    
                    worksheet = w.sheets['Dados']
                    
                    # ===== FORMATAÇÃO DO HEADER =====
                    header_fill = PatternFill(start_color='FF366092', end_color='FF366092', fill_type='solid')  # Azul escuro
                    header_font = Font(bold=True, color='FFFFFFFF', size=11)  # Texto branco
                    
                    # Formata todas as colunas do header
                    for col_idx in range(1, len(df_mest_final.columns) + 1):
                        header_cell = worksheet.cell(row=1, column=col_idx)
                        header_cell.fill = header_fill
                        header_cell.font = header_font
                        header_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    
                    # ===== FORMATAÇÃO DAS COLUNAS ESPECÍFICAS =====
                    # Mapeamento de colunas com cores
                    col_names = df_mest_final.columns.tolist()
                    
                    # Função para calcular largura baseada no maior valor da coluna
                    def calc_width(df, col_name, min_width=10, max_width=50):
                        if col_name not in df.columns:
                            return min_width
                        max_len = df[col_name].astype(str).str.len().max()
                        header_len = len(str(col_name))
                        largest = max(max_len, header_len)
                        width = min(max(largest + 2, min_width), max_width)
                        return width
                    
                    for col_idx, col_name in enumerate(col_names, 1):
                        # Define preenchimento e largura para cada tipo de coluna
                        if col_name == 'NOME':
                            col_fill = PatternFill(start_color='FFCCE5FF', end_color='FFCCE5FF', fill_type='solid')  # Azul claro suave
                            width = calc_width(df_mest_final, col_name, min_width=15, max_width=40)
                            worksheet.column_dimensions[get_column_letter(col_idx)].width = width
                        elif col_name == 'AREA':
                            col_fill = PatternFill(start_color='FFC6EFCE', end_color='FFC6EFCE', fill_type='solid')  # Verde claro suave
                            worksheet.column_dimensions[get_column_letter(col_idx)].width = 25
                        elif col_name == 'GESTOR':
                            col_fill = PatternFill(start_color='FFffbf5e', end_color='FFffbf5e', fill_type='solid')  # Laranja #ffbf5e
                            width = calc_width(df_mest_final, col_name, min_width=15, max_width=40)
                            worksheet.column_dimensions[get_column_letter(col_idx)].width = width
                        else:
                            col_fill = None
                            # Largura fixa para outras colunas
                            try:
                                datetime.datetime.strptime(str(col_name), '%d/%m')
                                worksheet.column_dimensions[get_column_letter(col_idx)].width = 3
                            except:
                                worksheet.column_dimensions[get_column_letter(col_idx)].width = 10
                        
                        # Aplica a cor a todas as linhas de dados desta coluna
                        if col_fill is not None:
                            for row_idx in range(2, worksheet.max_row + 1):
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                cell.fill = col_fill
                    
                    for col_data_obj in mapa_datas.values():
                        col_idx = list(df_mest_final.columns).index(col_data_obj) + 1
                        
                        header_cell = worksheet.cell(row=1, column=col_idx)
                        if isinstance(col_data_obj, (datetime.datetime, datetime.date)):
                            data_formatada = col_data_obj.strftime('%d/%m') if isinstance(col_data_obj, datetime.date) else col_data_obj.date().strftime('%d/%m')
                            header_cell.value = data_formatada
                        
                        for row_idx, row in enumerate(worksheet.iter_rows(min_col=col_idx, max_col=col_idx, min_row=2), start=2):
                            for cell in row:
                                cell.number_format = 'DD/MM'
                                
                                valor = str(cell.value).strip() if cell.value else ''
                                
                                if valor == 'P':
                                    cell.fill = PatternFill(start_color=MAPA_CORES['P'], end_color=MAPA_CORES['P'], fill_type='solid')
                                elif valor == 'FI':
                                    cell.fill = PatternFill(start_color=MAPA_CORES['FI'], end_color=MAPA_CORES['FI'], fill_type='solid')
                                elif valor == 'FA':
                                    cell.fill = PatternFill(start_color=MAPA_CORES['FA'], end_color=MAPA_CORES['FA'], fill_type='solid')
                                elif valor == 'FÉRIAS-BH':
                                    cell.fill = PatternFill(start_color=MAPA_CORES['FÉRIAS-BH'], end_color=MAPA_CORES['FÉRIAS-BH'], fill_type='solid')
                                    cell.font = Font(color='FFFFFFFF')
                                elif valor == 'DESLIGADO':
                                    cell.fill = PatternFill(start_color=MAPA_CORES['DESLIGADO'], end_color=MAPA_CORES['DESLIGADO'], fill_type='solid')
                                    cell.font = Font(color='FFFFFFFF')
                                elif valor == 'D':
                                    cell.fill = PatternFill(start_color=MAPA_CORES['DESCANSO'], end_color=MAPA_CORES['DESCANSO'], fill_type='solid')
                    
                    # ===== CRIAR GUIA DE RELATÓRIO =====
                    ws_relatorio = w.book.create_sheet('Relatório')
                    
                    # Linha 1: Título
                    titulo_cell = ws_relatorio.cell(row=1, column=1, value='📊 RELATÓRIO DE PROCESSAMENTO')
                    titulo_cell.font = Font(bold=True, size=14, color='FFFFFF')
                    titulo_cell.fill = PatternFill(start_color='FF366092', end_color='FF366092', fill_type='solid')
                    ws_relatorio.merge_cells('A1:D1')
                    
                    # Linha 2: Data/Hora
                    ws_relatorio.cell(row=2, column=1, value='Data do Processamento:')
                    ws_relatorio.cell(row=2, column=2, value=datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'))
                    
                    ws_relatorio.cell(row=2, column=3, value='Mês/Ano:')
                    ws_relatorio.cell(row=2, column=4, value=f"{mes:02d}/{ano}")
                    
                    # Linha 4: Resumo Geral
                    ws_relatorio.cell(row=4, column=1, value='RESUMO GERAL')
                    ws_relatorio.cell(row=4, column=1).font = Font(bold=True, size=12)
                    ws_relatorio.cell(row=4, column=1).fill = PatternFill(start_color='FFC5D9F1', end_color='FFC5D9F1', fill_type='solid')
                    ws_relatorio.merge_cells('A4:D4')
                    
                    ws_relatorio.cell(row=5, column=1, value='Total de Lançamentos Processados:')
                    ws_relatorio.cell(row=5, column=2, value=total_sucesso)
                    
                    ws_relatorio.cell(row=6, column=1, value='Total de Colaboradores Únicos:')
                    ws_relatorio.cell(row=6, column=2, value=len(total_nomes_unicos))
                    
                    ws_relatorio.cell(row=7, column=1, value='Total de Não Encontrados:')
                    ws_relatorio.cell(row=7, column=2, value=len(set(total_erros)))
                    
                    # Linha 9: Resumo por Dia
                    ws_relatorio.cell(row=9, column=1, value='RESUMO POR DIA')
                    ws_relatorio.cell(row=9, column=1).font = Font(bold=True, size=12)
                    ws_relatorio.cell(row=9, column=1).fill = PatternFill(start_color='FFC5D9F1', end_color='FFC5D9F1', fill_type='solid')
                    ws_relatorio.merge_cells('A9:F9')
                    
                    # Headers da tabela de resumo
                    headers_resumo = ['Data', 'Dia', 'FI', 'FA', 'FÉRIAS-BH', 'Total']
                    for col_idx, header in enumerate(headers_resumo, 1):
                        cell = ws_relatorio.cell(row=10, column=col_idx, value=header)
                        cell.font = Font(bold=True, color='FFFFFF')
                        cell.fill = PatternFill(start_color='FF4472C4', end_color='FF4472C4', fill_type='solid')
                    
                    # Preenche tabela de resumo
                    dias_semana_pt = {
                        'MON': 'SEG', 'TUE': 'TER', 'WED': 'QUA', 'THU': 'QUI',
                        'FRI': 'SEX', 'SAT': 'SÁB', 'SUN': 'DOM'
                    }
                    
                    row_idx = 11
                    for data_obj in sorted(mapa_datas.keys()):
                        col_data = mapa_datas[data_obj]
                        if col_data in df_mest.columns:
                            total_fi = (df_mest[col_data] == 'FI').sum()
                            total_fa = (df_mest[col_data] == 'FA').sum()
                            total_ferias = (df_mest[col_data] == 'FÉRIAS-BH').sum()
                            total_lancamentos = total_fi + total_fa + total_ferias
                            
                            if total_lancamentos > 0:
                                data_formatada = data_obj.strftime('%d/%m/%Y') if isinstance(data_obj, datetime.date) else str(data_obj)
                                dia_en = data_obj.strftime('%a').upper() if isinstance(data_obj, datetime.date) else '???'
                                dia_semana = dias_semana_pt.get(dia_en, dia_en)
                                
                                # Coluna Data (cinza)
                                cell_data = ws_relatorio.cell(row=row_idx, column=1, value=data_formatada)
                                cell_data.fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
                                
                                # Coluna Dia (cinza)
                                cell_dia = ws_relatorio.cell(row=row_idx, column=2, value=dia_semana)
                                cell_dia.fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
                                
                                # Coluna FI (vermelho)
                                cell_fi = ws_relatorio.cell(row=row_idx, column=3, value=int(total_fi))
                                cell_fi.fill = PatternFill(start_color=MAPA_CORES['FI'], end_color=MAPA_CORES['FI'], fill_type='solid')
                                
                                # Coluna FA (amarelo)
                                cell_fa = ws_relatorio.cell(row=row_idx, column=4, value=int(total_fa))
                                cell_fa.fill = PatternFill(start_color=MAPA_CORES['FA'], end_color=MAPA_CORES['FA'], fill_type='solid')
                                
                                # Coluna FÉRIAS-BH (preto)
                                cell_ferias = ws_relatorio.cell(row=row_idx, column=5, value=int(total_ferias))
                                cell_ferias.fill = PatternFill(start_color=MAPA_CORES['FÉRIAS-BH'], end_color=MAPA_CORES['FÉRIAS-BH'], fill_type='solid')
                                cell_ferias.font = Font(color='FFFFFFFF')  # Texto branco
                                
                                ws_relatorio.cell(row=row_idx, column=6, value=int(total_lancamentos))
                                
                                row_idx += 1
                    
                    # Linha de Resumo por Departamento (DIÁRIO)
                    row_departamento = row_idx + 2
                    ws_relatorio.cell(row=row_departamento, column=1, value='RESUMO POR DEPARTAMENTO (DIÁRIO)')
                    ws_relatorio.cell(row=row_departamento, column=1).font = Font(bold=True, size=12)
                    ws_relatorio.cell(row=row_departamento, column=1).fill = PatternFill(start_color='FFC5D9F1', end_color='FFC5D9F1', fill_type='solid')
                    ws_relatorio.merge_cells(f'A{row_departamento}:H{row_departamento}')
                    
                    # Mapeia setores para departamentos
                    setores_ma_bloq = ['MOVIMENTACAO E ARMAZENAGEM', 'BLOQ', 'PROJETO INTERPRISE - MOVIMENTACAO E ARMAZENAGEM']
                    setores_crdk_de = ['CRDK D&E|CD-RJ HB', 'CROSSDOCK DISTRIBUICAO E EXPEDICAO']
                    
                    # Headers do resumo por departamento (com datas)
                    row_departamento += 1
                    headers_depto = ['Data', 'Dia', 'Depto', 'FI', 'FA', 'Total']
                    for col_idx, header in enumerate(headers_depto, 1):
                        cell = ws_relatorio.cell(row=row_departamento, column=col_idx, value=header)
                        cell.font = Font(bold=True, color='FFFFFF')
                        cell.fill = PatternFill(start_color='FF4472C4', end_color='FF4472C4', fill_type='solid')
                    
                    # Função para contar FI e FA por departamento e data
                    def contar_fi_fa_por_depto_data(df, setores_lista, col_area, data_col):
                        total_fi = 0
                        total_fa = 0
                        
                        for setor in setores_lista:
                            # Filtra colaboradores deste setor
                            mask_setor = df[col_area].astype(str).str.contains(setor, case=False, na=False)
                            df_setor = df[mask_setor]
                            
                            if not df_setor.empty and data_col in df.columns:
                                # Conta FI e FA para esta data
                                total_fi += (df_setor[data_col] == 'FI').sum()
                                total_fa += (df_setor[data_col] == 'FA').sum()
                        
                        return total_fi, total_fa
                    
                    # Preenche tabela com dados por dia
                    if 'AREA' in df_mest.columns:
                        row_departamento += 1
                        for data_obj in sorted(mapa_datas.keys()):
                            col_data = mapa_datas[data_obj]
                            if col_data in df_mest.columns:
                                data_formatada = data_obj.strftime('%d/%m/%Y') if isinstance(data_obj, datetime.date) else str(data_obj)
                                dia_en = data_obj.strftime('%a').upper() if isinstance(data_obj, datetime.date) else '???'
                                dia_semana = dias_semana_pt.get(dia_en, dia_en)
                                
                                # M&A / BLOQ
                                fi_ma_bloq, fa_ma_bloq = contar_fi_fa_por_depto_data(df_mest, setores_ma_bloq, 'AREA', col_data)
                                
                                if fi_ma_bloq > 0 or fa_ma_bloq > 0:
                                    # Coluna Data (cinza)
                                    cell_data = ws_relatorio.cell(row=row_departamento, column=1, value=data_formatada)
                                    cell_data.fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
                                    
                                    # Coluna Dia (cinza)
                                    cell_dia = ws_relatorio.cell(row=row_departamento, column=2, value=dia_semana)
                                    cell_dia.fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
                                    
                                    # Coluna Depto (verde suave)
                                    cell_depto = ws_relatorio.cell(row=row_departamento, column=3, value='M&A / BLOQ')
                                    cell_depto.fill = PatternFill(start_color='FFD5E8D4', end_color='FFD5E8D4', fill_type='solid')
                                    
                                    # Coluna FI (vermelho suave)
                                    cell_fi = ws_relatorio.cell(row=row_departamento, column=4, value=int(fi_ma_bloq))
                                    cell_fi.fill = PatternFill(start_color=MAPA_CORES['FI'], end_color=MAPA_CORES['FI'], fill_type='solid')
                                    
                                    # Coluna FA (amarelo suave)
                                    cell_fa = ws_relatorio.cell(row=row_departamento, column=5, value=int(fa_ma_bloq))
                                    cell_fa.fill = PatternFill(start_color=MAPA_CORES['FA'], end_color=MAPA_CORES['FA'], fill_type='solid')
                                    
                                    ws_relatorio.cell(row=row_departamento, column=6, value=int(fi_ma_bloq + fa_ma_bloq))
                                    row_departamento += 1
                                
                                # CRDK / D&E
                                fi_crdk_de, fa_crdk_de = contar_fi_fa_por_depto_data(df_mest, setores_crdk_de, 'AREA', col_data)
                                
                                if fi_crdk_de > 0 or fa_crdk_de > 0:
                                    # Coluna Data (cinza)
                                    cell_data = ws_relatorio.cell(row=row_departamento, column=1, value=data_formatada)
                                    cell_data.fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
                                    
                                    # Coluna Dia (cinza)
                                    cell_dia = ws_relatorio.cell(row=row_departamento, column=2, value=dia_semana)
                                    cell_dia.fill = PatternFill(start_color='FFD3D3D3', end_color='FFD3D3D3', fill_type='solid')
                                    
                                    # Coluna Depto (verde suave)
                                    cell_depto = ws_relatorio.cell(row=row_departamento, column=3, value='CRDK / D&E')
                                    cell_depto.fill = PatternFill(start_color='FFD5E8D4', end_color='FFD5E8D4', fill_type='solid')
                                    
                                    # Coluna FI (vermelho suave)
                                    cell_fi = ws_relatorio.cell(row=row_departamento, column=4, value=int(fi_crdk_de))
                                    cell_fi.fill = PatternFill(start_color=MAPA_CORES['FI'], end_color=MAPA_CORES['FI'], fill_type='solid')
                                    
                                    # Coluna FA (amarelo suave)
                                    cell_fa = ws_relatorio.cell(row=row_departamento, column=5, value=int(fa_crdk_de))
                                    cell_fa.fill = PatternFill(start_color=MAPA_CORES['FA'], end_color=MAPA_CORES['FA'], fill_type='solid')
                                    
                                    ws_relatorio.cell(row=row_departamento, column=6, value=int(fi_crdk_de + fa_crdk_de))
                                    row_departamento += 1
                    
                    # Linha de Não Encontrados
                    row_nao_encontrados = row_departamento + 2
                    ws_relatorio.cell(row=row_nao_encontrados, column=1, value='COLABORADORES NÃO ENCONTRADOS')
                    ws_relatorio.cell(row=row_nao_encontrados, column=1).font = Font(bold=True, size=12)
                    ws_relatorio.cell(row=row_nao_encontrados, column=1).fill = PatternFill(start_color='FFC5D9F1', end_color='FFC5D9F1', fill_type='solid')
                    ws_relatorio.merge_cells(f'A{row_nao_encontrados}:D{row_nao_encontrados}')
                    
                    row_nao_encontrados += 1
                    if colaboradores_nao_encontrados:
                        for nome in sorted(colaboradores_nao_encontrados):
                            ws_relatorio.cell(row=row_nao_encontrados, column=1, value=nome)
                            row_nao_encontrados += 1
                    else:
                        ws_relatorio.cell(row=row_nao_encontrados, column=1, value='✅ Todos encontrados!')
                    
                    # Ajusta largura das colunas
                    ws_relatorio.column_dimensions['A'].width = 20
                    ws_relatorio.column_dimensions['B'].width = 15
                    ws_relatorio.column_dimensions['C'].width = 10
                    ws_relatorio.column_dimensions['D'].width = 10
                    ws_relatorio.column_dimensions['E'].width = 15
                    ws_relatorio.column_dimensions['F'].width = 10
                    
                    out.seek(0)
                
                st.download_button(
                    "📥 Download - Planilha MESTRA Completa",
                    out.getvalue(),
                    f"Mestra_Completa_{ano}-{mes:02d}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ Erro durante o processamento: {str(e)}")

# CSS final para garantir full-width em todo o app
st.markdown("""
<style>
    * {
        box-sizing: border-box !important;
    }
    
    [data-testid="stAppViewContainer"] {
        max-width: 100% !important;
    }
    
    [data-testid="stMain"] {
        max-width: 100% !important;
    }
    
    .main > .block-container {
        max-width: 100% !important;
        padding: 1rem !important;
    }
    
    .stTabs [role="tablist"] {
        width: 100% !important;
    }
    
    [data-testid="column"] {
        flex: 1 1 calc(100% - 1rem) !important;
        width: 100% !important;
    }
    
    .st-emotion-cache-9edo8l.e1wguzas2 {
        flex: 1 1 calc(100% - 1rem) !important;
        width: 100% !important;
    }
    
    section[data-testid="stContainer"] {
        width: 100% !important;
    }
    
    div.stDataFrame {
        width: 100% !important;
        max-width: 100% !important;
    }
    
    .dataframe {
        width: 100% !important;
    }
</style>
""", unsafe_allow_html=True)
