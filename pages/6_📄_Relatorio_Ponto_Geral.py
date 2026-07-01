import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import io
import zipfile
import re
from typing import Dict, List, Tuple
from unidecode import unidecode

st.set_page_config(page_title="Relatório de Ponto Geral", layout="wide")

st.title("📄 Relatório de Ponto Geral")
st.markdown("""
Esta página gera relatórios de ocorrências por tipo, agrupando colaboradores e gerando rankings.

**Fluxo:**
1. Carregue o arquivo Excel (XLSX) de ponto
2. Selecione os tipos de ocorrência
3. Baixe as planilhas geradas (zipadas)
""")

# ============================================================
# CORES CORPORATIVAS
# ============================================================
COR_VERDE_CLARO = '#8CC83C'
COR_VERDE_MEDIO = '#008A4B'
COR_VERDE_ESCURO = '#006450'
COR_CINZA_CLARO = '#F2F2F2'
COR_PRETO = '#000000'
COR_BRANCO = '#FFFFFF'
COR_VERMELHO = '#CC0000'
COR_LARANJA = '#FF8C00'


def calcular_tempo_servico(data_admissao_str: str) -> str:
    if pd.isna(data_admissao_str) or str(data_admissao_str).strip() == '':
        return 'N/A'
    try:
        data_str = str(data_admissao_str).strip()
        for fmt in ['%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d']:
            try:
                data_adm = datetime.strptime(data_str, fmt)
                break
            except ValueError:
                continue
        else:
            try:
                data_adm = pd.to_datetime(data_str, dayfirst=True)
                if pd.isna(data_adm):
                    return 'N/A'
            except:
                return 'N/A'
        hoje = datetime.now()
        anos = hoje.year - data_adm.year
        meses = hoje.month - data_adm.month
        if meses < 0:
            anos -= 1
            meses += 12
        if hoje.day < data_adm.day:
            meses -= 1
            if meses < 0:
                anos -= 1
                meses += 12
        if anos > 0 and meses > 0:
            return f'{anos}a {meses}m'
        elif anos > 0:
            return f'{anos}a'
        elif meses > 0:
            return f'{meses}m'
        else:
            return '< 1m'
    except Exception:
        return 'N/A'


def formatar_data_br(data_str: str) -> str:
    if pd.isna(data_str) or str(data_str).strip() == '':
        return ''
    try:
        data_str = str(data_str).strip()
        for fmt in ['%m/%d/%Y', '%m/%d/%y', '%Y-%m-%d', '%d/%m/%Y']:
            try:
                data_dt = datetime.strptime(data_str, fmt)
                return data_dt.strftime('%d/%m/%Y')
            except ValueError:
                continue
        try:
            data_dt = pd.to_datetime(data_str, dayfirst=False)
            if pd.notna(data_dt):
                return data_dt.strftime('%d/%m/%Y')
        except:
            pass
        return data_str
    except:
        return data_str


def processar_ocorrencia(
    df: pd.DataFrame,
    termo_ocorrencia: str,
    termo_justificativa: str,
    col_nome: str,
    col_cargo: str,
    col_depto: str,
    col_data_adm: str,
    col_ocorrencia: str,
    col_justificativa: str,
    col_data: str
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Processa um filtro de ocorrência+justificativa e retorna (detalhe, ranking)."""
    termo_occ_norm = unidecode(termo_ocorrencia).strip().lower()
    termo_just_norm = unidecode(termo_justificativa).strip().lower()
    
    occ_norm = df[col_ocorrencia].astype(str).apply(lambda x: unidecode(x).strip().lower())
    just_norm = df[col_justificativa].astype(str).apply(lambda x: unidecode(x).strip().lower())
    
    # Filtra pela ocorrência
    mask_occ = occ_norm == termo_occ_norm
    if mask_occ.sum() == 0:
        mask_occ = occ_norm.str.contains(termo_occ_norm, na=False)
    if mask_occ.sum() == 0:
        palavras = termo_occ_norm.split()
        for palavra in reversed(palavras):
            if len(palavra) > 3:
                mask_occ = occ_norm.str.contains(palavra, na=False)
                if mask_occ.sum() > 0:
                    break
    
    # Filtra pela justificativa
    mask_just = just_norm == termo_just_norm
    if mask_just.sum() == 0:
        mask_just = just_norm.str.contains(termo_just_norm, na=False)
    if mask_just.sum() == 0:
        palavras = termo_just_norm.split()
        for palavra in reversed(palavras):
            if len(palavra) > 3:
                mask_just = just_norm.str.contains(palavra, na=False)
                if mask_just.sum() > 0:
                    break
    
    # Se justificativa não achou nada, usa só a ocorrência
    if mask_just.sum() == 0:
        mask_just = pd.Series([True] * len(df))
    
    df_filtrado = df[mask_occ & mask_just].copy()
    
    if len(df_filtrado) == 0:
        colunas_detalhe = [
            'Colaborador', 'Cargo', 'Departamento', 'Data Admissão',
            'Tempo de Serviço', 'Quantidade Ocorrências', 'Datas das Ocorrências'
        ]
        df_detalhe = pd.DataFrame(columns=colunas_detalhe)
        df_ranking = pd.DataFrame(columns=['Colaborador', 'Cargo', 'Departamento', 'Data Admissão', 'Tempo de Serviço', 'Quantidade Ocorrências'])
        return df_detalhe, df_ranking
    
    df_filtrado['Data_Formatada'] = df_filtrado[col_data].apply(formatar_data_br)
    df_filtrado['Tempo_Servico'] = df_filtrado[col_data_adm].apply(calcular_tempo_servico)
    
    grupos = df_filtrado.groupby(col_nome)
    
    dados_detalhe = []
    for nome, grupo in grupos:
        primeiro = grupo.iloc[0]
        datas = sorted(grupo['Data_Formatada'].dropna().unique().tolist())
        datas_str = ', '.join(datas) if datas else ''
        
        dados_detalhe.append({
            'Colaborador': nome,
            'Cargo': primeiro[col_cargo] if col_cargo in grupo.columns else '',
            'Departamento': primeiro[col_depto] if col_depto in grupo.columns else '',
            'Data Admissão': primeiro[col_data_adm] if col_data_adm in grupo.columns else '',
            'Tempo de Serviço': primeiro['Tempo_Servico'],
            'Quantidade Ocorrências': len(grupo),
            'Datas das Ocorrências': datas_str
        })
    
    df_detalhe = pd.DataFrame(dados_detalhe)
    df_detalhe = df_detalhe.sort_values('Quantidade Ocorrências', ascending=False).reset_index(drop=True)
    
    df_ranking = df_detalhe[['Colaborador', 'Cargo', 'Departamento', 'Data Admissão', 'Tempo de Serviço', 'Quantidade Ocorrências']].copy()
    df_ranking = df_ranking.sort_values('Quantidade Ocorrências', ascending=False).reset_index(drop=True)
    df_ranking['Posição'] = range(1, len(df_ranking) + 1)
    df_ranking = df_ranking[['Posição', 'Colaborador', 'Cargo', 'Departamento', 'Data Admissão', 'Tempo de Serviço', 'Quantidade Ocorrências']]
    
    return df_detalhe, df_ranking


def gerar_planilha_ocorrencia(
    df_detalhe: pd.DataFrame,
    df_ranking: pd.DataFrame,
    nome_aba_detalhe: str,
    writer: pd.ExcelWriter
):
    """Escreve as duas abas de uma ocorrência no Excel writer."""
    workbook = writer.book
    
    header_fmt = workbook.add_format({
        'bold': True, 'font_size': 11,
        'font_color': COR_BRANCO, 'bg_color': COR_VERDE_ESCURO,
        'border': 1, 'border_color': COR_VERDE_ESCURO,
        'text_wrap': True, 'valign': 'vcenter', 'align': 'center'
    })
    
    header_ranking_fmt = workbook.add_format({
        'bold': True, 'font_size': 11,
        'font_color': COR_BRANCO, 'bg_color': COR_VERDE_MEDIO,
        'border': 1, 'border_color': COR_VERDE_MEDIO,
        'text_wrap': True, 'valign': 'vcenter', 'align': 'center'
    })
    
    cell_fmt = workbook.add_format({
        'font_size': 10, 'border': 1,
        'border_color': '#BFBFBF', 'text_wrap': True, 'valign': 'vcenter'
    })
    
    cell_alt_fmt = workbook.add_format({
        'font_size': 10, 'border': 1,
        'border_color': '#BFBFBF', 'bg_color': COR_CINZA_CLARO,
        'text_wrap': True, 'valign': 'vcenter'
    })
    
    cell_qtd_fmt = workbook.add_format({
        'font_size': 11, 'bold': True, 'font_color': COR_PRETO,
        'border': 1, 'border_color': '#BFBFBF',
        'text_wrap': True, 'valign': 'vcenter', 'align': 'center'
    })
    
    cell_qtd_alt_fmt = workbook.add_format({
        'font_size': 11, 'bold': True, 'font_color': COR_PRETO,
        'border': 1, 'border_color': '#BFBFBF', 'bg_color': COR_CINZA_CLARO,
        'text_wrap': True, 'valign': 'vcenter', 'align': 'center'
    })
    
    cell_alta_fmt = workbook.add_format({
        'font_size': 11, 'bold': True, 'font_color': COR_VERMELHO,
        'border': 1, 'border_color': '#BFBFBF', 'bg_color': '#FFF0F0',
        'text_wrap': True, 'valign': 'vcenter', 'align': 'center'
    })
    
    cell_media_fmt = workbook.add_format({
        'font_size': 11, 'bold': True, 'font_color': COR_LARANJA,
        'border': 1, 'border_color': '#BFBFBF', 'bg_color': '#FFF8E7',
        'text_wrap': True, 'valign': 'vcenter', 'align': 'center'
    })
    
    posicoes_format = {
        1: workbook.add_format({'font_size': 12, 'bold': True, 'font_color': COR_BRANCO, 'bg_color': '#D4A017', 'border': 1, 'border_color': '#BFBFBF', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'}),
        2: workbook.add_format({'font_size': 12, 'bold': True, 'font_color': '#333333', 'bg_color': '#C0C0C0', 'border': 1, 'border_color': '#BFBFBF', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'}),
        3: workbook.add_format({'font_size': 12, 'bold': True, 'font_color': COR_BRANCO, 'bg_color': '#CD7F32', 'border': 1, 'border_color': '#BFBFBF', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
    }
    
    # --- ABA 1: DETALHAMENTO ---
    sheet_name = nome_aba_detalhe[:31]
    df_detalhe.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
    ws = writer.sheets[sheet_name]
    
    for col_num, value in enumerate(df_detalhe.columns.values):
        ws.write(0, col_num, value, header_fmt)
    
    ws.set_column(0, 0, 38)  # Colaborador
    ws.set_column(1, 1, 22)  # Cargo
    ws.set_column(2, 2, 32)  # Departamento
    ws.set_column(3, 3, 16)  # Data Admissão
    ws.set_column(4, 4, 16)  # Tempo de Serviço
    ws.set_column(5, 5, 20)  # Quantidade
    ws.set_column(6, 6, 55)  # Datas
    
    for row_num in range(1, len(df_detalhe) + 1):
        is_alt = row_num % 2 == 0
        for col_num in range(len(df_detalhe.columns)):
            valor = df_detalhe.iloc[row_num - 1, col_num]
            if col_num == 5:
                qtd = valor
                if qtd >= 5:
                    fmt = cell_alta_fmt
                elif qtd >= 3:
                    fmt = cell_media_fmt
                elif is_alt:
                    fmt = cell_qtd_alt_fmt
                else:
                    fmt = cell_qtd_fmt
            elif is_alt:
                fmt = cell_alt_fmt
            else:
                fmt = cell_fmt
            ws.write(row_num, col_num, valor, fmt)
    
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(df_detalhe), len(df_detalhe.columns) - 1)
    
    # --- ABA 2: RANKING ---
    sheet_name_rank = f'Ofensores'
    # Garante nome único
    if sheet_name_rank in writer.sheets:
        sheet_name_rank = f'Ofensores 2'
    df_ranking.to_excel(writer, sheet_name=sheet_name_rank[:31], index=False, startrow=0)
    ws_rank = writer.sheets[sheet_name_rank[:31]]
    
    for col_num, value in enumerate(df_ranking.columns.values):
        ws_rank.write(0, col_num, value, header_ranking_fmt)
    
    ws_rank.set_column(0, 0, 10)
    ws_rank.set_column(1, 1, 38)
    ws_rank.set_column(2, 2, 22)
    ws_rank.set_column(3, 3, 32)
    ws_rank.set_column(4, 4, 16)
    ws_rank.set_column(5, 5, 16)
    ws_rank.set_column(6, 6, 20)
    
    for row_num in range(1, len(df_ranking) + 1):
        is_alt = row_num % 2 == 0
        posicao = df_ranking.iloc[row_num - 1, 0]
        for col_num in range(len(df_ranking.columns)):
            valor = df_ranking.iloc[row_num - 1, col_num]
            if col_num == 0 and posicao in posicoes_format:
                fmt = posicoes_format[posicao]
            elif col_num == 6:
                qtd = valor
                if qtd >= 5:
                    fmt = cell_alta_fmt
                elif qtd >= 3:
                    fmt = cell_media_fmt
                elif is_alt:
                    fmt = cell_qtd_alt_fmt
                else:
                    fmt = cell_qtd_fmt
            elif is_alt:
                fmt = cell_alt_fmt
            else:
                fmt = cell_fmt
            ws_rank.write(row_num, col_num, valor, fmt)
    
    ws_rank.freeze_panes(1, 0)
    ws_rank.autofilter(0, 0, len(df_ranking), len(df_ranking.columns) - 1)
    
    ws.set_row(0, 30)
    ws_rank.set_row(0, 30)


def sanitizar_nome_arquivo(nome: str) -> str:
    """Remove caracteres especiais para nome de arquivo."""
    nome = unidecode(nome)
    nome = re.sub(r'[<>:"/\\|?*]', '', nome)
    nome = re.sub(r'\s+', '_', nome.strip())
    nome = re.sub(r'_+', '_', nome)
    return nome[:80]


# ============================================================
# CONFIGURAÇÃO DE TODAS AS OCORRÊNCIAS
# ============================================================
# Cada item pode ser:
# - Dict simples: ocorrência com 1 justificativa (ocorrencia=justificativa)
# - Dict com 'multiplas': dicionário {nome_pasta: lista_de_justificativas}

OCORRENCIAS_CONFIG = [
    # === OCORRÊNCIAS DE AFASTAMENTO (todas dentro da pasta "Afastamentos_Atestados") ===
    {
        'tipo': 'multiplas_pasta_unica',
        'nome': 'Afastamentos/Atestados',
        'pasta': 'Afastamentos_Atestados',
        'itens': [
            {'nome': 'Afast Acid Trab <= 15 Dias', 'ocorrencia': 'Afast Acid Trab <= 15 Dias', 'justificativa': 'Afast Acid Trab <= 15 Dias', 'arquivo': 'Afast_Acid_Trab_15d'},
            {'nome': 'Afast Acid Trab > 15 Dias', 'ocorrencia': 'Afast Acid Trab > 15 Dias', 'justificativa': 'Afast Acid Trab > 15 Dias', 'arquivo': 'Afast_Acid_Trab_15d+'},
            {'nome': 'Afast Doenca <= 15 Dias', 'ocorrencia': 'Afast Doenca <= 15 Dias', 'justificativa': 'Afast Doenca <= 15 Dias', 'arquivo': 'Afast_Doenca_15d'},
            {'nome': 'Afast Doenca > 15 Dias', 'ocorrencia': 'Afast Doenca > 15 Dias', 'justificativa': 'Afast Doenca > 15 Dias', 'arquivo': 'Afast_Doenca_15d+'},
            {'nome': 'Afast Licenca Maternidade', 'ocorrencia': 'Afast Licenca Maternidade', 'justificativa': 'Afast Licenca Maternidade', 'arquivo': 'Afast_Licenca_Maternidade'},
            {'nome': 'Outros tipos de afastamento', 'ocorrencia': 'Outros tipos de afastamento', 'justificativa': 'Outros tipos de afastamento', 'arquivo': 'Outros_Afastamentos'},
        ]
    },
    
    # === OCORRÊNCIAS COM 1 JUSTIFICATIVA (justificativa = ocorrência) ===
    {'tipo': 'unica', 'nome': 'Ferias Normais', 'ocorrencia': 'Ferias Normais', 'justificativa': 'Ferias Normais', 'arquivo': 'Ferias_Normais'},
    
    # === SEM MARCAÇÕES (ambos dentro da pasta "Sem_Marcacoes") ===
    {
        'tipo': 'multiplas_pasta_unica',
        'nome': 'Sem marcações',
        'pasta': 'Sem_Marcacoes',
        'itens': [
            {'nome': 'Sem marcacao de entrada', 'ocorrencia': 'Sem marcação de entrada', 'justificativa': 'Sem marcação de entrada', 'arquivo': 'Sem_Marcacao_Entrada'},
            {'nome': 'Sem marcacao de saida', 'ocorrencia': 'Sem marcação de saída', 'justificativa': 'Sem marcação de saída', 'arquivo': 'Sem_Marcacao_Saida'},
        ]
    },
    
    # === OCORRÊNCIAS COM MÚLTIPLAS JUSTIFICATIVAS ===
    {
        'tipo': 'multiplas',
        'nome': 'Entrada em atraso',
        'ocorrencia': 'Entrada em atraso',
        'justificativas': [
            'Banco de Horas - Fechamento Semestral (Fev/Ago)',
            'Banco de Horas - Fechamento Semestral (Fev/Ago) S/D',
            'Banco de Horas - Fechamento Trimestral Fev/Maio/Ago/Nov S/D',
            'Banco de Horas Distribuição - Fechamento Trimestral (Fev/Mai/Ago/Nov) S/D',
            'Declaração de Horas',
            'Liberação Empresa - Horas',
            'Parte Ou Testemunha de Processo Judicial',
        ],
        'pasta': 'Entrada_em_Atraso'
    },
    {
        'tipo': 'multiplas',
        'nome': 'Falta',
        'ocorrencia': 'Falta',
        'justificativas': [
            'Amamentação',
            'Aniversário - Dia Livre',
            'Banco de Horas - Fechamento Trimestral Fev/Maio/Ago/Nov S/D',
            'Banco de Horas Distribuição - Fechamento Trimestral (Fev/Mai/Ago/Nov) S/D',
            'Curso de Aprendizagem',
            'Declaração de Horas',
            'Falta',
            'Folga',
            'Folga Ouro da Casa',
            'Integração',
            'Liberação da Empresa - Dia',
            'Obito de Familiar',
            'Parte Ou Testemunha de Processo Judicial',
            'Serviço Externo',
        ],
        'pasta': 'Falta'
    },
]


def gerar_excel_ocorrencia(
    df: pd.DataFrame,
    config: Dict,
    col_nome: str,
    col_cargo: str,
    col_depto: str,
    col_data_adm: str,
    col_ocorrencia: str,
    col_justificativa: str,
    col_data: str
) -> Tuple[bytes, str, int]:
    """
    Gera o Excel para uma ocorrência única.
    Retorna (bytes_do_excel, nome_do_arquivo, qtd_colaboradores)
    """
    df_detalhe, df_ranking = processar_ocorrencia(
        df,
        config['ocorrencia'],
        config['justificativa'],
        col_nome, col_cargo, col_depto, col_data_adm,
        col_ocorrencia, col_justificativa, col_data
    )
    
    if len(df_detalhe) == 0:
        return None, config.get('arquivo', '') + '.xlsx', 0
    
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        gerar_planilha_ocorrencia(df_detalhe, df_ranking, config['nome'], writer)
    
    nome_arquivo = f"{config['arquivo']}.xlsx"
    return excel_buffer.getvalue(), nome_arquivo, len(df_detalhe)


def gerar_pasta_ocorrencia(
    df: pd.DataFrame,
    config: Dict,
    col_nome: str,
    col_cargo: str,
    col_depto: str,
    col_data_adm: str,
    col_ocorrencia: str,
    col_justificativa: str,
    col_data: str
) -> Tuple[Dict[str, bytes], str, int]:
    """
    Gera Excel para cada justificativa de uma ocorrência com múltiplas justificativas.
    Retorna (dict {nome_arquivo: bytes}, nome_da_pasta, total_colaboradores)
    """
    pasta = config['pasta']
    arquivos = {}
    total_colaboradores = 0
    
    for justificativa in config['justificativas']:
        config_unica = {
            'ocorrencia': config['ocorrencia'],
            'justificativa': justificativa,
            'nome': f"{config['nome']} - {justificativa}",
            'arquivo': sanitizar_nome_arquivo(justificativa)
        }
        
        excel_bytes, nome_arquivo, qtd = gerar_excel_ocorrencia(
            df, config_unica,
            col_nome, col_cargo, col_depto, col_data_adm,
            col_ocorrencia, col_justificativa, col_data
        )
        
        if excel_bytes is not None:
            arquivos[nome_arquivo] = excel_bytes
            total_colaboradores += qtd
    
    return arquivos, pasta, total_colaboradores


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

uploaded_file = st.file_uploader(
    "1. Carregue o arquivo Excel de Ponto (XLSX)",
    type=["xlsx", "xlsm"]
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"Arquivo carregado! {len(df)} linhas, {len(df.columns)} colunas.")
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
        st.stop()
    
    if len(df.columns) < 39:
        st.error(f"O arquivo precisa ter pelo menos 39 colunas (até AM). Encontradas: {len(df.columns)}")
        st.write("Primeiras colunas encontradas:", list(df.columns[:20]))
        st.stop()
    
    col_nome = df.columns[3]
    col_cargo = df.columns[7]
    col_depto = df.columns[8]
    col_escala = df.columns[11]
    col_data_adm = df.columns[16]
    col_marcacoes = df.columns[23]
    col_ocorrencia = df.columns[25]
    col_justificativa = df.columns[27]
    col_data = df.columns[38]
    
    st.info(f"""
    **Colunas detectadas:**
    - Colaborador (D): {col_nome}
    - Cargo (H): {col_cargo}
    - Departamento (I): {col_depto}
    - Escala (L): {col_escala}
    - Data Admissão (Q): {col_data_adm}
    - Marcações (X): {col_marcacoes}
    - Ocorrência (Z): {col_ocorrencia}
    - Justificativa (AB): {col_justificativa}
    - Data (AM): {col_data}
    """)
    
    with st.expander("🔍 Preview dos dados", expanded=False):
        st.write("Amostra das colunas principais:")
        preview_cols = [col_nome, col_cargo, col_depto, col_ocorrencia, col_justificativa, col_data]
        st.dataframe(df[preview_cols].head(20), use_container_width=True)
    
    with st.expander("📊 Ocorrências disponíveis no arquivo", expanded=False):
        ocorrencias_count = df[col_ocorrencia].value_counts().reset_index()
        ocorrencias_count.columns = ['Ocorrência', 'Quantidade']
        st.dataframe(ocorrencias_count, use_container_width=True)
        
        st.write("**Valores únicos em Justificativa (AB):**")
        justificativas_count = df[col_justificativa].value_counts().reset_index()
        justificativas_count.columns = ['Justificativa', 'Quantidade']
        st.dataframe(justificativas_count, use_container_width=True)
    
    st.subheader("2. Selecione os tipos de ocorrência para gerar")
    
    # Lista todas as opções
    opcoes = []
    for config in OCORRENCIAS_CONFIG:
        if config['tipo'] == 'unica':
            opcoes.append(config['nome'])
        elif config['tipo'] == 'multiplas_pasta_unica':
            opcoes.append(f"📁 {config['nome']} ({len(config['itens'])} tipos)")
        else:
            opcoes.append(f"📁 {config['nome']} ({len(config['justificativas'])} justificativas)")
    
    selecionados = st.multiselect(
        "Selecione os tipos de ocorrência:",
        options=opcoes,
        default=opcoes,
        help="Tipos com múltiplas justificativas geram uma pasta com várias planilhas"
    )
    
    if st.button("🚀 Gerar Relatórios", type="primary", use_container_width=True):
        if not selecionados:
            st.warning("Selecione pelo menos um tipo de ocorrência.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Filtra configs selecionadas
            configs_selecionadas = []
            for config in OCORRENCIAS_CONFIG:
                if config['tipo'] == 'unica' and config['nome'] in selecionados:
                    configs_selecionadas.append(config)
                elif config['tipo'] == 'multiplas':
                    label = f"📁 {config['nome']} ({len(config['justificativas'])} justificativas)"
                    if label in selecionados:
                        configs_selecionadas.append(config)
                elif config['tipo'] == 'multiplas_pasta_unica':
                    label = f"📁 {config['nome']} ({len(config['itens'])} tipos)"
                    if label in selecionados:
                        configs_selecionadas.append(config)
            
            total = len(configs_selecionadas)
            if total == 0:
                st.warning("Nenhuma ocorrência selecionada.")
                st.stop()
            
            # Primeiro processa tudo para gerar estatísticas
            resultados_processados = []
            for i, config in enumerate(configs_selecionadas):
                status_text.text(f"Processando: {config['nome']}...")
                
                if config['tipo'] == 'unica':
                    excel_bytes, nome_arquivo, qtd = gerar_excel_ocorrencia(
                        df, config,
                        col_nome, col_cargo, col_depto, col_data_adm,
                        col_ocorrencia, col_justificativa, col_data
                    )
                    resultados_processados.append({
                        'config': config,
                        'excel_bytes': excel_bytes,
                        'nome_arquivo': nome_arquivo,
                        'qtd': qtd,
                        'tipo': 'unica'
                    })
                
                elif config['tipo'] == 'multiplas_pasta_unica':
                    # Pasta única com vários itens individuais
                    pasta = config['pasta']
                    arquivos = {}
                    total_colab = 0
                    for item in config['itens']:
                        excel_bytes, nome_arquivo, qtd = gerar_excel_ocorrencia(
                            df, item,
                            col_nome, col_cargo, col_depto, col_data_adm,
                            col_ocorrencia, col_justificativa, col_data
                        )
                        if excel_bytes is not None:
                            arquivos[nome_arquivo] = excel_bytes
                            total_colab += qtd
                    resultados_processados.append({
                        'config': config,
                        'arquivos': arquivos,
                        'pasta': pasta,
                        'total_colab': total_colab,
                        'tipo': 'multiplas'
                    })
                
                else:  # múltiplas (com várias justificativas)
                    arquivos, pasta, total_colab = gerar_pasta_ocorrencia(
                        df, config,
                        col_nome, col_cargo, col_depto, col_data_adm,
                        col_ocorrencia, col_justificativa, col_data
                    )
                    resultados_processados.append({
                        'config': config,
                        'arquivos': arquivos,
                        'pasta': pasta,
                        'total_colab': total_colab,
                        'tipo': 'multiplas'
                    })
                
                progress_bar.progress((i + 1) / (total + 1))
            
            # Gera o ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for resultado in resultados_processados:
                    status_text.text(f"Compactando: {resultado['config']['nome']}...")
                    
                    if resultado['tipo'] == 'unica':
                        if resultado['excel_bytes'] is not None:
                            zf.writestr(resultado['nome_arquivo'], resultado['excel_bytes'])
                    else:  # múltiplas
                        for nome_arquivo, excel_bytes in resultado['arquivos'].items():
                            caminho = f"{resultado['pasta']}/{nome_arquivo}"
                            zf.writestr(caminho, excel_bytes)
                    
                    progress_bar.progress((i + 1) / (total + 1) + 0.05)
            
            status_text.text("✅ Relatórios gerados com sucesso!")
            
            if len(zip_buffer.getvalue()) > 0:
                st.download_button(
                    label="📥 Baixar ZIP com todas as planilhas",
                    data=zip_buffer.getvalue(),
                    file_name="Relatorio_Ponto_Geral.zip",
                    mime="application/zip",
                    use_container_width=True
                )
            else:
                st.warning("Nenhum relatório foi gerado (nenhum registro encontrado).")
            
            # Mostra resumo
            st.subheader("📋 Resumo dos Relatórios Gerados")
            resumo_data = []
            total_com_registros = 0
            
            for resultado in resultados_processados:
                if resultado['tipo'] == 'unica':
                    tem_registros = resultado['qtd'] > 0
                    if tem_registros:
                        total_com_registros += 1
                    resumo_data.append({
                        'Ocorrência': resultado['config']['nome'],
                        'Colaboradores': resultado['qtd'],
                        'Arquivo': resultado['nome_arquivo'] if tem_registros else '—',
                        'Status': '✅' if tem_registros else '⚠️'
                    })
                else:
                    qtd_justificativas = len(resultado['arquivos'])
                    tem_registros = resultado['total_colab'] > 0
                    if tem_registros:
                        total_com_registros += 1
                    resumo_data.append({
                        'Ocorrência': f"📁 {resultado['config']['nome']}",
                        'Colaboradores': resultado['total_colab'],
                        'Arquivo': f"{qtd_justificativas} planilhas em pasta" if tem_registros else '—',
                        'Status': '✅' if tem_registros else '⚠️'
                    })
            
            df_resumo = pd.DataFrame(resumo_data)
            st.dataframe(df_resumo, use_container_width=True, hide_index=True)
            
            if total_com_registros == 0:
                st.warning("Nenhum registro encontrado para os tipos selecionados.")

else:
    st.info("👆 Carregue o arquivo Excel de ponto para começar.")
    
    st.markdown("""
    ### Estrutura esperada do arquivo
    
    O arquivo Excel deve conter as seguintes colunas (posições):
    
    | Coluna | Descrição | Exemplo |
    |--------|-----------|---------|
    | **D** | Colaborador | ADALBERTO BRANDAO DE AVILA |
    | **H** | Cargo | ASSISTENTE III |
    | **I** | Departamento | 3287 - EQUIPES ANALISTAS CDs |
    | **L** | Escala | Escala: 08:00 12:00 13:00 17:48 - 5X2 |
    | **Q** | Data Admissão | 05/12/07 |
    | **X** | Marcações | 07:58 12:00 13:00 17:48 |
    | **Z** | Ocorrência | Falta, Entrada em atraso, etc. |
    | **AB** | Justificativa | Justificativa da ocorrência |
    | **AM** | Data | 6/11/2026 |
    
    ### Estrutura do ZIP gerado
    
    **📁 Pasta `Afastamentos_Atestados/`** (6 tipos de afastamento):
    - `Afast_Acid_Trab_15d.xlsx`
    - `Afast_Acid_Trab_15d+.xlsx`
    - `Afast_Doenca_15d.xlsx`
    - `Afast_Doenca_15d+.xlsx`
    - `Afast_Licenca_Maternidade.xlsx`
    - `Outros_Afastamentos.xlsx`
    
    **📄 Arquivos na raiz** (ocorrências com 1 justificativa):
    - `Ferias_Normais.xlsx`
    
    **📁 Pasta `Sem_Marcacoes/`** (2 tipos):
    - `Sem_Marcacao_Entrada.xlsx`
    - `Sem_Marcacao_Saida.xlsx`
    
    **📁 Pasta `Entrada_em_Atraso/`** (7 justificativas):
    - `Banco_de_Horas_Fechamento_Semestral_Fev_Ago.xlsx`
    - `Declaracao_de_Horas.xlsx`
    - `Liberacao_Empresa_Horas.xlsx`
    - ...
    
    **📁 Pasta `Falta/`** (14 justificativas):
    - `Amamentacao.xlsx`
    - `Aniversario_Dia_Livre.xlsx`
    - `Falta.xlsx`
    - `Folga.xlsx`
    - `Integracao.xlsx`
    - `Servico_Externo.xlsx`
    - ...
    
    ### Layout das planilhas
    
    **Aba 1 - Detalhamento:** Cabeçalho verde escuro, linhas alternadas, destaques nas quantidades
    **Aba 2 - Ranking Ofensores:** Pódio ouro/prata/bronze, inclui Data Adm e Tempo de Serviço
    """)