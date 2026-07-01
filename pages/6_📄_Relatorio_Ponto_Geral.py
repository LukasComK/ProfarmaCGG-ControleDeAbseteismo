import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import io
import zipfile
import re
from typing import Dict, List, Tuple
from unidecode import unidecode
from funcoes_processamento_csv import determinar_turno

st.set_page_config(page_title="Relatório de Ponto Geral", layout="wide")

st.title("📄 Relatório de Ponto Geral")
st.markdown("""
Esta página gera relatórios de ocorrências por tipo, agrupando colaboradores e gerando rankings.
Os relatórios são enriquecidos com **Gestor, Supervisor e Turno** a partir do CSV de base ativos.
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


def safe_unidecode(valor):
    """Aplica unidecode com segurança tratando NaN e valores vazios."""
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    try:
        return unidecode(str(valor)).strip().lower()
    except:
        return str(valor).strip().lower()


def processar_csv_gestores(
    csv_file
) -> Tuple[Dict[str, Dict], Dict[str, str]]:
    """
    Processa o CSV de base ativos e retorna:
    - mapa_colaboradores: {nome_norm: {'gestor': str, 'jornada': str, 'turno': str}}
    - mapa_supervisores: {gestor_norm: supervisor_nome}
    """
    # Tenta ler o CSV
    df_csv = None
    erros_tentativa = []
    
    tentativas = [
        {'sep': ';', 'encoding': 'latin-1', 'skiprows': 1},
        {'sep': ';', 'encoding': 'utf-8', 'skiprows': 1},
        {'sep': ',', 'encoding': 'utf-8', 'skiprows': 0},
        {'sep': ';', 'encoding': 'latin-1', 'skiprows': 0},
    ]
    
    for params in tentativas:
        try:
            csv_file.seek(0)
            df_csv = pd.read_csv(csv_file, **params, engine='python')
            if len(df_csv.columns) >= 30:
                break
        except Exception as e:
            erros_tentativa.append(str(e))
            df_csv = None
    
    if df_csv is None or len(df_csv.columns) < 30:
        st.error(f"Não foi possível ler o CSV. Colunas encontradas: {len(df_csv.columns) if df_csv is not None else 0}")
        return {}, {}
    
    # Identifica colunas: Colaborador (D=3), Nome Gestor (Z=25), Jornada (AR=43?)
    # Usa o nome exato se possível, ou índice
    col_colaborador = df_csv.columns[3] if len(df_csv.columns) > 3 else None
    col_gestor = df_csv.columns[25] if len(df_csv.columns) > 25 else None
    
    # Jornada - tenta encontrar pelo nome, senão usa índice
    col_jornada = None
    for nome_possivel in ['Jornada', 'JORNADA', 'Codigo Jornada']:
        if nome_possivel in df_csv.columns:
            col_jornada = nome_possivel
            break
    if col_jornada is None and len(df_csv.columns) > 43:
        col_jornada = df_csv.columns[43]  # Posição próxima da Jornada
    
    if col_colaborador is None or col_gestor is None:
        st.error("CSV não tem colunas suficientes (precisa de Colaborador e Gestor).")
        return {}, {}
    
    st.success(f"CSV carregado! {len(df_csv)} colaboradores, {len(df_csv.columns)} colunas.")
    if col_jornada:
        st.caption(f"Colunas usadas: Colaborador={col_colaborador}, Gestor={col_gestor}, Jornada={col_jornada}")
    else:
        st.caption(f"Colunas usadas: Colaborador={col_colaborador}, Gestor={col_gestor}, Jornada=não encontrada")
    
    # Constrói mapa de colaboradores
    mapa_colaboradores = {}
    for idx, row in df_csv.iterrows():
        nome = str(row[col_colaborador]).strip() if pd.notna(row[col_colaborador]) else ''
        if not nome:
            continue
        nome_norm = safe_unidecode(nome)
        
        gestor = str(row[col_gestor]).strip() if pd.notna(row[col_gestor]) else ''
        
        jornada = ''
        turno = 'Indeterminado'
        if col_jornada and pd.notna(row.get(col_jornada, np.nan)):
            jornada = str(row[col_jornada]).strip()
            turno = determinar_turno(jornada)
        
        mapa_colaboradores[nome_norm] = {
            'nome_original': nome,
            'gestor': gestor,
            'jornada': jornada,
            'turno': turno
        }
    
    # Constrói mapa de supervisores (gestor do gestor)
    mapa_supervisores = {}
    for nome_norm, info in mapa_colaboradores.items():
        gestor_norm = safe_unidecode(info['gestor'])
        if gestor_norm and gestor_norm in mapa_colaboradores:
            supervisor_info = mapa_colaboradores[gestor_norm]
            mapa_supervisores[gestor_norm] = supervisor_info['gestor']
    
    return mapa_colaboradores, mapa_supervisores


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
    col_data: str,
    mapa_colaboradores: Dict[str, Dict] = None,
    mapa_supervisores: Dict[str, str] = None
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Processa um filtro de ocorrência+justificativa e retorna (detalhe, ranking)."""
    termo_occ_norm = unidecode(termo_ocorrencia).strip().lower()
    termo_just_norm = unidecode(termo_justificativa).strip().lower()
    
    occ_norm = df[col_ocorrencia].apply(safe_unidecode)
    just_norm = df[col_justificativa].apply(safe_unidecode)
    
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
    
    if mask_just.sum() == 0:
        mask_just = pd.Series([True] * len(df))
    
    df_filtrado = df[mask_occ & mask_just].copy()
    
    if len(df_filtrado) == 0:
        colunas_detalhe = [
            'Colaborador', 'Cargo', 'Departamento', 'Gestor', 'Supervisor', 'Turno',
            'Data Admissão', 'Tempo de Serviço', 'Quantidade Ocorrências', 'Datas das Ocorrências'
        ]
        df_detalhe = pd.DataFrame(columns=colunas_detalhe)
        df_ranking = pd.DataFrame(columns=['Colaborador', 'Cargo', 'Departamento', 'Gestor', 'Supervisor', 'Turno', 'Data Admissão', 'Tempo de Serviço', 'Quantidade Ocorrências'])
        return df_detalhe, df_ranking
    
    df_filtrado['Data_Formatada'] = df_filtrado[col_data].apply(formatar_data_br)
    df_filtrado['Tempo_Servico'] = df_filtrado[col_data_adm].apply(calcular_tempo_servico)
    
    # Função para buscar info do colaborador no mapa
    def get_info_colaborador(nome):
        if mapa_colaboradores is None:
            return {'gestor': '', 'turno': '', 'supervisor': ''}
        nome_norm = safe_unidecode(nome)
        info = mapa_colaboradores.get(nome_norm, {})
        gestor = info.get('gestor', '')
        turno = info.get('turno', 'Indeterminado')
        
        # Busca supervisor
        supervisor = ''
        if mapa_supervisores and gestor:
            gestor_norm = safe_unidecode(gestor)
            supervisor = mapa_supervisores.get(gestor_norm, '')
        
        return {'gestor': gestor, 'turno': turno, 'supervisor': supervisor}
    
    grupos = df_filtrado.groupby(col_nome)
    
    dados_detalhe = []
    for nome, grupo in grupos:
        primeiro = grupo.iloc[0]
        datas = sorted(grupo['Data_Formatada'].dropna().unique().tolist())
        datas_str = ', '.join(datas) if datas else ''
        
        info = get_info_colaborador(nome)
        
        dados_detalhe.append({
            'Colaborador': nome,
            'Cargo': primeiro[col_cargo] if col_cargo in grupo.columns else '',
            'Departamento': primeiro[col_depto] if col_depto in grupo.columns else '',
            'Gestor': info['gestor'],
            'Supervisor': info['supervisor'],
            'Turno': info['turno'],
            'Data Admissão': primeiro[col_data_adm] if col_data_adm in grupo.columns else '',
            'Tempo de Serviço': primeiro['Tempo_Servico'],
            'Quantidade Ocorrências': len(grupo),
            'Datas das Ocorrências': datas_str
        })
    
    df_detalhe = pd.DataFrame(dados_detalhe)
    df_detalhe = df_detalhe.sort_values('Quantidade Ocorrências', ascending=False).reset_index(drop=True)
    
    # Ranking
    cols_ranking = ['Colaborador', 'Cargo', 'Departamento', 'Gestor', 'Supervisor', 'Turno', 'Data Admissão', 'Tempo de Serviço', 'Quantidade Ocorrências']
    df_ranking = df_detalhe[cols_ranking].copy()
    df_ranking = df_ranking.sort_values('Quantidade Ocorrências', ascending=False).reset_index(drop=True)
    df_ranking['Posição'] = range(1, len(df_ranking) + 1)
    df_ranking = df_ranking[['Posição'] + cols_ranking]
    
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
    
    # Larguras dinâmicas baseadas nas colunas
    col_widths = {
        'Colaborador': 36, 'Cargo': 20, 'Departamento': 28, 'Gestor': 30,
        'Supervisor': 30, 'Turno': 14, 'Data Admissão': 16, 'Tempo de Serviço': 16,
        'Quantidade Ocorrências': 20, 'Datas das Ocorrências': 50
    }
    for col_num, col_name in enumerate(df_detalhe.columns):
        width = col_widths.get(col_name, 18)
        ws.set_column(col_num, col_num, width)
    
    # Índice da coluna Quantidade Ocorrências
    qtd_col_idx = list(df_detalhe.columns).index('Quantidade Ocorrências') if 'Quantidade Ocorrências' in df_detalhe.columns else -1
    
    for row_num in range(1, len(df_detalhe) + 1):
        is_alt = row_num % 2 == 0
        for col_num in range(len(df_detalhe.columns)):
            valor = df_detalhe.iloc[row_num - 1, col_num]
            if col_num == qtd_col_idx:
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
    sheet_name_rank = 'Ofensores'
    if sheet_name_rank in writer.sheets:
        sheet_name_rank = 'Ofensores 2'
    df_ranking.to_excel(writer, sheet_name=sheet_name_rank[:31], index=False, startrow=0)
    ws_rank = writer.sheets[sheet_name_rank[:31]]
    
    for col_num, value in enumerate(df_ranking.columns.values):
        ws_rank.write(0, col_num, value, header_ranking_fmt)
    
    # Larguras ranking
    col_widths_rank = {
        'Posição': 8, 'Colaborador': 36, 'Cargo': 20, 'Departamento': 28,
        'Gestor': 30, 'Supervisor': 30, 'Turno': 14, 'Data Admissão': 16,
        'Tempo de Serviço': 16, 'Quantidade Ocorrências': 20
    }
    qtd_col_idx_rank = list(df_ranking.columns).index('Quantidade Ocorrências') if 'Quantidade Ocorrências' in df_ranking.columns else -1
    posicao_col_idx = 0
    
    for col_num, col_name in enumerate(df_ranking.columns):
        width = col_widths_rank.get(col_name, 18)
        ws_rank.set_column(col_num, col_num, width)
    
    for row_num in range(1, len(df_ranking) + 1):
        is_alt = row_num % 2 == 0
        posicao = df_ranking.iloc[row_num - 1, posicao_col_idx]
        for col_num in range(len(df_ranking.columns)):
            valor = df_ranking.iloc[row_num - 1, col_num]
            if col_num == posicao_col_idx and posicao in posicoes_format:
                fmt = posicoes_format[posicao]
            elif col_num == qtd_col_idx_rank:
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
    nome = unidecode(nome)
    nome = re.sub(r'[<>:"/\\|?*]', '', nome)
    nome = re.sub(r'\s+', '_', nome.strip())
    nome = re.sub(r'_+', '_', nome)
    return nome[:80]


# ============================================================
# CONFIGURAÇÃO DE TODAS AS OCORRÊNCIAS
# ============================================================
OCORRENCIAS_CONFIG = [
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
    {'tipo': 'unica', 'nome': 'Ferias Normais', 'ocorrencia': 'Ferias Normais', 'justificativa': 'Ferias Normais', 'arquivo': 'Ferias_Normais'},
    {
        'tipo': 'multiplas_pasta_unica',
        'nome': 'Sem marcações',
        'pasta': 'Sem_Marcacoes',
        'itens': [
            {'nome': 'Sem marcacao de entrada', 'ocorrencia': 'Sem marcação de entrada', 'justificativa': 'Sem marcação de entrada', 'arquivo': 'Sem_Marcacao_Entrada'},
            {'nome': 'Sem marcacao de saida', 'ocorrencia': 'Sem marcação de saída', 'justificativa': 'Sem marcação de saída', 'arquivo': 'Sem_Marcacao_Saida'},
        ]
    },
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
            'Amamentação', 'Aniversário - Dia Livre',
            'Banco de Horas - Fechamento Trimestral Fev/Maio/Ago/Nov S/D',
            'Banco de Horas Distribuição - Fechamento Trimestral (Fev/Mai/Ago/Nov) S/D',
            'Curso de Aprendizagem', 'Declaração de Horas', 'Falta', 'Folga',
            'Folga Ouro da Casa', 'Integração', 'Liberação da Empresa - Dia',
            'Obito de Familiar', 'Parte Ou Testemunha de Processo Judicial', 'Serviço Externo',
        ],
        'pasta': 'Falta'
    },
]


def gerar_excel_ocorrencia(
    df: pd.DataFrame,
    config: Dict,
    col_nome: str, col_cargo: str, col_depto: str, col_data_adm: str,
    col_ocorrencia: str, col_justificativa: str, col_data: str,
    mapa_colaboradores: Dict = None, mapa_supervisores: Dict = None
) -> Tuple[bytes, str, int]:
    df_detalhe, df_ranking = processar_ocorrencia(
        df, config['ocorrencia'], config['justificativa'],
        col_nome, col_cargo, col_depto, col_data_adm,
        col_ocorrencia, col_justificativa, col_data,
        mapa_colaboradores, mapa_supervisores
    )
    
    if len(df_detalhe) == 0:
        return None, config.get('arquivo', '') + '.xlsx', 0
    
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        gerar_planilha_ocorrencia(df_detalhe, df_ranking, config['nome'], writer)
    
    nome_arquivo = f"{config['arquivo']}.xlsx"
    return excel_buffer.getvalue(), nome_arquivo, len(df_detalhe)


def gerar_pasta_ocorrencia(
    df: pd.DataFrame, config: Dict,
    col_nome: str, col_cargo: str, col_depto: str, col_data_adm: str,
    col_ocorrencia: str, col_justificativa: str, col_data: str,
    mapa_colaboradores: Dict = None, mapa_supervisores: Dict = None
) -> Tuple[Dict[str, bytes], str, int]:
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
            col_ocorrencia, col_justificativa, col_data,
            mapa_colaboradores, mapa_supervisores
        )
        if excel_bytes is not None:
            arquivos[nome_arquivo] = excel_bytes
            total_colaboradores += qtd
    
    return arquivos, pasta, total_colaboradores


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

st.subheader("1. Carregue os arquivos")

col1, col2 = st.columns(2)

with col1:
    uploaded_file = st.file_uploader(
        "Arquivo Excel de Ponto (XLSX)",
        type=["xlsx", "xlsm"],
        help="Arquivo com as colunas D, H, I, L, Q, X, Z, AB, AM"
    )

with col2:
    uploaded_csv = st.file_uploader(
        "Arquivo CSV de Base Ativos",
        type=["csv"],
        help="Arquivo CSV com Colaborador (D), Nome Gestor (Z) e Jornada (~AR)"
    )

# Processa CSV se carregado
mapa_colaboradores = {}
mapa_supervisores = {}

if uploaded_csv is not None:
    with st.spinner("Processando CSV de gestores..."):
        mapa_colaboradores, mapa_supervisores = processar_csv_gestores(uploaded_csv)
    
    if mapa_colaboradores:
        st.success(f"✅ {len(mapa_colaboradores)} colaboradores mapeados no CSV")

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"Arquivo de ponto carregado! {len(df)} linhas, {len(df.columns)} colunas.")
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
    **Colunas detectadas no Ponto:**
    - Colaborador (D): {col_nome}
    - Cargo (H): {col_cargo}
    - Departamento (I): {col_depto}
    - Data Admissão (Q): {col_data_adm}
    - Ocorrência (Z): {col_ocorrencia}
    - Justificativa (AB): {col_justificativa}
    - Data (AM): {col_data}
    """)
    
    with st.expander("🔍 Preview dos dados", expanded=False):
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
            
            resultados_processados = []
            for i, config in enumerate(configs_selecionadas):
                status_text.text(f"Processando: {config['nome']}...")
                
                if config['tipo'] == 'unica':
                    excel_bytes, nome_arquivo, qtd = gerar_excel_ocorrencia(
                        df, config,
                        col_nome, col_cargo, col_depto, col_data_adm,
                        col_ocorrencia, col_justificativa, col_data,
                        mapa_colaboradores, mapa_supervisores
                    )
                    resultados_processados.append({
                        'config': config, 'excel_bytes': excel_bytes,
                        'nome_arquivo': nome_arquivo, 'qtd': qtd, 'tipo': 'unica'
                    })
                
                elif config['tipo'] == 'multiplas_pasta_unica':
                    pasta = config['pasta']
                    arquivos = {}
                    total_colab = 0
                    for item in config['itens']:
                        excel_bytes, nome_arquivo, qtd = gerar_excel_ocorrencia(
                            df, item,
                            col_nome, col_cargo, col_depto, col_data_adm,
                            col_ocorrencia, col_justificativa, col_data,
                            mapa_colaboradores, mapa_supervisores
                        )
                        if excel_bytes is not None:
                            arquivos[nome_arquivo] = excel_bytes
                            total_colab += qtd
                    resultados_processados.append({
                        'config': config, 'arquivos': arquivos,
                        'pasta': pasta, 'total_colab': total_colab, 'tipo': 'multiplas'
                    })
                
                else:
                    arquivos, pasta, total_colab = gerar_pasta_ocorrencia(
                        df, config,
                        col_nome, col_cargo, col_depto, col_data_adm,
                        col_ocorrencia, col_justificativa, col_data,
                        mapa_colaboradores, mapa_supervisores
                    )
                    resultados_processados.append({
                        'config': config, 'arquivos': arquivos,
                        'pasta': pasta, 'total_colab': total_colab, 'tipo': 'multiplas'
                    })
                
                progress_bar.progress((i + 1) / (total + 1))
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for resultado in resultados_processados:
                    status_text.text(f"Compactando: {resultado['config']['nome']}...")
                    if resultado['tipo'] == 'unica':
                        if resultado['excel_bytes'] is not None:
                            zf.writestr(resultado['nome_arquivo'], resultado['excel_bytes'])
                    else:
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
                    qtd_arquivos = len(resultado['arquivos'])
                    tem_registros = resultado['total_colab'] > 0
                    if tem_registros:
                        total_com_registros += 1
                    resumo_data.append({
                        'Ocorrência': f"📁 {resultado['config']['nome']}",
                        'Colaboradores': resultado['total_colab'],
                        'Arquivo': f"{qtd_arquivos} planilhas" if tem_registros else '—',
                        'Status': '✅' if tem_registros else '⚠️'
                    })
            
            df_resumo = pd.DataFrame(resumo_data)
            st.dataframe(df_resumo, use_container_width=True, hide_index=True)
            
            if total_com_registros == 0:
                st.warning("Nenhum registro encontrado para os tipos selecionados.")

else:
    st.info("👆 Carregue os arquivos para começar.")
    
    st.markdown("""
    ### Arquivos necessários
    
    **1. Excel de Ponto (XLSX)** - Controle de ponto com as colunas:
    - D=Colaborador, H=Cargo, I=Departamento, L=Escala, Q=Data Admissão
    - X=Maracões, Z=Ocorrência, AB=Justificativa, AM=Data
    
    **2. CSV de Base Ativos** (opcional, mas recomendado) - Para enriquecer com:
    - **Gestor** (coluna "Nome Gestor")
    - **Supervisor** (gestor do gestor)
    - **Turno** (calculado pela jornada)
    
    ### Colunas adicionadas nas planilhas
    
    Todas as planilhas geradas agora incluem:
    | Coluna | Origem |
    |--------|--------|
    | Gestor | CSV → "Nome Gestor" |
    | Supervisor | CSV → Gestor do Gestor |
    | Turno | CSV → Jornada (TURNO 1/2/3) |
    """)