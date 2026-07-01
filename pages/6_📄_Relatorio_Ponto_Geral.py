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
    if pd.isna(valor) or str(valor).strip() == '':
        return ''
    try:
        return unidecode(str(valor)).strip().lower()
    except:
        return str(valor).strip().lower()


def processar_csv_gestores(csv_file) -> Tuple[Dict[str, Dict], Dict[str, str]]:
    df_csv = None
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
        except:
            df_csv = None
    if df_csv is None or len(df_csv.columns) < 30:
        st.error(f"Não foi possível ler o CSV. Colunas: {len(df_csv.columns) if df_csv is not None else 0}")
        return {}, {}
    col_colaborador = df_csv.columns[3] if len(df_csv.columns) > 3 else None
    col_gestor = df_csv.columns[25] if len(df_csv.columns) > 25 else None
    col_jornada = None
    for nome in ['Jornada', 'JORNADA', 'Codigo Jornada']:
        if nome in df_csv.columns:
            col_jornada = nome
            break
    if col_jornada is None and len(df_csv.columns) > 43:
        col_jornada = df_csv.columns[43]
    if col_colaborador is None or col_gestor is None:
        st.error("CSV não tem colunas suficientes.")
        return {}, {}
    st.success(f"CSV carregado! {len(df_csv)} colaboradores.")
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
        mapa_colaboradores[nome_norm] = {'nome_original': nome, 'gestor': gestor, 'jornada': jornada, 'turno': turno}
    mapa_supervisores = {}
    for nome_norm, info in mapa_colaboradores.items():
        gestor_norm = safe_unidecode(info['gestor'])
        if gestor_norm and gestor_norm in mapa_colaboradores:
            mapa_supervisores[gestor_norm] = mapa_colaboradores[gestor_norm]['gestor']
    return mapa_colaboradores, mapa_supervisores


def processar_ocorrencia(
    df: pd.DataFrame,
    termo_ocorrencia: str, termo_justificativa: str,
    col_nome: str, col_cargo: str, col_depto: str, col_data_adm: str,
    col_ocorrencia: str, col_justificativa: str, col_data: str,
    mapa_colaboradores: Dict[str, Dict] = None,
    mapa_supervisores: Dict[str, str] = None
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    termo_occ_norm = unidecode(termo_ocorrencia).strip().lower()
    termo_just_norm = unidecode(termo_justificativa).strip().lower()
    
    import re as regex_module
    termo_occ_escape = regex_module.escape(termo_occ_norm)
    termo_just_escape = regex_module.escape(termo_just_norm)
    
    occ_norm = df[col_ocorrencia].apply(safe_unidecode)
    just_norm = df[col_justificativa].apply(safe_unidecode)
    
    # Filtra ocorrencia
    mask_occ = occ_norm == termo_occ_norm
    if mask_occ.sum() == 0:
        mask_occ = occ_norm.str.contains(termo_occ_escape, na=False, regex=True)
    if mask_occ.sum() == 0:
        palavras = termo_occ_norm.split()
        for palavra in reversed(palavras):
            if len(palavra) > 3:
                mask_occ = occ_norm.str.contains(regex_module.escape(palavra), na=False, regex=True)
                if mask_occ.sum() > 0:
                    break
    
    # Filtra justificativa
    mask_just = just_norm == termo_just_norm
    if mask_just.sum() == 0:
        mask_just = just_norm.str.contains(termo_just_escape, na=False, regex=True)
    if mask_just.sum() == 0:
        palavras = termo_just_norm.split()
        for palavra in reversed(palavras):
            if len(palavra) > 3:
                mask_just = just_norm.str.contains(regex_module.escape(palavra), na=False, regex=True)
                if mask_just.sum() > 0:
                    break
    if mask_just.sum() == 0:
        mask_just = pd.Series([True] * len(df))
    
    df_filtrado = df[mask_occ & mask_just].copy()
    
    if len(df_filtrado) == 0:
        colunas_detalhe = ['Colaborador', 'Cargo', 'Departamento', 'Gestor', 'Supervisor', 'Turno', 'Data Admissão', 'Tempo de Serviço', 'Quantidade Ocorrências', 'Datas das Ocorrências']
        return pd.DataFrame(columns=colunas_detalhe), pd.DataFrame(columns=['Posição', 'Colaborador', 'Cargo', 'Departamento', 'Gestor', 'Supervisor', 'Turno', 'Data Admissão', 'Tempo de Serviço', 'Quantidade Ocorrências'])
    
    # Prepara dados
    df_filtrado['Data_Formatada'] = df_filtrado[col_data].apply(formatar_data_br)
    df_filtrado['Tempo_Servico'] = df_filtrado[col_data_adm].apply(calcular_tempo_servico)
    
    def get_info_colaborador(nome):
        if mapa_colaboradores is None:
            return {'gestor': '', 'turno': '', 'supervisor': ''}
        nome_norm = safe_unidecode(nome)
        info = mapa_colaboradores.get(nome_norm, {})
        gestor = info.get('gestor', '')
        turno = info.get('turno', 'Indeterminado')
        supervisor = ''
        if mapa_supervisores and gestor:
            gestor_norm = safe_unidecode(gestor)
            supervisor = mapa_supervisores.get(gestor_norm, '')
        return {'gestor': gestor, 'turno': turno, 'supervisor': supervisor}
    
    df_filtrado['Gestor'] = df_filtrado[col_nome].apply(lambda x: get_info_colaborador(x)['gestor'])
    df_filtrado['Supervisor'] = df_filtrado[col_nome].apply(lambda x: get_info_colaborador(x)['supervisor'])
    df_filtrado['Turno'] = df_filtrado[col_nome].apply(lambda x: get_info_colaborador(x)['turno'])
    
    # Cria coluna de valor ANTES de qualquer rename - usa o índice para acessar
    # Isso é mais seguro do que usar col_justificativa que pode ser um nome estranho
    idx_justificativa = 27  # AB
    if len(df_filtrado.columns) > idx_justificativa:
        col_name_just = df_filtrado.columns[idx_justificativa]
        df_filtrado['_VALOR'] = df_filtrado[col_name_just].fillna('').astype(str)
    else:
        df_filtrado['_VALOR'] = df_filtrado[col_justificativa].fillna('').astype(str)
    df_filtrado.loc[df_filtrado['_VALOR'].str.lower().isin(['nan', 'nat', '']), '_VALOR'] = ''
    
    # DEBUG: guarda info para mostrar no Streamlit
    debug_valores = df_filtrado['_VALOR'].value_counts().head(10).to_dict()
    debug_qtde = len(df_filtrado)
    st.session_state['_debug_qtde'] = debug_qtde
    st.session_state['_debug_valores'] = debug_valores
    
    # Renomeia colunas para padrao
    rename_map = {}
    for col_orig, col_novo in [(col_nome, 'Colaborador'), (col_cargo, 'Cargo'), (col_depto, 'Departamento'), (col_data_adm, 'Data Admissão')]:
        if col_orig != col_novo and col_orig in df_filtrado.columns:
            rename_map[col_orig] = col_novo
    if 'Tempo_Servico' in df_filtrado.columns:
        rename_map['Tempo_Servico'] = 'Tempo de Serviço'
    if rename_map:
        df_filtrado = df_filtrado.rename(columns=rename_map)
    
    # ========================================================
    # CONSTRUCAO DIRETA DO DETALHAMENTO (MAIS ROBUSTA)
    # ========================================================
    # Agrupa por Colaborador e Data, concatenando os valores
    df_agg = df_filtrado.groupby(['Colaborador', 'Data_Formatada'], as_index=False)['_VALOR'].agg(
        lambda x: ' | '.join(sorted(set(v for v in x if v.strip() != '')))
    )
    
    # Pega dados fixos de cada colaborador (primeira linha)
    cols_fixas = ['Colaborador', 'Cargo', 'Departamento', 'Gestor', 'Supervisor', 'Turno', 'Data Admissão', 'Tempo de Serviço']
    df_fixas = df_filtrado[cols_fixas].drop_duplicates(subset=['Colaborador']).set_index('Colaborador')
    
    # Cria a tabela pivotada manualmente para ter controle total
    # 1. Lista todos os colaboradores unicos
    colaboradores = df_agg['Colaborador'].unique()
    # 2. Lista todas as datas unicas
    todas_datas = sorted(df_agg['Data_Formatada'].unique().tolist())
    try:
        todas_datas.sort(key=lambda x: datetime.strptime(x, '%d/%m/%Y') if x else datetime.min)
    except:
        pass
    # 3. Dicionario para acesso rapido: (colab, data) -> valor
    lookup = {}
    qtd_por_colab = {}
    datas_por_colab = {}
    for _, row in df_agg.iterrows():
        colab = str(row['Colaborador']).strip()
        data = str(row['Data_Formatada']).strip()
        valor = str(row['_VALOR']).strip()
        lookup[(colab, data)] = valor
        if colab not in qtd_por_colab:
            qtd_por_colab[colab] = 0
            datas_por_colab[colab] = []
        qtd_por_colab[colab] += 1
        datas_por_colab[colab].append(data)
    
    # 4. Monta o DataFrame final
    linhas = []
    for colab in colaboradores:
        linha = {}
        # Dados fixos
        if colab in df_fixas.index:
            for c in cols_fixas:
                if c != 'Colaborador':
                    linha[c] = df_fixas.loc[colab, c]
        # Colaborador
        linha['Colaborador'] = colab
        # Quantidade e datas concatenadas
        linha['Quantidade Ocorrências'] = qtd_por_colab.get(colab, 0)
        linha['Datas das Ocorrências'] = ', '.join(sorted(datas_por_colab.get(colab, [])))
        # Cada data vira uma coluna
        for data in todas_datas:
            valor = lookup.get((colab, data), '')
            # Se tiver valor, mostra o texto da justificativa
            linha[data] = valor
        linhas.append(linha)
    
    df_detalhe = pd.DataFrame(linhas)
    df_detalhe = df_detalhe.sort_values('Quantidade Ocorrências', ascending=False).reset_index(drop=True)
    
    # Colunas: fixas + qtd + datas concatenadas + colunas de data
    cols_finais = cols_fixas + ['Quantidade Ocorrências', 'Datas das Ocorrências'] + todas_datas
    cols_existentes = [c for c in cols_finais if c in df_detalhe.columns]
    df_detalhe = df_detalhe[cols_existentes]
    
    # Ranking
    cols_ranking = cols_fixas + ['Quantidade Ocorrências', 'Datas das Ocorrências']
    cols_ranking_existentes = [c for c in cols_ranking if c in df_detalhe.columns]
    df_ranking = df_detalhe[cols_ranking_existentes].copy()
    df_ranking = df_ranking.sort_values('Quantidade Ocorrências', ascending=False).reset_index(drop=True)
    df_ranking.insert(0, 'Posição', range(1, len(df_ranking) + 1))
    
    return df_detalhe, df_ranking


def gerar_planilha_ocorrencia(df_detalhe: pd.DataFrame, df_ranking: pd.DataFrame, nome_aba_detalhe: str, writer: pd.ExcelWriter):
    workbook = writer.book
    header_fmt = workbook.add_format({'bold': True, 'font_size': 11, 'font_color': COR_BRANCO, 'bg_color': COR_VERDE_ESCURO, 'border': 1, 'border_color': COR_VERDE_ESCURO, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
    header_ranking_fmt = workbook.add_format({'bold': True, 'font_size': 11, 'font_color': COR_BRANCO, 'bg_color': COR_VERDE_MEDIO, 'border': 1, 'border_color': COR_VERDE_MEDIO, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
    cell_fmt = workbook.add_format({'font_size': 10, 'border': 1, 'border_color': '#BFBFBF', 'text_wrap': True, 'valign': 'vcenter'})
    cell_alt_fmt = workbook.add_format({'font_size': 10, 'border': 1, 'border_color': '#BFBFBF', 'bg_color': COR_CINZA_CLARO, 'text_wrap': True, 'valign': 'vcenter'})
    cell_qtd_fmt = workbook.add_format({'font_size': 11, 'bold': True, 'font_color': COR_PRETO, 'border': 1, 'border_color': '#BFBFBF', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
    cell_qtd_alt_fmt = workbook.add_format({'font_size': 11, 'bold': True, 'font_color': COR_PRETO, 'border': 1, 'border_color': '#BFBFBF', 'bg_color': COR_CINZA_CLARO, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
    cell_alta_fmt = workbook.add_format({'font_size': 11, 'bold': True, 'font_color': COR_VERMELHO, 'border': 1, 'border_color': '#BFBFBF', 'bg_color': '#FFF0F0', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
    cell_media_fmt = workbook.add_format({'font_size': 11, 'bold': True, 'font_color': COR_LARANJA, 'border': 1, 'border_color': '#BFBFBF', 'bg_color': '#FFF8E7', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
    posicoes_format = {
        1: workbook.add_format({'font_size': 12, 'bold': True, 'font_color': COR_BRANCO, 'bg_color': '#D4A017', 'border': 1, 'border_color': '#BFBFBF', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'}),
        2: workbook.add_format({'font_size': 12, 'bold': True, 'font_color': '#333333', 'bg_color': '#C0C0C0', 'border': 1, 'border_color': '#BFBFBF', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'}),
        3: workbook.add_format({'font_size': 12, 'bold': True, 'font_color': COR_BRANCO, 'bg_color': '#CD7F32', 'border': 1, 'border_color': '#BFBFBF', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
    }
    col_widths = {'Colaborador': 36, 'Cargo': 20, 'Departamento': 28, 'Gestor': 30, 'Supervisor': 30, 'Turno': 14, 'Data Admissão': 16, 'Tempo de Serviço': 16, 'Quantidade Ocorrências': 20, 'Datas das Ocorrências': 50}
    col_widths_rank = {'Posição': 8, 'Colaborador': 36, 'Cargo': 20, 'Departamento': 28, 'Gestor': 30, 'Supervisor': 30, 'Turno': 14, 'Data Admissão': 16, 'Tempo de Serviço': 16, 'Quantidade Ocorrências': 20}
    
    sheet_name = nome_aba_detalhe[:31]
    df_detalhe.to_excel(writer, sheet_name=sheet_name, index=False, startrow=0)
    ws = writer.sheets[sheet_name]
    for col_num, value in enumerate(df_detalhe.columns.values):
        ws.write(0, col_num, value, header_fmt)
    for col_num, col_name in enumerate(df_detalhe.columns):
        ws.set_column(col_num, col_num, col_widths.get(col_name, 18))
    qtd_col_idx = list(df_detalhe.columns).index('Quantidade Ocorrências') if 'Quantidade Ocorrências' in df_detalhe.columns else -1
    for row_num in range(1, len(df_detalhe) + 1):
        is_alt = row_num % 2 == 0
        for col_num in range(len(df_detalhe.columns)):
            valor = df_detalhe.iloc[row_num - 1, col_num]
            if col_num == qtd_col_idx:
                qtd = valor
                if qtd >= 5: fmt = cell_alta_fmt
                elif qtd >= 3: fmt = cell_media_fmt
                elif is_alt: fmt = cell_qtd_alt_fmt
                else: fmt = cell_qtd_fmt
            elif is_alt: fmt = cell_alt_fmt
            else: fmt = cell_fmt
            ws.write(row_num, col_num, valor, fmt)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(df_detalhe), len(df_detalhe.columns) - 1)
    
    sheet_name_rank = 'Ofensores'
    if sheet_name_rank in writer.sheets:
        sheet_name_rank = 'Ofensores 2'
    df_ranking.to_excel(writer, sheet_name=sheet_name_rank[:31], index=False, startrow=0)
    ws_rank = writer.sheets[sheet_name_rank[:31]]
    for col_num, value in enumerate(df_ranking.columns.values):
        ws_rank.write(0, col_num, value, header_ranking_fmt)
    for col_num, col_name in enumerate(df_ranking.columns):
        ws_rank.set_column(col_num, col_num, col_widths_rank.get(col_name, 18))
    qtd_col_idx_rank = list(df_ranking.columns).index('Quantidade Ocorrências') if 'Quantidade Ocorrências' in df_ranking.columns else -1
    for row_num in range(1, len(df_ranking) + 1):
        is_alt = row_num % 2 == 0
        posicao = df_ranking.iloc[row_num - 1, 0]
        for col_num in range(len(df_ranking.columns)):
            valor = df_ranking.iloc[row_num - 1, col_num]
            if col_num == 0 and posicao in posicoes_format: fmt = posicoes_format[posicao]
            elif col_num == qtd_col_idx_rank:
                qtd = valor
                if qtd >= 5: fmt = cell_alta_fmt
                elif qtd >= 3: fmt = cell_media_fmt
                elif is_alt: fmt = cell_qtd_alt_fmt
                else: fmt = cell_qtd_fmt
            elif is_alt: fmt = cell_alt_fmt
            else: fmt = cell_fmt
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


OCORRENCIAS_CONFIG = [
    {'tipo': 'multiplas_pasta_unica', 'nome': 'Afastamentos/Atestados', 'pasta': 'Afastamentos_Atestados',
     'itens': [
         {'nome': 'Afast Acid Trab <= 15 Dias', 'ocorrencia': 'Afast Acid Trab <= 15 Dias', 'justificativa': 'Afast Acid Trab <= 15 Dias', 'arquivo': 'Afast_Acid_Trab_15d'},
         {'nome': 'Afast Acid Trab > 15 Dias', 'ocorrencia': 'Afast Acid Trab > 15 Dias', 'justificativa': 'Afast Acid Trab > 15 Dias', 'arquivo': 'Afast_Acid_Trab_15d+'},
         {'nome': 'Afast Doenca <= 15 Dias', 'ocorrencia': 'Afast Doenca <= 15 Dias', 'justificativa': 'Afast Doenca <= 15 Dias', 'arquivo': 'Afast_Doenca_15d'},
         {'nome': 'Afast Doenca > 15 Dias', 'ocorrencia': 'Afast Doenca > 15 Dias', 'justificativa': 'Afast Doenca > 15 Dias', 'arquivo': 'Afast_Doenca_15d+'},
         {'nome': 'Afast Licenca Maternidade', 'ocorrencia': 'Afast Licenca Maternidade', 'justificativa': 'Afast Licenca Maternidade', 'arquivo': 'Afast_Licenca_Maternidade'},
         {'nome': 'Outros tipos de afastamento', 'ocorrencia': 'Outros tipos de afastamento', 'justificativa': 'Outros tipos de afastamento', 'arquivo': 'Outros_Afastamentos'}
     ]},
    {'tipo': 'unica', 'nome': 'Ferias Normais', 'ocorrencia': 'Ferias Normais', 'justificativa': 'Ferias Normais', 'arquivo': 'Ferias_Normais'},
    {'tipo': 'multiplas_pasta_unica', 'nome': 'Sem marcações', 'pasta': 'Sem_Marcacoes',
     'itens': [
         {'nome': 'Sem marcacao de entrada', 'ocorrencia': 'Sem marcação de entrada', 'justificativa': 'Sem marcação de entrada', 'arquivo': 'Sem_Marcacao_Entrada'},
         {'nome': 'Sem marcacao de saida', 'ocorrencia': 'Sem marcação de saída', 'justificativa': 'Sem marcação de saída', 'arquivo': 'Sem_Marcacao_Saida'}
     ]},
    {'tipo': 'multiplas', 'nome': 'Entrada em atraso', 'ocorrencia': 'Entrada em atraso',
     'justificativas': ['Banco de Horas - Fechamento Semestral (Fev/Ago)', 'Banco de Horas - Fechamento Semestral (Fev/Ago) S/D', 'Banco de Horas - Fechamento Trimestral Fev/Maio/Ago/Nov S/D', 'Banco de Horas Distribuição - Fechamento Trimestral (Fev/Mai/Ago/Nov) S/D', 'Declaração de Horas', 'Liberação Empresa - Horas', 'Parte Ou Testemunha de Processo Judicial'],
     'pasta': 'Entrada_em_Atraso'},
    {'tipo': 'multiplas', 'nome': 'Falta', 'ocorrencia': 'Falta',
     'justificativas': ['Amamentação', 'Aniversário - Dia Livre', 'Banco de Horas - Fechamento Trimestral Fev/Maio/Ago/Nov S/D', 'Banco de Horas Distribuição - Fechamento Trimestral (Fev/Mai/Ago/Nov) S/D', 'Curso de Aprendizagem', 'Declaração de Horas', 'Falta', 'Folga', 'Folga Ouro da Casa', 'Integração', 'Liberação da Empresa - Dia', 'Obito de Familiar', 'Parte Ou Testemunha de Processo Judicial', 'Serviço Externo'],
     'pasta': 'Falta'}
]


def gerar_excel_ocorrencia(df: pd.DataFrame, config: Dict, col_nome: str, col_cargo: str, col_depto: str, col_data_adm: str, col_ocorrencia: str, col_justificativa: str, col_data: str, mapa_colaboradores: Dict = None, mapa_supervisores: Dict = None) -> Tuple[bytes, str, int]:
    df_detalhe, df_ranking = processar_ocorrencia(df, config['ocorrencia'], config['justificativa'], col_nome, col_cargo, col_depto, col_data_adm, col_ocorrencia, col_justificativa, col_data, mapa_colaboradores, mapa_supervisores)
    if len(df_detalhe) == 0:
        return None, config.get('arquivo', '') + '.xlsx', 0
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        gerar_planilha_ocorrencia(df_detalhe, df_ranking, config['nome'], writer)
    return excel_buffer.getvalue(), f"{config['arquivo']}.xlsx", len(df_detalhe)


def gerar_pasta_ocorrencia(df: pd.DataFrame, config: Dict, col_nome: str, col_cargo: str, col_depto: str, col_data_adm: str, col_ocorrencia: str, col_justificativa: str, col_data: str, mapa_colaboradores: Dict = None, mapa_supervisores: Dict = None) -> Tuple[Dict[str, bytes], str, int]:
    pasta = config['pasta']
    arquivos = {}
    total_colaboradores = 0
    for justificativa in config['justificativas']:
        config_unica = {'ocorrencia': config['ocorrencia'], 'justificativa': justificativa, 'nome': f"{config['nome']} - {justificativa}", 'arquivo': sanitizar_nome_arquivo(justificativa)}
        excel_bytes, nome_arquivo, qtd = gerar_excel_ocorrencia(df, config_unica, col_nome, col_cargo, col_depto, col_data_adm, col_ocorrencia, col_justificativa, col_data, mapa_colaboradores, mapa_supervisores)
        if excel_bytes is not None:
            arquivos[nome_arquivo] = excel_bytes
            total_colaboradores += qtd
    return arquivos, pasta, total_colaboradores


# ============================================================
# INTERFACE
# ============================================================

st.subheader("1. Carregue os arquivos")
col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("Arquivo Excel de Ponto (XLSX)", type=["xlsx", "xlsm"])
with col2:
    uploaded_csv = st.file_uploader("Arquivo CSV de Base Ativos", type=["csv"])

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
        st.success(f"Arquivo carregado! {len(df)} linhas, {len(df.columns)} colunas.")
    except Exception as e:
        st.error(f"Erro: {e}")
        st.stop()
    
    if len(df.columns) < 39:
        st.error(f"Precisa de 39 colunas. Encontradas: {len(df.columns)}")
        st.stop()
    
    col_nome = df.columns[3]; col_cargo = df.columns[7]; col_depto = df.columns[8]
    col_escala = df.columns[11]; col_data_adm = df.columns[16]; col_marcacoes = df.columns[23]
    col_ocorrencia = df.columns[25]; col_justificativa = df.columns[27]; col_data = df.columns[38]
    
    st.info(f"D={col_nome} | H={col_cargo} | I={col_depto} | Q={col_data_adm} | Z={col_ocorrencia} | AB={col_justificativa} | AM={col_data}")
    
    with st.expander("📊 Ocorrências disponíveis", expanded=False):
        oc = df[col_ocorrencia].value_counts().reset_index()
        oc.columns = ['Ocorrência', 'Quantidade']
        st.dataframe(oc, use_container_width=True)
    
    st.subheader("2. Selecione os tipos")
    opcoes = []
    for config in OCORRENCIAS_CONFIG:
        if config['tipo'] == 'unica': opcoes.append(config['nome'])
        elif config['tipo'] == 'multiplas_pasta_unica': opcoes.append(f"📁 {config['nome']} ({len(config['itens'])} tipos)")
        else: opcoes.append(f"📁 {config['nome']} ({len(config['justificativas'])} justificativas)")
    selecionados = st.multiselect("Selecione:", options=opcoes, default=opcoes)
    
    if st.button("🚀 Gerar Relatórios", type="primary", use_container_width=True):
        if not selecionados:
            st.warning("Selecione pelo menos um tipo.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            configs_selecionadas = []
            for config in OCORRENCIAS_CONFIG:
                if config['tipo'] == 'unica' and config['nome'] in selecionados: configs_selecionadas.append(config)
                elif config['tipo'] == 'multiplas':
                    label = f"📁 {config['nome']} ({len(config['justificativas'])} justificativas)"
                    if label in selecionados: configs_selecionadas.append(config)
                elif config['tipo'] == 'multiplas_pasta_unica':
                    label = f"📁 {config['nome']} ({len(config['itens'])} tipos)"
                    if label in selecionados: configs_selecionadas.append(config)
            
            total = len(configs_selecionadas)
            if total == 0:
                st.warning("Nenhuma selecionada.")
                st.stop()
            
            resultados_processados = []
            for i, config in enumerate(configs_selecionadas):
                status_text.text(f"Processando: {config['nome']}...")
                if config['tipo'] == 'unica':
                    eb, na, qtd = gerar_excel_ocorrencia(df, config, col_nome, col_cargo, col_depto, col_data_adm, col_ocorrencia, col_justificativa, col_data, mapa_colaboradores, mapa_supervisores)
                    resultados_processados.append({'config': config, 'excel_bytes': eb, 'nome_arquivo': na, 'qtd': qtd, 'tipo': 'unica'})
                elif config['tipo'] == 'multiplas_pasta_unica':
                    arquivos = {}; total_colab = 0
                    for item in config['itens']:
                        eb, na, qtd = gerar_excel_ocorrencia(df, item, col_nome, col_cargo, col_depto, col_data_adm, col_ocorrencia, col_justificativa, col_data, mapa_colaboradores, mapa_supervisores)
                        if eb is not None: arquivos[na] = eb; total_colab += qtd
                    resultados_processados.append({'config': config, 'arquivos': arquivos, 'pasta': config['pasta'], 'total_colab': total_colab, 'tipo': 'multiplas'})
                else:
                    arquivos, pasta, total_colab = gerar_pasta_ocorrencia(df, config, col_nome, col_cargo, col_depto, col_data_adm, col_ocorrencia, col_justificativa, col_data, mapa_colaboradores, mapa_supervisores)
                    resultados_processados.append({'config': config, 'arquivos': arquivos, 'pasta': pasta, 'total_colab': total_colab, 'tipo': 'multiplas'})
                progress_bar.progress((i + 1) / (total + 1))
            
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                for resultado in resultados_processados:
                    status_text.text(f"Compactando: {resultado['config']['nome']}...")
                    if resultado['tipo'] == 'unica':
                        if resultado['excel_bytes'] is not None: zf.writestr(resultado['nome_arquivo'], resultado['excel_bytes'])
                    else:
                        for nome_arquivo, excel_bytes in resultado['arquivos'].items():
                            zf.writestr(f"{resultado['pasta']}/{nome_arquivo}", excel_bytes)
                    progress_bar.progress((i + 1) / (total + 1) + 0.05)
            
            status_text.text("✅ Relatórios gerados!")
            if len(zip_buffer.getvalue()) > 0:
                st.download_button("📥 Baixar ZIP", data=zip_buffer.getvalue(), file_name="Relatorio_Ponto_Geral.zip", mime="application/zip", use_container_width=True)
            
            st.subheader("📋 Resumo")
            resumo = []
            for r in resultados_processados:
                if r['tipo'] == 'unica': resumo.append({'Ocorrência': r['config']['nome'], 'Colaboradores': r['qtd'], 'Status': '✅' if r['qtd'] > 0 else '⚠️'})
                else: resumo.append({'Ocorrência': f"📁 {r['config']['nome']}", 'Colaboradores': r['total_colab'], 'Status': '✅' if r['total_colab'] > 0 else '⚠️'})
            st.dataframe(pd.DataFrame(resumo), use_container_width=True, hide_index=True)
            
            # DEBUG
            with st.expander("🔬 Debug dos dados", expanded=False):
                debug_qtde = st.session_state.get('_debug_qtde', 0)
                debug_valores = st.session_state.get('_debug_valores', {})
                st.write(f"**Linhas filtradas no último processamento:** {debug_qtde}")
                if debug_valores:
                    st.write("**Valores únicos em _VALOR:**")
                    st.json(debug_valores)
                else:
                    st.warning("Nenhum valor encontrado em _VALOR! Verifique o filtro.")
else:
    st.info(" Carregue os arquivos para começar.")