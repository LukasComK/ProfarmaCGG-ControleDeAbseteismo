"""
Página: Relatório de Ponto Geral
Descrição: Geração de relatórios de ponto com ocorrências, gestores e turnos
"""

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


# ===== FUNÇÕES AUXILIARES PARA MEDIDA DISCIPLINAR E DEMISSÕES =====
def limpar_nome(nome):
    if pd.isna(nome) or str(nome).strip().upper() == "NAN":
        return ""
    nome_limpo = str(nome).replace('"', '').replace("'", "")
    return " ".join(nome_limpo.strip().upper().split())


def buscar_info_aproximada(nome_alvo, base_dados):
    if not nome_alvo:
        return None if isinstance(base_dados, dict) else False

    if nome_alvo in base_dados:
        val = base_dados[nome_alvo] if isinstance(base_dados, dict) else True
        if isinstance(val, bool) and isinstance(base_dados, dict):
            return None
        return val

    if len(nome_alvo) >= 8:
        for chave in base_dados:
            if isinstance(chave, str) and len(chave) >= 8:
                if nome_alvo in chave or chave in nome_alvo:
                    val = base_dados[chave]
                    if isinstance(val, bool) and isinstance(base_dados, dict):
                        return None
                    return val if isinstance(base_dados, dict) else True

    return None if isinstance(base_dados, dict) else False


def extrair_meses_de_valor(valor) -> set:
    meses = set()
    if pd.isna(valor):
        return meses

    if isinstance(valor, pd.Timestamp):
        meses.add(str(valor.month).zfill(2))
        return meses

    if hasattr(valor, 'month') and hasattr(valor, 'day') and hasattr(valor, 'year'):
        try:
            meses.add(str(valor.month).zfill(2))
            return meses
        except:
            pass

    texto = str(valor).strip().lower()
    if not texto:
        return meses

    match_dt = re.search(r'\b\d{1,2}/(\d{1,2})(?:/\d{2,4})?\b', texto)
    if match_dt:
        meses.add(match_dt.group(1).zfill(2))

    match_iso = re.search(r'\b\d{4}-(\d{1,2})-\d{1,2}\b', texto)
    if match_iso:
        meses.add(match_iso.group(1).zfill(2))

    mapa_meses = {
        'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04',
        'mai': '05', 'jun': '06', 'jul': '07', 'ago': '08',
        'set': '09', 'out': '10', 'nov': '11', 'dez': '12'
    }
    for nome_mes, num_mes in mapa_meses.items():
        if nome_mes in texto:
            meses.add(num_mes)

    return meses


def formatar_data_exibicao(valor) -> str:
    if pd.isna(valor) or str(valor).strip() in ['', 'NaT', 'NaN', 'nan']:
        return 'Data Indisponível'

    if isinstance(valor, pd.Timestamp):
        return valor.strftime('%d/%m/%Y')

    if hasattr(valor, 'strftime') and hasattr(valor, 'year'):
        try:
            return valor.strftime('%d/%m/%Y')
        except:
            pass

    texto = str(valor).strip()
    try:
        dt = pd.to_datetime(texto, dayfirst=True, errors='coerce')
        if pd.notna(dt):
            return dt.strftime('%d/%m/%Y')
    except:
        pass

    return texto[:10] if texto else 'Data Indisponível'


def normalizar_nome_chave(nome) -> str:
    return safe_unidecode(limpar_nome(nome))


def carregar_arquivo(f, sheet=None, todas_abas=False):
    raw_bytes = f.getvalue()
    
    # CSV
    if f.name.lower().endswith('.csv'):
        try:
            csv_str = raw_bytes.decode('latin1', errors='ignore')
            linhas = csv_str.split('\n')[:5]
            sep = ';' if any(';' in l for l in linhas) else ','
            df_csv = pd.read_csv(io.StringIO(csv_str), sep=sep, header=None, names=range(250), quoting=3, on_bad_lines='skip', engine='python')
            if todas_abas:
                return {"Aba_CSV": df_csv}
            return df_csv
        except Exception as e:
            raise ValueError(f"Erro ao processar o arquivo CSV '{f.name}': {str(e)}")

    # XLSX/XLS
    try:
        if todas_abas:
            return pd.read_excel(io.BytesIO(raw_bytes), sheet_name=None, header=None)
        if sheet:
            try:
                return pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sheet, header=None)
            except:
                pass
        return pd.read_excel(io.BytesIO(raw_bytes), header=None)
    except Exception as e_excel:
        try:
            if todas_abas:
                return pd.read_excel(io.BytesIO(raw_bytes), sheet_name=None, header=None, engine='openpyxl')
            if sheet:
                try:
                    return pd.read_excel(io.BytesIO(raw_bytes), sheet_name=sheet, header=None, engine='openpyxl')
                except:
                    pass
            return pd.read_excel(io.BytesIO(raw_bytes), header=None, engine='openpyxl')
        except:
            try:
                html_str = raw_bytes.decode('latin1', errors='ignore')
                dfs = pd.read_html(io.StringIO(html_str), header=None)
                if dfs:
                    if todas_abas:
                        return {"Aba_HTML": dfs[0]}
                    return dfs[0]
            except:
                try:
                    csv_str = raw_bytes.decode('latin1', errors='ignore')
                    df_csv = pd.read_csv(io.StringIO(csv_str), sep=';', header=None)
                    if todas_abas:
                        return {"Aba_CSV": df_csv}
                    return df_csv
                except:
                    raise ValueError(f"Não foi possível ler '{f.name}'.")


def tem_mes_correspondente(lista_datas_fi, texto_medida):
    if not texto_medida:
        return False
    
    texto_medida = str(texto_medida).lower()

    meses_fi = set()
    for d in lista_datas_fi:
        meses_fi.update(extrair_meses_de_valor(d))

    meses_med = extrair_meses_de_valor(texto_medida)

    # Se não detectou data/mês nenhum na medida, mantém a regra anterior
    if not meses_med:
        return True
        
    return bool(meses_fi.intersection(meses_med))


def processar_medidas(arquivo_medida, termo_busca="FALTA INJUSTIFICADA"):
    medidas_dict = {}
    if arquivo_medida is None:
        return medidas_dict
    try:
        df_med = carregar_arquivo(arquivo_medida)
        st.info(f"📋 Medida Disciplinar ({termo_busca}): {len(df_med)} linhas, {len(df_med.columns)} colunas")
        col_max_med = len(df_med.columns)
        for r in range(len(df_med)):
            if col_max_med > 1:
                nome_val = df_med.iloc[r, 1]
                if pd.isna(nome_val):
                    continue
                nome = limpar_nome(nome_val)
            else:
                continue
            tipo_val = df_med.iloc[r, 3] if col_max_med > 3 else ""
            tipo = limpar_nome(tipo_val)
            obs_val = str(df_med.iloc[r, 4]).strip() if col_max_med > 4 and pd.notna(df_med.iloc[r, 4]) else ""
            data_val = formatar_data_exibicao(df_med.iloc[r, 5]) if col_max_med > 5 else "Data Indisponível"
            eh_atraso_busca = termo_busca.upper().startswith('ATRAS')
            tem_atraso_no_tipo = 'ATRAS' in tipo
            
            if eh_atraso_busca:
                partes = []
                if obs_val:
                    partes.append(f"OBS: {obs_val}")
                if data_val:
                    partes.append(f"DATA: {data_val}")
                detalhe_data = " | ".join(partes)
                tipo_match = tem_atraso_no_tipo
            else:
                detalhe_data = obs_val
                tipo_match = termo_busca.upper() in tipo
            
            if nome and tipo_match:
                medidas_dict[nome] = detalhe_data
        
        st.info(f"✅ Medida Disciplinar ({termo_busca}): {len(medidas_dict)} registros")
    except Exception as e:
        st.warning(f"Aviso: Erro ao processar Medida Disciplinar ({termo_busca}): {e}")
        import traceback
        st.warning(traceback.format_exc())
    return medidas_dict


def processar_demissoes(arquivo_dem):
    demissoes_dict = {}
    if arquivo_dem is None:
        return demissoes_dict
    try:
        if arquivo_dem.name.lower().endswith('.csv'):
            try:
                raw_bytes = arquivo_dem.getvalue()
                csv_str = raw_bytes.decode('latin1', errors='ignore')
                sep = ';' if any(';' in csv_str[:500]) else ','
                df_dem = pd.read_csv(io.StringIO(csv_str), sep=sep, header=None, quoting=3, on_bad_lines='skip', engine='python')
                st.info(f"📋 Demissões (CSV): {len(df_dem)} linhas, {len(df_dem.columns)} colunas")
                col_max_dem = len(df_dem.columns)
                for r in range(len(df_dem)):
                    nome = limpar_nome(df_dem.iloc[r, 6]) if col_max_dem > 6 else ""
                    if nome:
                        val_data = df_dem.iloc[r, 12] if col_max_dem > 12 else None
                        d_data = formatar_data_exibicao(val_data)
                        val_tipo = df_dem.iloc[r, 14] if col_max_dem > 14 else None
                        d_tipo = str(val_tipo).strip().capitalize() if pd.notna(val_tipo) else "Tipo Indisponível"
                        demissoes_dict[nome] = {'data': d_data, 'tipo': d_tipo}
            except Exception as e_csv:
                st.warning(f"Aviso: Erro ao processar Demissões CSV: {e_csv}")
            return demissoes_dict
        
        dfs_dem = carregar_arquivo(arquivo_dem, todas_abas=True)
        if isinstance(dfs_dem, dict):
            st.info(f"📋 Demissões (XLSX): {len(dfs_dem)} abas")
            for nome_aba, df_dem_aba in dfs_dem.items():
                col_max_dem = len(df_dem_aba.columns)
                for r in range(len(df_dem_aba)):
                    if col_max_dem > 1:
                        nome = limpar_nome(df_dem_aba.iloc[r, 1])
                    else:
                        continue
                    if nome:
                        val_data = df_dem_aba.iloc[r, 3] if col_max_dem > 3 else None
                        d_data = formatar_data_exibicao(val_data)
                        val_tipo = df_dem_aba.iloc[r, 5] if col_max_dem > 5 else None
                        d_tipo = str(val_tipo).strip().capitalize() if pd.notna(val_tipo) else "Tipo Indisponível"
                        info_atual = demissoes_dict.get(nome)
                        if not info_atual or (info_atual.get('data') in ['', 'Data Indisponível'] and d_data != 'Data Indisponível'):
                            demissoes_dict[nome] = {'data': d_data, 'tipo': d_tipo}
    except Exception as e:
        st.warning(f"Aviso: Erro ao processar Demissões: {e}")
    return demissoes_dict


def processar_ocorrencia(
    df: pd.DataFrame,
    termo_ocorrencia: str, termo_justificativa: str,
    col_nome: str, col_cargo: str, col_depto: str, col_data_adm: str,
    col_ocorrencia: str, col_justificativa: str, col_data: str,
    col_marcacoes: str = None, col_atraso_calc: str = None,
    mapa_colaboradores: Dict[str, Dict] = None,
    mapa_supervisores: Dict[str, str] = None,
    medidas_dict: Dict = None,
    demissoes_dict: Dict = None,
    medidas_atraso_dict: Dict = None
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
    
    eh_falta = 'falta' in termo_occ_norm
    eh_atraso = 'atraso' in termo_occ_norm
    
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
    
    # ========================================================
    # CONSTRUCAO DO DETALHAMENTO
    # ========================================================
    
    # Cria a coluna de valor
    def extrair_valor_marcacao(row):
        # Para Entrada em atraso, usa Col Y (marcações) e Col AA (cálculo)
        if eh_atraso and col_marcacoes and col_atraso_calc:
            marc_str = str(row[col_marcacoes]).strip() if pd.notna(row[col_marcacoes]) else ""
            calc_str = str(row[col_atraso_calc]).strip() if pd.notna(row[col_atraso_calc]) else ""
            
            # Formato: "06:00  06:50" -> separar
            if "  " in marc_str:
                partes = marc_str.split("  ")
                escala = partes[0].strip()
                entrou = partes[1].strip()
            else:
                escala = marc_str
                entrou = ""
            
            # Formatar saída
            saida = f"Escala {escala} - Entrou {entrou}"
            if calc_str and calc_str.lower() not in ['nan', 'nat']:
                saida += f" - {calc_str}"
            return saida
        
        # Padrão: usa coluna de justificativa
        for col_ref in (col_justificativa, col_ocorrencia):
            if col_ref in row.index:
                valor = str(row[col_ref]).strip()
                if valor and valor.lower() not in ['nan', 'nat']:
                    return valor
        return ''

    df_filtrado['_VALOR'] = df_filtrado.apply(extrair_valor_marcacao, axis=1)
    
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
    if eh_falta or eh_atraso:
        dict_medidas_usar = medidas_dict if eh_falta else medidas_atraso_dict
        
        def get_medida(nome):
            if not dict_medidas_usar:
                return "Não solicitada"
            nome_limpo = limpar_nome(nome)
            med = buscar_info_aproximada(nome_limpo, dict_medidas_usar)
            if med:
                datas_colab = df_filtrado[df_filtrado['Colaborador'] == nome]['Data_Formatada'].tolist()
                if tem_mes_correspondente(datas_colab, med):
                    return f"Sim ({med})"
                return "Não solicitada"
            return "Não solicitada"
        
        df_filtrado['Medida Disciplinar'] = df_filtrado['Colaborador'].apply(get_medida)
        
        def get_demissao(nome):
            if not demissoes_dict:
                return "Sem projeção"
            nome_limpo = limpar_nome(nome)
            dem = buscar_info_aproximada(nome_limpo, demissoes_dict)
            if dem:
                return f"{dem['data']} - {dem['tipo']}"
            return "Sem projeção"
        df_filtrado['Projeção Desligamento'] = df_filtrado['Colaborador'].apply(get_demissao)
    
    # === COLUNAS FIXAS DO RELATORIO ===
    cols_fixas = ['Colaborador', 'Cargo', 'Departamento', 'Gestor', 'Supervisor', 'Turno', 'Data Admissão', 'Tempo de Serviço']
    if eh_falta or eh_atraso:
        cols_fixas.append('Medida Disciplinar')
        cols_fixas.append('Projeção Desligamento')
    
    cols_fixas_existentes = [c for c in cols_fixas if c in df_filtrado.columns]
    
    todas_datas = sorted(df_filtrado['Data_Formatada'].dropna().unique().tolist())
    try:
        todas_datas.sort(key=lambda x: datetime.strptime(x, '%d/%m/%Y') if x else datetime.min)
    except:
        pass
    
    from collections import defaultdict
    lookup = defaultdict(list)
    
    for _, row in df_filtrado.iterrows():
        colab = str(row['Colaborador']).strip()
        data = str(row['Data_Formatada']).strip()
        valor = str(row['_VALOR']).strip()
        if valor and valor.lower() not in ['nan', 'nat']:
            lookup[(colab, data)].append(valor)
    
    lookup_final = {}
    for chave, valores in lookup.items():
        valores_unicos = sorted(set(valores))
        lookup_final[chave] = ' | '.join(valores_unicos)
    
    df_fixas = df_filtrado[cols_fixas_existentes].drop_duplicates(subset=['Colaborador']).reset_index(drop=True)
    
    linhas = []
    for _, fixa_row in df_fixas.iterrows():
        colab = str(fixa_row['Colaborador']).strip()
        linha = {c: fixa_row[c] for c in cols_fixas_existentes}
        
        datas_com_valor = []
        qtd = 0
        for data in todas_datas:
            valor = lookup_final.get((colab, data), '')
            linha[data] = valor
            if valor:
                qtd += 1
                datas_com_valor.append(data)
        
        linha['Quantidade Ocorrências'] = qtd
        linha['Datas das Ocorrências'] = ', '.join(datas_com_valor)
        linhas.append(linha)
    
    df_detalhe = pd.DataFrame(linhas)
    df_detalhe = df_detalhe.sort_values('Quantidade Ocorrências', ascending=False).reset_index(drop=True)
    
    cols_finais = cols_fixas_existentes + ['Quantidade Ocorrências', 'Datas das Ocorrências'] + todas_datas
    cols_existentes = [c for c in cols_finais if c in df_detalhe.columns]
    df_detalhe = df_detalhe[cols_existentes]
    
    cols_ranking = cols_fixas + ['Quantidade Ocorrências', 'Datas das Ocorrências']
    cols_ranking_existentes = [c for c in cols_ranking if c in df_detalhe.columns]
    df_ranking = df_detalhe[cols_ranking_existentes].copy()
    df_ranking = df_ranking.sort_values('Quantidade Ocorrências', ascending=False).reset_index(drop=True)
    df_ranking.insert(0, 'Posição', range(1, len(df_ranking) + 1))
    
    return df_detalhe, df_ranking


def gerar_planilha_ocorrencia(df_detalhe: pd.DataFrame, df_ranking: pd.DataFrame, nome_aba_detalhe: str, writer: pd.ExcelWriter,
                              col_widths_extra: Dict = None):
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
    if col_widths_extra:
        col_widths.update(col_widths_extra)
    col_widths_rank = {'Posição': 8, 'Colaborador': 36, 'Cargo': 20, 'Departamento': 28, 'Gestor': 30, 'Supervisor': 30, 'Turno': 14, 'Data Admissão': 16, 'Tempo de Serviço': 16, 'Quantidade Ocorrências': 20}
    if col_widths_extra:
        col_widths_rank.update(col_widths_extra)
    
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


def localizar_coluna(df: pd.DataFrame, candidatos: List[str], indice_padrao: int = None):
    for nome in candidatos:
        if nome in df.columns:
            return nome
    if indice_padrao is not None and len(df.columns) > indice_padrao:
        return df.columns[indice_padrao]
    return None


def extrair_horarios(texto) -> List[str]:
    if pd.isna(texto):
        return []
    encontrados = re.findall(r'\b\d{2}:\d{2}\b', str(texto))
    horarios = []
    for horario in encontrados:
        if horario not in horarios:
            horarios.append(horario)
        if len(horarios) == 4:
            break
    return horarios


def horario_para_minutos(horario: str):
    if not horario:
        return None
    try:
        horas, minutos = horario.split(':')
        return int(horas) * 60 + int(minutos)
    except Exception:
        return None


def minutos_para_horario(minutos: int) -> str:
    if minutos is None:
        return ''
    minutos = int(round(minutos)) % 1440
    horas = minutos // 60
    resto = minutos % 60
    return f'{horas:02d}:{resto:02d}'


def diferenca_circular_minutos(real: str, oficial: str):
    real_min = horario_para_minutos(real)
    oficial_min = horario_para_minutos(oficial)
    if real_min is None or oficial_min is None:
        return None
    delta = real_min - oficial_min
    if delta > 720:
        delta -= 1440
    elif delta < -720:
        delta += 1440
    return int(delta)


def formatar_delta_minutos(delta: float) -> str:
    if delta is None or pd.isna(delta):
        return ''
    delta = int(round(delta))
    sinal = '+' if delta > 0 else ''
    return f'{sinal}{delta} min'


def parse_duracao_minutos(valor):
    if pd.isna(valor):
        return None
    texto = str(valor).strip()
    if not texto or texto.lower() in {'nan', 'nat'}:
        return None
    sinal = -1 if texto.startswith('-') else 1
    texto = texto.lstrip('+-').strip()
    match = re.match(r'^(\d{1,2}):(\d{2})$', texto)
    if not match:
        return None
    horas = int(match.group(1))
    minutos = int(match.group(2))
    return sinal * (horas * 60 + minutos)


def formatar_duracao_minutos(valor) -> str:
    if valor is None or pd.isna(valor):
        return ''
    valor = int(round(valor))
    sinal = '-' if valor < 0 else ''
    valor = abs(valor)
    horas = valor // 60
    minutos = valor % 60
    return f'{sinal}{horas:02d}:{minutos:02d}'


def valor_excel_seguro(valor):
    if valor is None:
        return ''
    if isinstance(valor, float) and (pd.isna(valor) or np.isinf(valor)):
        return ''
    if pd.isna(valor):
        return ''
    return valor


def identificar_dia_evento(grupo_dia: pd.DataFrame) -> Tuple[str, str]:
    texto_parts = []
    for col in ['Ocorrencia', 'Justificativa', 'TipoOcorrencia']:
        if col in grupo_dia.columns:
            serie = grupo_dia[col].dropna().astype(str).str.strip()
            texto_parts.extend([item for item in serie.tolist() if item and item.lower() not in {'nan', 'nat', '0', '0.0', '00:00'}])

    texto = ' | '.join(texto_parts).lower()
    if not texto:
        return '', ''

    if 'falta' in texto or 'sem marcação' in texto or 'sem marcacao' in texto or 'ausência' in texto or 'ausencia' in texto:
        return 'falta', 'Falta/Ausência'

    if 'feriado' in texto:
        return 'extra_forte', 'Trabalho em feriado'
    if 'folga' in texto and ('hora extra' in texto or 'pagar horas extras' in texto or 'trabalh' in texto):
        return 'extra_forte', 'Trabalho em folga'
    if 'hora extra folga' in texto:
        return 'extra_forte', 'Hora extra em folga'
    if 'pagar horas extras' in texto:
        return 'extra_forte', 'Dia fora da jornada com hora extra'
    if 'serviço externo' in texto or 'servico externo' in texto:
        return 'extra_moderado', 'Dia fora da jornada'

    return '', ''


def processar_alteracoes_escala(
    df: pd.DataFrame,
    col_nome: str,
    col_cargo: str,
    col_depto: str,
    col_data_adm: str,
    col_data: str,
    col_escala: str,
    col_marcacoes: str,
    mapa_colaboradores: Dict = None,
    mapa_supervisores: Dict = None,
    col_escala_codigo: str = None,
    tolerancia_minutos: int = 10,
    minimo_dias: int = 3,
    consistencia_minima: float = 70.0,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
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

    coluna_escala_base = col_escala_codigo if col_escala_codigo and col_escala_codigo in df.columns else col_escala
    df_base = df.copy()
    df_base['_Data_Ord'] = pd.to_datetime(df_base[col_data], errors='coerce', dayfirst=True)
    df_base['_Data'] = df_base['_Data_Ord'].dt.strftime('%d/%m/%Y')
    df_base['_Horarios_Oficiais'] = df_base.apply(
        lambda row: extrair_horarios(str(row.get(coluna_escala_base, '')).strip())
        if coluna_escala_base and pd.notna(row.get(coluna_escala_base, np.nan))
        else extrair_horarios(str(row.get(col_escala, '')).strip()) if pd.notna(row.get(col_escala, np.nan)) else [],
        axis=1,
    )
    df_base['_Horarios_Reais'] = df_base[col_marcacoes].apply(extrair_horarios) if col_marcacoes in df_base.columns else [[] for _ in range(len(df_base))]
    df_base['_Qtd_Reais'] = df_base['_Horarios_Reais'].apply(len)
    for col_balanco in ['BancoDeHoras', 'HoraExtra', 'Desconto']:
        if col_balanco in df_base.columns:
            df_base[f'_{col_balanco}_Min'] = df_base[col_balanco].apply(parse_duracao_minutos)
        else:
            df_base[f'_{col_balanco}_Min'] = np.nan

    registros_diarios = []
    for (nome_colab, data_ord), grupo_dia in df_base.groupby([col_nome, '_Data_Ord'], dropna=False):
        if pd.isna(nome_colab) or pd.isna(data_ord):
            continue

        grupo_dia = grupo_dia.copy()
        primeira_linha = grupo_dia.iloc[0]
        info = get_info_colaborador(nome_colab)

        oficiais = []
        for itens in grupo_dia['_Horarios_Oficiais']:
            if itens:
                oficiais = itens
                break

        linhas_com_marcacao = grupo_dia[grupo_dia['_Qtd_Reais'] > 0]
        if len(linhas_com_marcacao) > 0:
            linha_real = linhas_com_marcacao.sort_values(['_Qtd_Reais', '_Data_Ord'], ascending=[False, True]).iloc[0]
            reais = linha_real['_Horarios_Reais']
        else:
            reais = []

        if not reais:
            lista_reais = []
            for itens in grupo_dia['_Horarios_Reais']:
                if itens:
                    lista_reais = itens
                    break
            reais = lista_reais

        tipo_evento, alerta_dia = identificar_dia_evento(grupo_dia)

        saldo_bh_min = None
        for col_balanco in ['BancoDeHoras', 'HoraExtra', 'Desconto']:
            valores_balanco = [v for v in grupo_dia[f'_{col_balanco}_Min'].tolist() if pd.notna(v) and int(v) != 0]
            if valores_balanco:
                saldo_bh_min = max(valores_balanco, key=lambda v: abs(v))
                break

        entrada_prevista = oficiais[0] if len(oficiais) >= 1 else ''
        saida_prevista = oficiais[3] if len(oficiais) >= 4 else (oficiais[-1] if oficiais else '')
        entrada_real = reais[0] if len(reais) >= 1 else ''
        saida_real = reais[3] if len(reais) >= 4 else (reais[-1] if len(reais) >= 2 else '')

        delta_entrada = diferenca_circular_minutos(entrada_real, entrada_prevista) if entrada_real and entrada_prevista else None
        delta_saida = diferenca_circular_minutos(saida_real, saida_prevista) if saida_real and saida_prevista else None

        registros_diarios.append({
            'Colaborador': nome_colab,
            'Cargo': primeira_linha[col_cargo],
            'Departamento': primeira_linha[col_depto],
            'Gestor': info['gestor'],
            'Supervisor': info['supervisor'],
            'Turno': info['turno'],
            'Data Admissão': formatar_data_br(primeira_linha[col_data_adm]),
            'Tempo de Serviço': calcular_tempo_servico(primeira_linha[col_data_adm]),
            'Data': primeira_linha['_Data'] if pd.notna(primeira_linha['_Data']) else '',
            'Data_Ord': data_ord,
            'Escala Oficial': ' '.join(oficiais[:4]) if oficiais else '',
            'Batidas do Dia': ' '.join(reais[:4]) if reais else '',
            'Entrada Prevista': entrada_prevista,
            'Entrada Real': entrada_real,
            'Saída Prevista': saida_prevista,
            'Saída Real': saida_real,
            'Delta Entrada': delta_entrada,
            'Delta Saída': delta_saida,
            'Saldo BH Min': saldo_bh_min,
            'Saldo BH': formatar_duracao_minutos(saldo_bh_min),
            'Tem Jornada Comparavel': bool(len(reais) >= 2),
            'Alerta Dia': alerta_dia,
            'Tipo Evento': tipo_evento,
            'Dia Extra': tipo_evento in {'extra_forte', 'extra_moderado'},
            'Dia Falta': tipo_evento == 'falta',
        })

    df_detalhe = pd.DataFrame(registros_diarios)
    if df_detalhe.empty:
        colunas_vazias = ['Colaborador', 'Cargo', 'Departamento', 'Gestor', 'Supervisor', 'Turno', 'Data', 'Escala Oficial', 'Batidas do Dia', 'Entrada Prevista', 'Entrada Real', 'Saída Prevista', 'Saída Real', 'Delta Entrada', 'Delta Saída', 'Saldo BH', 'Alerta Dia']
        df_vazio = pd.DataFrame(columns=colunas_vazias)
        return df_vazio, df_vazio

    df_detalhe = df_detalhe.sort_values(['Colaborador', 'Data_Ord'], na_position='last').reset_index(drop=True)

    linhas_resumo = []
    for colaborador, grupo in df_detalhe.groupby('Colaborador', sort=False):
        total_dias = len(grupo)
        dias_extra = int(grupo['Dia Extra'].sum()) if 'Dia Extra' in grupo.columns else 0
        dias_falta = int(grupo['Dia Falta'].sum()) if 'Dia Falta' in grupo.columns else 0
        dias_extra_fortes = int((grupo['Tipo Evento'] == 'extra_forte').sum()) if 'Tipo Evento' in grupo.columns else 0
        taxa_falta = (dias_falta / total_dias) if total_dias else 0.0

        grupo_regular = grupo[(grupo['Tem Jornada Comparavel']) & (~grupo['Dia Extra']) & (~grupo['Dia Falta'])].copy()
        dias_regulares = len(grupo_regular)
        med_entrada = float(np.median(grupo_regular['Delta Entrada'])) if dias_regulares > 0 and grupo_regular['Delta Entrada'].notna().any() else None
        med_saida = float(np.median(grupo_regular['Delta Saída'])) if dias_regulares > 0 and grupo_regular['Delta Saída'].notna().any() else None

        if dias_regulares > 0 and med_entrada is not None and med_saida is not None:
            consistencia_horario = ((grupo_regular['Delta Entrada'].sub(med_entrada).abs() <= tolerancia_minutos) & (grupo_regular['Delta Saída'].sub(med_saida).abs() <= tolerancia_minutos)).mean() * 100
        else:
            consistencia_horario = 0.0

        saldo_vals = grupo['Saldo BH Min'].dropna()
        saldo_medio = float(np.median(saldo_vals)) if len(saldo_vals) > 0 else None
        saldo_abs_medio = float(np.median(np.abs(saldo_vals))) if len(saldo_vals) > 0 else None
        saldo_pos_pct = float((saldo_vals > 0).mean() * 100) if len(saldo_vals) > 0 else 0.0
        saldo_neg_pct = float((saldo_vals < 0).mean() * 100) if len(saldo_vals) > 0 else 0.0

        penalty_extra = min(15, dias_extra * 5)
        penalty_falta = min(30, dias_falta * 8)
        consistencia_final = max(0.0, consistencia_horario - penalty_extra - penalty_falta)

        flag_horario = dias_regulares >= minimo_dias and taxa_falta < 0.40 and med_entrada is not None and med_saida is not None and (abs(med_entrada) >= 15 or abs(med_saida) >= 15) and consistencia_horario >= consistencia_minima
        flag_saldo = total_dias >= minimo_dias and taxa_falta < 0.50 and saldo_abs_medio is not None and saldo_abs_medio >= 240 and max(saldo_pos_pct, saldo_neg_pct) >= 70

        if dias_falta >= dias_regulares and dias_extra_fortes == 0:
            continue

        if not flag_horario and not flag_saldo:
            continue

        primeira_linha = grupo.iloc[0]
        escala_oficial = primeira_linha['Escala Oficial']
        entrada_prevista = primeira_linha['Entrada Prevista']
        saida_prevista = primeira_linha['Saída Prevista']
        entrada_real_media = minutos_para_horario(horario_para_minutos(entrada_prevista) + int(round(med_entrada))) if entrada_prevista and med_entrada is not None else ''
        saida_real_media = minutos_para_horario(horario_para_minutos(saida_prevista) + int(round(med_saida))) if saida_prevista and med_saida is not None else ''

        motivos = []
        if flag_saldo:
            saldo_texto = formatar_duracao_minutos(saldo_medio)
            orientacao = 'positivo' if (saldo_medio or 0) > 0 else 'negativo'
            motivos.append(f'Saldo de horas {orientacao} ({saldo_texto})')
        if flag_horario:
            motivos.append('Entrada e saída deslocadas de forma consistente')
        if dias_extra > 0:
            motivos.append(f'{dias_extra} dia(s) fora da jornada')
        if dias_falta > 0:
            motivos.append(f'{dias_falta} dia(s) com falta/ausência')
        if taxa_falta >= 0.40:
            motivos.append(f'Faltas em {round(taxa_falta * 100)}% dos dias')

        if dias_extra_fortes > 0 and flag_horario:
            recomendacao = 'Troca de escala evidente'
        elif flag_saldo and flag_horario:
            recomendacao = 'Rever escala agora'
        elif flag_saldo:
            recomendacao = 'Revisar escala por saldo de horas'
        else:
            recomendacao = 'Revisar escala e batidas'

        pontuacao = 0.0
        if flag_horario:
            pontuacao += 60.0
        if flag_saldo and saldo_abs_medio is not None:
            pontuacao += 25.0 + min(15.0, float(saldo_abs_medio) / 60.0 * 3.0)
        if dias_extra > 0:
            pontuacao += min(10.0, dias_extra * 2.5)
        if dias_extra_fortes > 0:
            pontuacao += min(20.0, dias_extra_fortes * 10.0)
        if dias_falta > 0:
            pontuacao -= min(20.0, dias_falta * 3.0)
        if taxa_falta >= 0.40:
            pontuacao -= min(20.0, taxa_falta * 20.0)
        if not flag_horario and dias_extra > 0:
            pontuacao *= 0.6

        linhas_resumo.append({
            'Colaborador': primeira_linha['Colaborador'],
            'Cargo': primeira_linha['Cargo'],
            'Departamento': primeira_linha['Departamento'],
            'Gestor': primeira_linha['Gestor'],
            'Turno': primeira_linha['Turno'],
            'Escala Oficial': escala_oficial,
            'Entrada Prevista': entrada_prevista,
            'Entrada Real Média': entrada_real_media,
            'Saída Prevista': saida_prevista,
            'Saída Real Média': saida_real_media,
            'Saldo Médio BH': formatar_duracao_minutos(saldo_medio),
            'Dias Analisados': total_dias,
            'Dias Fora da Jornada': dias_extra,
            'Dias com Falta': dias_falta,
            'Falta %': round(taxa_falta * 100, 1),
            'Consistência': round(consistencia_final, 1),
            'Pontuação': round(pontuacao, 1),
            'Evidência': 'Alta' if (dias_extra_fortes > 0 and flag_horario) else ('Média' if flag_horario else 'Baixa'),
            'Alerta': ' | '.join(motivos),
            'Recomendação': recomendacao,
        })

    df_resumo = pd.DataFrame(linhas_resumo)
    if not df_resumo.empty:
        df_resumo = df_resumo.sort_values(['Pontuação', 'Consistência', 'Dias Analisados', 'Dias Fora da Jornada', 'Dias com Falta'], ascending=[False, False, False, True, True]).reset_index(drop=True)
        df_resumo.insert(0, 'Posição', range(1, len(df_resumo) + 1))

    df_detalhe = df_detalhe.drop(columns=['Data_Ord', 'Saldo BH Min'], errors='ignore')
    return df_resumo, df_detalhe


def gerar_planilha_alteracoes_escala(df_resumo: pd.DataFrame, df_detalhe: pd.DataFrame) -> bytes:
    buffer_saida = io.BytesIO()
    with pd.ExcelWriter(buffer_saida, engine='xlsxwriter') as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({'bold': True, 'font_size': 11, 'font_color': COR_BRANCO, 'bg_color': COR_VERDE_ESCURO, 'border': 1, 'border_color': COR_VERDE_ESCURO, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
        cell_fmt = workbook.add_format({'font_size': 10, 'border': 1, 'border_color': '#BFBFBF', 'text_wrap': True, 'valign': 'vcenter'})
        cell_alt_fmt = workbook.add_format({'font_size': 10, 'border': 1, 'border_color': '#BFBFBF', 'bg_color': COR_CINZA_CLARO, 'text_wrap': True, 'valign': 'vcenter'})
        cell_qtd_fmt = workbook.add_format({'font_size': 11, 'bold': True, 'font_color': COR_PRETO, 'border': 1, 'border_color': '#BFBFBF', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
        cell_qtd_alt_fmt = workbook.add_format({'font_size': 11, 'bold': True, 'font_color': COR_PRETO, 'border': 1, 'border_color': '#BFBFBF', 'bg_color': COR_CINZA_CLARO, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
        cell_alta_fmt = workbook.add_format({'font_size': 11, 'bold': True, 'font_color': COR_VERMELHO, 'border': 1, 'border_color': '#BFBFBF', 'bg_color': '#FFF0F0', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})
        cell_media_fmt = workbook.add_format({'font_size': 11, 'bold': True, 'font_color': COR_LARANJA, 'border': 1, 'border_color': '#BFBFBF', 'bg_color': '#FFF8E7', 'text_wrap': True, 'valign': 'vcenter', 'align': 'center'})

        if df_resumo.empty:
            df_resumo = pd.DataFrame(columns=['Posição', 'Colaborador', 'Cargo', 'Departamento', 'Gestor', 'Turno', 'Escala Oficial', 'Entrada Prevista', 'Entrada Real Média', 'Saída Prevista', 'Saída Real Média', 'Saldo Médio BH', 'Dias Analisados', 'Dias Fora da Jornada', 'Dias com Falta', 'Falta %', 'Consistência', 'Pontuação', 'Evidência', 'Alerta', 'Recomendação'])

        if df_detalhe.empty:
            df_detalhe = pd.DataFrame(columns=['Colaborador', 'Cargo', 'Departamento', 'Gestor', 'Supervisor', 'Turno', 'Data', 'Escala Oficial', 'Batidas do Dia', 'Entrada Prevista', 'Entrada Real', 'Saída Prevista', 'Saída Real', 'Delta Entrada', 'Delta Saída', 'Saldo BH'])

        df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
        ws_resumo = writer.sheets['Resumo']
        for col_num, value in enumerate(df_resumo.columns.values):
            ws_resumo.write(0, col_num, value, header_fmt)
        for col_num, col_name in enumerate(df_resumo.columns):
            largura = 18
            if col_name in {'Colaborador'}:
                largura = 36
            elif col_name in {'Cargo'}:
                largura = 20
            elif col_name in {'Departamento'}:
                largura = 28
            elif col_name in {'Gestor'}:
                largura = 30
            elif col_name in {'Turno'}:
                largura = 12
            elif col_name in {'Escala Oficial'}:
                largura = 28
            elif col_name in {'Entrada Prevista', 'Entrada Real Média', 'Saída Prevista', 'Saída Real Média', 'Saldo Médio BH'}:
                largura = 16
            elif col_name in {'Falta %'}:
                largura = 10
            elif col_name in {'Recomendação'}:
                largura = 26
            elif col_name in {'Evidência'}:
                largura = 12
            elif col_name in {'Alerta'}:
                largura = 36
            elif col_name in {'Consistência', 'Pontuação'}:
                largura = 12
            ws_resumo.set_column(col_num, col_num, largura)
        idx_consistencia = list(df_resumo.columns).index('Consistência') if 'Consistência' in df_resumo.columns else -1
        idx_pontuacao = list(df_resumo.columns).index('Pontuação') if 'Pontuação' in df_resumo.columns else -1
        for row_num in range(1, len(df_resumo) + 1):
            is_alt = row_num % 2 == 0
            for col_num in range(len(df_resumo.columns)):
                valor = valor_excel_seguro(df_resumo.iloc[row_num - 1, col_num])
                if col_num in {idx_consistencia, idx_pontuacao}:
                    valor_num = pd.to_numeric(valor, errors='coerce')
                    if pd.notna(valor_num) and float(valor_num) >= 80 and col_num == idx_consistencia:
                        fmt = cell_alta_fmt
                    elif pd.notna(valor_num) and float(valor_num) >= 70 and col_num == idx_consistencia:
                        fmt = cell_media_fmt
                    elif pd.notna(valor_num) and float(valor_num) >= 80 and col_num == idx_pontuacao:
                        fmt = cell_alta_fmt
                    elif pd.notna(valor_num) and float(valor_num) >= 50 and col_num == idx_pontuacao:
                        fmt = cell_media_fmt
                    elif is_alt:
                        fmt = cell_qtd_alt_fmt
                    else:
                        fmt = cell_qtd_fmt
                elif is_alt:
                    fmt = cell_alt_fmt
                else:
                    fmt = cell_fmt
                ws_resumo.write(row_num, col_num, valor, fmt)
        ws_resumo.freeze_panes(1, 0)
        ws_resumo.autofilter(0, 0, len(df_resumo), len(df_resumo.columns) - 1)
        ws_resumo.set_row(0, 30)

        df_detalhe.to_excel(writer, sheet_name='Detalhe', index=False)
        ws_detalhe = writer.sheets['Detalhe']
        for col_num, value in enumerate(df_detalhe.columns.values):
            ws_detalhe.write(0, col_num, value, header_fmt)
        for col_num, col_name in enumerate(df_detalhe.columns):
            largura = 18
            if col_name in {'Colaborador'}:
                largura = 36
            elif col_name in {'Cargo'}:
                largura = 20
            elif col_name in {'Departamento'}:
                largura = 28
            elif col_name in {'Gestor', 'Supervisor'}:
                largura = 30
            elif col_name in {'Escala Oficial', 'Batidas do Dia'}:
                largura = 24
            elif col_name in {'Entrada Prevista', 'Entrada Real', 'Saída Prevista', 'Saída Real', 'Saldo BH'}:
                largura = 16
            elif col_name in {'Data'}:
                largura = 12
            ws_detalhe.set_column(col_num, col_num, largura)
        for row_num in range(1, len(df_detalhe) + 1):
            is_alt = row_num % 2 == 0
            for col_num in range(len(df_detalhe.columns)):
                valor = valor_excel_seguro(df_detalhe.iloc[row_num - 1, col_num])
                if is_alt:
                    fmt = cell_alt_fmt
                else:
                    fmt = cell_fmt
                ws_detalhe.write(row_num, col_num, valor, fmt)
        ws_detalhe.freeze_panes(1, 0)
        ws_detalhe.autofilter(0, 0, len(df_detalhe), len(df_detalhe.columns) - 1)
        ws_detalhe.set_row(0, 30)

    return buffer_saida.getvalue()


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


def gerar_excel_ocorrencia(df: pd.DataFrame, config: Dict, col_nome: str, col_cargo: str, col_depto: str, col_data_adm: str, col_ocorrencia: str, col_justificativa: str, col_data: str, col_marcacoes: str = None, col_atraso_calc: str = None,
                           mapa_colaboradores: Dict = None, mapa_supervisores: Dict = None,
                           medidas_dict: Dict = None, demissoes_dict: Dict = None, medidas_atraso_dict: Dict = None) -> Tuple[bytes, str, int]:
    df_detalhe, df_ranking = processar_ocorrencia(df, config['ocorrencia'], config['justificativa'], col_nome, col_cargo, col_depto, col_data_adm, col_ocorrencia, col_justificativa, col_data, col_marcacoes, col_atraso_calc, mapa_colaboradores, mapa_supervisores, medidas_dict, demissoes_dict, medidas_atraso_dict)
    if len(df_detalhe) == 0:
        return None, config.get('arquivo', '') + '.xlsx', 0
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        col_widths_extra = None
        if 'Medida Disciplinar' in df_detalhe.columns or 'Projeção Desligamento' in df_detalhe.columns:
            col_widths_extra = {'Medida Disciplinar': 30, 'Projeção Desligamento': 28}
        gerar_planilha_ocorrencia(df_detalhe, df_ranking, config['nome'], writer, col_widths_extra)
    return excel_buffer.getvalue(), f"{config['arquivo']}.xlsx", len(df_detalhe)


def gerar_pasta_ocorrencia(df: pd.DataFrame, config: Dict, col_nome: str, col_cargo: str, col_depto: str, col_data_adm: str, col_ocorrencia: str, col_justificativa: str, col_data: str, col_marcacoes: str = None, col_atraso_calc: str = None,
                           mapa_colaboradores: Dict = None, mapa_supervisores: Dict = None,
                           medidas_dict: Dict = None, demissoes_dict: Dict = None, medidas_atraso_dict: Dict = None) -> Tuple[Dict[str, bytes], str, int]:
    pasta = config['pasta']
    arquivos = {}
    total_colaboradores = 0
    for justificativa in config['justificativas']:
        config_unica = {'ocorrencia': config['ocorrencia'], 'justificativa': justificativa, 'nome': f"{config['nome']} - {justificativa}", 'arquivo': sanitizar_nome_arquivo(justificativa)}
        excel_bytes, nome_arquivo, qtd = gerar_excel_ocorrencia(df, config_unica, col_nome, col_cargo, col_depto, col_data_adm, col_ocorrencia, col_justificativa, col_data, col_marcacoes, col_atraso_calc, mapa_colaboradores, mapa_supervisores, medidas_dict, demissoes_dict, medidas_atraso_dict)
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

# Uploads opcionais para Medida Disciplinar e Demissões
st.subheader("(Opcional) Arquivos para cruzamento de dados")
col_med, col_dem = st.columns(2)
with col_med:
    f_med = st.file_uploader("Planilha de Medida Disciplinar", type=["xlsx", "xls", "xlsm"], key="medida_disciplinar_ponto_geral")
with col_dem:
    f_dem = st.file_uploader("Planilha de Demissões", type=["xlsx", "xls", "xlsm"], key="demissoes_ponto_geral")

# Processa arquivos opcionais
medidas_dict = {}
medidas_atraso_dict = {}
demissoes_dict = {}
if f_med:
    medidas_dict = processar_medidas(f_med, "FALTA INJUSTIFICADA")
    medidas_atraso_dict = processar_medidas(f_med, "ATRASOS")
if f_dem:
    demissoes_dict = processar_demissoes(f_dem)
if f_med:
    st.info(f"📊 Medida Disciplinar: {len(medidas_dict)} Falta Injustificada + {len(medidas_atraso_dict)} Atraso")
if f_dem and demissoes_dict:
    st.success(f"✅ Demissões: {len(demissoes_dict)} registros")
    with st.expander("🔍 Ver registros de Demissões"):
        for nome, info in list(demissoes_dict.items())[:20]:
            st.write(f"- {nome}: {info['data']} - {info['tipo']}")
elif f_dem and not demissoes_dict:
    st.warning("⚠️ Demissões: Nenhum registro encontrado")

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
    col_marcacoes_atraso = df.columns[24]  # Col Y - Escala + Entrada real (ex: "06:00  06:50")
    col_atraso_calc = df.columns[26]       # Col AA - Cálculo do atraso (ex: "50 Minutos")
    
    st.info(f"D={col_nome} | H={col_cargo} | I={col_depto} | Q={col_data_adm} | Y={col_marcacoes_atraso} | Z={col_ocorrencia} | AA={col_atraso_calc} | AB={col_justificativa} | AM={col_data}")
    
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
                    eb, na, qtd = gerar_excel_ocorrencia(df, config, col_nome, col_cargo, col_depto, col_data_adm, col_ocorrencia, col_justificativa, col_data, col_marcacoes_atraso, col_atraso_calc, mapa_colaboradores, mapa_supervisores, medidas_dict, demissoes_dict, medidas_atraso_dict)
                    resultados_processados.append({'config': config, 'excel_bytes': eb, 'nome_arquivo': na, 'qtd': qtd, 'tipo': 'unica'})
                elif config['tipo'] == 'multiplas_pasta_unica':
                    arquivos = {}; total_colab = 0
                    for item in config['itens']:
                        eb, na, qtd = gerar_excel_ocorrencia(df, item, col_nome, col_cargo, col_depto, col_data_adm, col_ocorrencia, col_justificativa, col_data, col_marcacoes_atraso, col_atraso_calc, mapa_colaboradores, mapa_supervisores, medidas_dict, demissoes_dict, medidas_atraso_dict)
                        if eb is not None: arquivos[na] = eb; total_colab += qtd
                    resultados_processados.append({'config': config, 'arquivos': arquivos, 'pasta': config['pasta'], 'total_colab': total_colab, 'tipo': 'multiplas'})
                else:
                    arquivos, pasta, total_colab = gerar_pasta_ocorrencia(df, config, col_nome, col_cargo, col_depto, col_data_adm, col_ocorrencia, col_justificativa, col_data, col_marcacoes_atraso, col_atraso_calc, mapa_colaboradores, mapa_supervisores, medidas_dict, demissoes_dict, medidas_atraso_dict)
                    resultados_processados.append({'config': config, 'arquivos': arquivos, 'pasta': pasta, 'total_colab': total_colab, 'tipo': 'multiplas'})
                progress_bar.progress((i + 1) / (total + 1))

            status_text.text("Processando: Possiveis alteracoes de escala...")
            col_escala_codigo = localizar_coluna(df, ['EscalaCodigoDescricao', 'JornadaCodigoDescricaoStr'], 22)
            df_resumo_escala, df_detalhe_escala = processar_alteracoes_escala(
                df,
                col_nome,
                col_cargo,
                col_depto,
                col_data_adm,
                col_data,
                col_escala,
                col_marcacoes,
                mapa_colaboradores,
                mapa_supervisores,
                col_escala_codigo=col_escala_codigo,
            )
            excel_escala = None
            if not df_resumo_escala.empty:
                excel_escala = gerar_planilha_alteracoes_escala(df_resumo_escala, df_detalhe_escala)
                resultados_processados.append({'config': {'nome': 'Possiveis alteracoes de escala'}, 'excel_bytes': excel_escala, 'nome_arquivo': 'Possiveis_Alteracoes_Escala.xlsx', 'qtd': len(df_resumo_escala), 'tipo': 'unica'})

            progress_bar.progress(0.96)
            
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

            st.subheader("3. Possiveis alteracoes de escala")
            if df_resumo_escala.empty:
                st.warning("Nenhuma alteracao de escala consistente foi encontrada.")
            else:
                st.success(f"Encontrados {len(df_resumo_escala)} colaboradores com possivel alteracao de escala.")
                st.dataframe(df_resumo_escala, use_container_width=True, hide_index=True)
                st.caption("A planilha de escala ja foi incluida no mesmo ZIP gerado acima.")
            
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
    st.info(" 🌐 Carregue os arquivos para começar.")