"""
Módulo de processamento de CSV de colaboradores
Extrai dados, calcula turno e supervisor
"""

import pandas as pd
import re
from typing import Tuple, Dict, List


def determinar_turno(jornada: str) -> str:
    """
    Determina o turno baseado no horário inicial da jornada.
    
    TURNO 1: 06:00, 07:00, 08:00, 09:00
    TURNO 2: 10:00, 11:00, 12:00, 13:00, 13:40, 14:00
    TURNO 3: 20:00, 21:00, 22:00
    
    Args:
        jornada: String contendo a jornada (ex: "06:00 - 14:00" ou "06:00 10:00 11:00 14:20 - 6x1")
    
    Returns:
        String com o turno (TURNO 1, TURNO 2, TURNO 3 ou "Indeterminado")
    """
    # Debug: verifica o que está chegando
    if pd.isna(jornada) or jornada == "":
        return "Indeterminado"
    
    try:
        jornada_str = str(jornada).strip()
        
        # Extrai TODOS os horários no formato HH:MM usando regex
        # Procura por padrão: 2 dígitos : 2 dígitos
        horarios_encontrados = re.findall(r'\b\d{2}:\d{2}\b', jornada_str)
        
        if not horarios_encontrados:
            return "Indeterminado"
        
        # Pega o PRIMEIRO horário encontrado
        primeiro_horario = horarios_encontrados[0]
        
        # Horários TURNO 1
        turno_1_horarios = ["06:00", "07:00", "08:00", "09:00"]
        if primeiro_horario in turno_1_horarios:
            return "TURNO 1"
        
        # Horários TURNO 2
        turno_2_horarios = ["10:00", "11:00", "12:00", "13:00", "13:40", "14:00"]
        if primeiro_horario in turno_2_horarios:
            return "TURNO 2"
        
        # Horários TURNO 3
        turno_3_horarios = ["20:00", "21:00", "22:00"]
        if primeiro_horario in turno_3_horarios:
            return "TURNO 3"
        
        return "Indeterminado"
    
    except Exception as e:
        return f"Erro: {str(e)}"


def extrair_tabela_supervisores(df: pd.DataFrame, mapa_colunas: Dict) -> Dict[str, str]:
    """
    PASSO 1: Filtra todos os ENCARREGADOS (I, II, III)
    e guarda {nome_encarregado → gestor_dele}
    
    Isso cria a lista de supervisores que será usada depois.
    
    Args:
        df: DataFrame com todos os dados do CSV
        mapa_colunas: Dicionário com mapeamento de colunas
    
    Returns:
        Dicionário {ENCARREGADO_UPPER: GESTOR}
    """
    
    col_cargo = mapa_colunas.get('cargo')
    col_colaborador = mapa_colunas.get('colaborador')
    col_gestor = mapa_colunas.get('gestor')
    
    tabela_supervisores = {}
    
    # Filtra ENCARREGADOS
    cargos_encarregados = ["ENCARREGADO I", "ENCARREGADO II", "ENCARREGADO III"]
    cargos_upper = [c.upper().strip() for c in cargos_encarregados]
    
    mask = df[col_cargo].astype(str).str.upper().str.strip().isin(cargos_upper)
    df_encarregados = df[mask]
    
    # Guarda cada encarregado com seu gestor
    for idx, row in df_encarregados.iterrows():
        nome_encarregado = str(row[col_colaborador]).strip().upper()
        nome_gestor = str(row[col_gestor]).strip() if pd.notna(row[col_gestor]) else ""
        
        if nome_encarregado and nome_gestor:
            tabela_supervisores[nome_encarregado] = nome_gestor
    
    return tabela_supervisores


def encontrar_supervisor(nome_gestor: str, tabela_supervisores: Dict[str, str]) -> str:
    """
    Busca o supervisor usando a tabela de ENCARREGADOS.
    
    Args:
        nome_gestor: Nome do gestor
        tabela_supervisores: Dict {ENCARREGADO: GESTOR}
    
    Returns:
        Nome do supervisor ou "Não encontrado"
    """
    if pd.isna(nome_gestor) or nome_gestor == "":
        return "Não encontrado"
    
    nome_upper = str(nome_gestor).upper().strip()
    
    return tabela_supervisores.get(nome_upper, "Não encontrado")


def processar_csv_colaboradores(
    df: pd.DataFrame,
    cargos_filtro: List[str] = None,
    mapa_colunas: Dict = None
) -> Tuple[pd.DataFrame, Dict]:
    """
    Processa o CSV de colaboradores, extraindo dados e calculando turno e supervisor.
    
    Args:
        df: DataFrame com os dados do CSV
        cargos_filtro: Lista de cargos a filtrar. Se None, usa os padrões.
        mapa_colunas: Dicionário com mapeamento de colunas encontradas
    
    Returns:
        Tuple contendo:
        - DataFrame processado (apenas as 9 colunas solicitadas)
        - Dict com informações de processamento
    """
    
    # Cargos padrão se não fornecidos
    if cargos_filtro is None:
        cargos_filtro = [
            "AUXILIAR DEPOSITO I",
            "AUXILIAR DEPOSITO II",
            "AUXILIAR DEPOSITO III",
            "OPERADOR EMPILHADEIRA"
        ]
    
    info_processamento = {
        "total_linhas_original": len(df),
        "total_linhas_processado": 0,
        "linhas_filtradas": 0,
        "cargos_nao_encontrados": 0,
        "erros": []
    }
    
    try:
        # Se não recebeu mapa de colunas, tenta detectar
        if mapa_colunas is None:
            _, _, mapa_colunas = validar_csv(df)
        
        # Verifica se conseguiu encontrar as colunas
        if not mapa_colunas or len(mapa_colunas) < 7:
            raise ValueError("Não foi possível identificar todas as colunas necessárias")
        
        # Pega os nomes reais das colunas
        col_colaborador = mapa_colunas.get('colaborador')
        col_cargo = mapa_colunas.get('cargo')
        col_situacao = mapa_colunas.get('situacao')
        col_cc = mapa_colunas.get('cc')
        col_gestor = mapa_colunas.get('gestor')
        col_unidade = mapa_colunas.get('unidade')
        col_jornada = mapa_colunas.get('jornada')
        
        # Filtra por cargo (normaliza para maiúsculas para comparação)
        cargos_filtro_upper = [c.upper().strip() for c in cargos_filtro]
        mask = df[col_cargo].astype(str).str.upper().str.strip().isin(cargos_filtro_upper)
        df_filtrado = df[mask].copy()
        
        # IMPORTANTE: Reset dos índices para evitar desalinhamento
        df_filtrado = df_filtrado.reset_index(drop=True)
        
        info_processamento["linhas_filtradas"] = len(df) - len(df_filtrado)
        info_processamento["cargos_nao_encontrados"] = info_processamento["linhas_filtradas"]
        
        # Cria novo DataFrame com apenas as 9 colunas solicitadas
        # Usando explicitamente os índices para garantir alinhamento
        df_resultado = pd.DataFrame(index=range(len(df_filtrado)))
        
        # PRIMEIRO: Extrai tabela de supervisores (antes de filtrar)
        tabela_supervisores = extrair_tabela_supervisores(df, mapa_colunas)
        
        # Adiciona as colunas na ordem solicitada
        df_resultado["Colaborador"] = df_filtrado[col_colaborador].values
        df_resultado["Cargo"] = df_filtrado[col_cargo].values
        df_resultado["Descrição Situação"] = df_filtrado[col_situacao].values
        df_resultado["Descrição CC"] = df_filtrado[col_cc].values
        df_resultado["Nome Gestor"] = df_filtrado[col_gestor].values
        
        # SUPERVISOR - Usa a tabela de supervisores extraída
        supervisores = []
        for gestor in df_filtrado[col_gestor].values:
            supervisor = encontrar_supervisor(gestor, tabela_supervisores)
            supervisores.append(supervisor)
        
        df_resultado["Supervisor"] = supervisores
        
        df_resultado["Descrição da Unidade Organizacional"] = df_filtrado[col_unidade].values
        
        # TURNO - Calcula para cada jornada
        # Aplica com verificação de erros
        try:
            turno_result = df_filtrado[col_jornada].apply(determinar_turno)
            df_resultado["Turno"] = turno_result
        except Exception as e:
            # Se houver erro, preenche com "Erro" para debug
            df_resultado["Turno"] = "Erro: " + str(e)
        
        df_resultado["Jornada"] = df_filtrado[col_jornada].values
        
        info_processamento["total_linhas_processado"] = len(df_resultado)
        
        return df_resultado, info_processamento
        
    except Exception as e:
        info_processamento["erros"].append(str(e))
        raise


def validar_csv(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    """
    Valida se o CSV tem a estrutura esperada.
    Suporta diferentes nomes de coluna.
    
    Args:
        df: DataFrame a validar
    
    Returns:
        Tuple (é_válido, lista_de_erros, dicionário_de_colunas_encontradas)
    """
    erros = []
    
    # Colunas que podem estar presentes (diferentes nomes)
    colunas_obrigatorias = {
        'colaborador': ['Colaborador', 'COLABORADOR', 'Nome', 'NOME'],
        'cargo': ['Cargo', 'CARGO'],
        'situacao': ['Descri??o Situa??o', 'Descrição Situação', 'Situação', 'SITUACAO'],
        'cc': ['Descri??o CC', 'Descrição CC', 'CC', 'Codigo CC', 'C?digo CC'],
        'gestor': ['Nome Gestor', 'GESTOR', 'Gestor', 'Matr?cula Gestor'],
        'unidade': ['Descri??o da Unidade Organizacional', 'Descrição da Unidade Organizacional', 'Unidade', 'UNIDADE'],
        'jornada': ['Jornada', 'JORNADA', 'Codigo Jornada', 'C?digo Jornada']
    }
    
    colunas_encontradas = {}
    for chave, nomes_possiveis in colunas_obrigatorias.items():
        encontrado = False
        for nome in nomes_possiveis:
            if nome in df.columns:
                colunas_encontradas[chave] = nome
                encontrado = True
                break
        
        # Se não encontrou de forma exata, tenta busca parcial
        if not encontrado:
            for col in df.columns:
                # Remove caracteres especiais para comparação
                col_clean = col.lower().replace('?', '').replace('ç', 'c').replace('ã', 'a').replace('á', 'a').replace('é', 'e').replace('ó', 'o')
                chave_clean = chave.lower()
                
                # Procura por palavras-chave
                palavras_chave = {
                    'colaborador': ['colaborador'],
                    'cargo': ['cargo'],
                    'situacao': ['situacao', 'situação'],
                    'cc': ['cc', 'centro', 'custo'],
                    'gestor': ['gestor'],
                    'unidade': ['unidade', 'organizacional'],
                    'jornada': ['jornada']
                }
                
                if chave in palavras_chave:
                    for palavra in palavras_chave[chave]:
                        if palavra in col_clean:
                            colunas_encontradas[chave] = col
                            encontrado = True
                            break
            
            if not encontrado:
                erros.append(f"Coluna '{chave}' não encontrada. Procure por: {', '.join(nomes_possiveis)}")
    
    if len(df) == 0:
        erros.append("CSV vazio")
    
    # Retorna as colunas encontradas para uso posterior
    return len(erros) == 0, erros, colunas_encontradas
