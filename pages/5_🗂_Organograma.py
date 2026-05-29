"""
Página de Organograma Interativo
Gera um organograma hierárquico com caixas clicáveis.
"""

import json
import traceback
import unicodedata

import pandas as pd
import streamlit as st


st.set_page_config(page_title="ORGANOGRAMA", layout="wide")

st.header("🗂 ORGANOGRAMA")
st.write("Organograma interativo baseado em hierarquia de gestores e colaboradores. Clique nas caixas para expandir/colapsar.")


def limpar_texto(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    return "" if texto.lower() == "nan" else texto


def limpar_numero(valor):
    """Remove .0 from floats that are actually integers"""
    if pd.isna(valor):
        return ""
    try:
        num = float(valor)
        if num == int(num):
            return str(int(num))
    except (ValueError, TypeError):
        pass
    return limpar_texto(valor)


def normalizar_chave(valor):
    texto = limpar_texto(valor).lower()
    nfd = unicodedata.normalize("NFD", texto)
    return "".join(c for c in nfd if unicodedata.category(c) != "Mn")


def achar_coluna_por_keywords(df, keywords_list):
    cols = list(df.columns)
    norm_map = {c: normalizar_chave(c) for c in cols}

    # First pass: try exact matches (keyword is exactly the normalized column)
    for kw in keywords_list:
        kw_norm = normalizar_chave(kw)
        for c in cols:
            if norm_map[c] == kw_norm:
                return c
    
    # Second pass: try substring matches, but prefer longer/more specific keywords first
    # Sort keywords by length descending (longer keywords are more specific)
    sorted_keywords = sorted(keywords_list, key=len, reverse=True)
    for kw in sorted_keywords:
        kw_norm = normalizar_chave(kw)
        for c in cols:
            if kw_norm in norm_map[c]:
                return c
    
    return None


def ler_planilha(uploaded_file):
    if uploaded_file.name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)

    encodings = ["utf-8", "latin-1", "iso-8859-1", "cp1252", "utf-16"]
    separators = [";", ",", "\t", "|"]

    for encoding in encodings:
        for separator in separators:
            for skiprows in (0, 1):
                try:
                    uploaded_file.seek(0)
                    df_try = pd.read_csv(uploaded_file, encoding=encoding, sep=separator, skiprows=skiprows)
                    if len(df_try.columns) > 1:
                        return df_try
                except Exception:
                    pass

    uploaded_file.seek(0)
    return pd.read_csv(uploaded_file, encoding="latin-1", sep=";")


uploaded = st.file_uploader(
    "📤 Envie o CSV/XLSX com os dados",
    type=["csv", "xlsx", "xlsm"],
    key="org_file_uploader"
)

if uploaded is None:
    st.info("Envie a planilha para gerar o organograma")
    st.stop()

try:
    df = ler_planilha(uploaded)

    st.success(f"Arquivo carregado: {uploaded.name} — {len(df)} linhas")

    filtro_busca = normalizar_chave(
        st.text_input(
            "Buscar área, gestor ou colaborador",
            value="",
            placeholder="Ex.: RJ GER OPERACOES M&A",
        )
    )

    col_colaborador = achar_coluna_por_keywords(df, ["colaborador", "nome social", "nome"])
    col_matricula = achar_coluna_por_keywords(df, ["matricula"])
    col_uo_codigo = achar_coluna_por_keywords(df, ["codigo da unidade organizacional", "codigo uo", "cod uo"])
    col_uo_descricao = achar_coluna_por_keywords(df, ["descricao da unidade organizacional", "descricao unidade", "unidade organizacional"])
    col_gestor = achar_coluna_por_keywords(df, ["nome gestor", "gestor", "manager"])
    col_uo_nivel_codigo = achar_coluna_por_keywords(df, ["codigo nivel uo", "codigo_nivel_uo"])
    col_uo_nivel_descricao = achar_coluna_por_keywords(df, ["nivel uo", "descricao nivel uo"])
    col_cargo = achar_coluna_por_keywords(df, ["descricao do cargo", "cargo"])
    col_matricula_gestor = achar_coluna_por_keywords(df, ["matricula gestor"])

    st.write(
        {
            "colaborador": col_colaborador,
            "matricula": col_matricula,
            "codigo_uo": col_uo_codigo,
            "descricao_uo": col_uo_descricao,
            "gestor": col_gestor,
            "nivel_codigo": col_uo_nivel_codigo,
            "nivel_descricao": col_uo_nivel_descricao,
            "cargo": col_cargo,
            "matricula_gestor": col_matricula_gestor,
        }
    )

    if col_colaborador is None or col_gestor is None:
        st.error("Não foi possível detectar as colunas de colaborador e gestor.")
        st.stop()

    df = df.copy()
    df[col_colaborador] = df[col_colaborador].apply(limpar_texto)
    if col_matricula:
        df[col_matricula] = df[col_matricula].apply(limpar_texto)
    if col_uo_codigo:
        df[col_uo_codigo] = df[col_uo_codigo].apply(limpar_texto)
    if col_uo_descricao:
        df[col_uo_descricao] = df[col_uo_descricao].apply(limpar_texto)
    if col_uo_nivel_codigo:
        df[col_uo_nivel_codigo] = df[col_uo_nivel_codigo].apply(limpar_texto)
    if col_uo_nivel_descricao:
        df[col_uo_nivel_descricao] = df[col_uo_nivel_descricao].apply(limpar_texto)
    if col_matricula_gestor:
        df[col_matricula_gestor] = df[col_matricula_gestor].apply(limpar_numero)

    df["__colaborador_norm"] = df[col_colaborador].apply(normalizar_chave)
    df["__gestor_norm"] = df[col_gestor].apply(normalizar_chave)
    df["__matricula_norm"] = df[col_matricula].apply(normalizar_chave) if col_matricula else ""

    if col_uo_nivel_codigo:
        nivel_extraido = df[col_uo_nivel_codigo].astype(str).str.extract(r"(\d+)")[0]
        df["__nivel_num"] = pd.to_numeric(nivel_extraido, errors="coerce")
    else:
        df["__nivel_num"] = pd.NA

    def nivel_num_from_text(valor):
        texto = limpar_texto(valor)
        if not texto:
            return 999999
        try:
            return int(str(texto).strip())
        except (ValueError, TypeError):
            try:
                extraido = pd.Series([texto]).astype(str).str.extract(r"(\d+)")[0].iloc[0]
                return int(extraido) if limpar_texto(extraido) else 999999
            except Exception:
                return 999999

    def pessoa_key_from_row(row, idx):
        if col_matricula and limpar_texto(row[col_matricula]):
            return normalizar_chave(row[col_matricula])
        return f"{normalizar_chave(row[col_colaborador])}_{idx}"

    person_rows = {}
    for idx, row in df.iterrows():
        nome = limpar_texto(row[col_colaborador])
        if not nome:
            continue
        key = pessoa_key_from_row(row, idx)
        nivel_num = row["__nivel_num"] if pd.notna(row["__nivel_num"]) else 999999
        current = person_rows.get(key)
        if current is None or nivel_num < current["nivel_num"]:
            person_rows[key] = {"row": row, "nome": nome, "key": key, "nivel_num": nivel_num}

    nome_to_key = {normalizar_chave(pessoa["nome"]): key for key, pessoa in person_rows.items()}
    gestores_norm = set(df["__colaborador_norm"].tolist())

    def build_node(person_key, stack=None):
        stack = set() if stack is None else set(stack)
        if person_key in stack:
            return None

        pessoa = person_rows.get(person_key)
        if pessoa is None:
            return None

        stack.add(person_key)
        row = pessoa["row"]
        nome = pessoa["nome"]
        gestor_nome = limpar_texto(row[col_gestor])
        uo_desc = limpar_texto(row[col_uo_descricao]) if col_uo_descricao else limpar_texto(row[col_uo_codigo])
        uo_code = limpar_texto(row[col_uo_codigo]) if col_uo_codigo else ""
        nivel_codigo = limpar_texto(row[col_uo_nivel_codigo]) if col_uo_nivel_codigo else ""
        nivel_descricao = limpar_texto(row[col_uo_nivel_descricao]) if col_uo_nivel_descricao else ""
        cargo = limpar_texto(row[col_cargo]) if col_cargo else ""

        cargo_norm = normalizar_chave(cargo)
        cargo_indica_gestao = any(
            keyword in cargo_norm
            for keyword in [
                "gestor",
                "gerente",
                "coordenador",
                "coordenacao",
                "coord",
                "encarregado",
                "supervisor",
                "sup",
                "lider",
                "líder",
                "lideranca",
                "liderança",
            ]
        )

        children = []
        subordinates = df[df["__gestor_norm"] == normalizar_chave(nome)]
        if not subordinates.empty:
            sort_cols = ["__nivel_num"]
            if col_colaborador:
                sort_cols.append(col_colaborador)
            subordinates = subordinates.sort_values(by=sort_cols, na_position="last")

            seen = set()
            gestor_children = []
            leaf_people = []
            
            for idx, child_row in subordinates.iterrows():
                child_name = limpar_texto(child_row[col_colaborador])
                if not child_name:
                    continue
                if normalizar_chave(child_name) == normalizar_chave(nome):
                    continue

                child_key = pessoa_key_from_row(child_row, idx)
                if child_key in seen:
                    continue
                seen.add(child_key)

                has_subordinates = not df[df["__gestor_norm"] == normalizar_chave(child_name)].empty
                
                if has_subordinates:
                    child_node = build_node(child_key, stack)
                    if child_node is not None:
                        gestor_children.append(child_node)
                else:
                    child_uo_desc = limpar_texto(child_row[col_uo_descricao]) if col_uo_descricao else limpar_texto(child_row[col_uo_codigo])
                    child_uo_code = limpar_texto(child_row[col_uo_codigo]) if col_uo_codigo else ""
                    leaf_people.append({
                        "child_name": child_name,
                        "child_key": child_key,
                        "matricula": limpar_texto(child_row[col_matricula]) if col_matricula else "",
                        "cargo": limpar_texto(child_row[col_cargo]) if col_cargo else "",
                        "uo_desc": child_uo_desc,
                        "uo_code": child_uo_code,
                        "codigo_nivel_uo": limpar_texto(child_row[col_uo_nivel_codigo]) if col_uo_nivel_codigo else "",
                        "gestor": nome,
                    })
            
            children.extend(gestor_children)
            
            if leaf_people:
                # Group leaf people by their UO description to avoid mixing different UOs
                groups = {}
                for p in leaf_people:
                    key_uo = p.get('uo_desc') or p.get('uo_code') or 'COLABORADORES'
                    groups.setdefault(key_uo, []).append(p)

                for i, (group_uo_desc, group_people) in enumerate(groups.items()):
                    agg_uo_code = group_people[0].get('uo_code', '') if group_people else ''
                    agg_level_values = [nivel_num_from_text(p.get('codigo_nivel_uo', '')) for p in group_people]
                    agg_level_values = [v for v in agg_level_values if v != 999999]
                    agg_level_code = str(min(agg_level_values)) if agg_level_values else nivel_codigo
                    agg_node = {
                        "id": f"agg_{person_key}_{i}",
                        "name": group_uo_desc,
                        "title": group_uo_desc,
                        "subtitle": f"{len(group_people)} pessoas",
                        "type": "uo",
                        "is_aggregated": True,
                        "aggregated_people": group_people,
                        "info": {
                            "gestor": nome,
                            "uo": group_uo_desc,
                            "codigo_uo": agg_uo_code,
                            "codigo_nivel_uo": agg_level_code,
                            "nivel_uo": nivel_descricao,
                        },
                    }
                    # Adicionar agregada diretamente como irmã dos gestores
                    children.append(agg_node)

        node = {
            "id": person_key,
            "name": uo_desc or nome,
            "title": uo_desc or nome,
            "subtitle": nome,
            "type": "gestor" if (children or cargo_indica_gestao) else "colaborador",
            "info": {
                "gestor": gestor_nome,
                "uo": uo_desc,
                "codigo_uo": uo_code,
                "codigo_nivel_uo": nivel_codigo,
                "nivel_uo": nivel_descricao,
                "cargo": cargo,
            },
        }

        if children:
            node["children"] = children

        return node

    def node_level_sort_value(node):
        info = node.get("info", {}) or {}
        if node.get("is_aggregated"):
            agg_values = [nivel_num_from_text(p.get("codigo_nivel_uo", "")) for p in node.get("aggregated_people", [])]
            agg_values = [v for v in agg_values if v != 999999]
            if agg_values:
                return min(agg_values)
        return nivel_num_from_text(info.get("codigo_nivel_uo") or info.get("codigo_nivel") or "")

    def contar_nos_sort(node):
        return 1 + sum(contar_nos_sort(child) for child in node.get("children", []))

    def sort_node_children_recursive(node):
        children = node.get("children")
        if not children:
            return

        def child_sort_key(item):
            type_rank = {"gestor": 0, "uo": 1, "colaborador": 2}.get(item.get("type", ""), 3)
            return (
                type_rank,
                node_level_sort_value(item),
                -contar_nos_sort(item),
                item.get("subtitle", ""),
                item.get("title", ""),
            )

        children.sort(key=child_sort_key)
        for child in children:
            sort_node_children_recursive(child)

    top_level = []
    for key, pessoa in person_rows.items():
        gestor_norm = normalizar_chave(limpar_texto(pessoa["row"][col_gestor]))
        if not gestor_norm or gestor_norm not in gestores_norm:
            node = build_node(key)
            if node is not None:
                # Include gestores even when they do not have subordinates
                if node.get("children") or node.get("type") == "gestor":
                    top_level.append(node)
    
        # DEBUG: Show which people are top_level and how many children they have
    for node in top_level:
        sort_node_children_recursive(node)

    st.write(f"DEBUG: {len(top_level)} top_level nodes with children")
    for node in top_level[:10]:
        st.write(f"  - {node.get('subtitle', 'UNKNOWN')} ({node.get('name', '')}): {len(node.get('children', []))} children")

    all_people = []
    for idx, row in df.iterrows():
        nome = limpar_texto(row[col_colaborador])
        if not nome:
            continue

        uo_desc = limpar_texto(row[col_uo_descricao]) if col_uo_descricao else limpar_texto(row[col_uo_codigo])
        uo_code = limpar_texto(row[col_uo_codigo]) if col_uo_codigo else ""
        if uo_code:
            uo_key = normalizar_chave(uo_code)
        else:
            uo_key = normalizar_chave(uo_desc)

        all_people.append(
            {
                "nome": nome,
                "matricula": limpar_texto(row[col_matricula]) if col_matricula else "",
                "gestor": limpar_texto(row[col_gestor]),
                "uo": uo_desc,
                "codigo_uo": uo_code,
                "codigo_nivel_uo": limpar_texto(row[col_uo_nivel_codigo]) if col_uo_nivel_codigo else "",
                "nivel_uo": limpar_texto(row[col_uo_nivel_descricao]) if col_uo_nivel_descricao else "",
                "cargo": limpar_texto(row[col_cargo]) if col_cargo else "",
                "matricula_gestor": limpar_numero(row[col_matricula_gestor]) if col_matricula_gestor else "",
                "uo_key": uo_key,
                "colaborador_norm": normalizar_chave(nome),
            }
        )

    def contar_nos(node):
        return 1 + sum(contar_nos(child) for child in node.get("children", []))

    def node_contem_texto(node, texto_busca):
        """Verifica recursivamente se o texto de busca existe em algum lugar da árvore do nó"""
        
        partes = [
            node.get("name", ""),
            node.get("title", ""),
            node.get("subtitle", ""),
        ]
        info = node.get("info", {})
        partes.extend([info.get("gestor", ""), info.get("uo", ""), info.get("codigo_uo", "")])
        
        # Normaliza e verifica o nó atual
        texto_nodo_norm = normalizar_chave(" ".join(partes))
        if texto_busca in texto_nodo_norm:
            return True
        
        # Procura em pessoas agregadas (nós de agregação têm isso ao invés de children)
        aggregated_people = node.get("aggregated_people", [])
        
        for person in aggregated_people:
            person_texto = normalizar_chave(" ".join([
                person.get("child_name", ""),
                person.get("gestor", ""),
                person.get("uo_desc", ""),
                person.get("uo_code", ""),
                person.get("cargo", ""),
            ]))
            if texto_busca in person_texto:
                return True
        
        # Procura recursivamente nos filhos
        for child in node.get("children", []):
            if node_contem_texto(child, texto_busca):
                return True
        
        return False

    def node_texto_busca(node):
        """Retorna texto normalizado do nó (sem recursão, para compatibilidade)"""
        partes = [
            node.get("name", ""),
            node.get("title", ""),
            node.get("subtitle", ""),
        ]
        info = node.get("info", {})
        partes.extend([info.get("gestor", ""), info.get("uo", ""), info.get("codigo_uo", "")])
        return normalizar_chave(" ".join(partes))

    candidatos = top_level
    if filtro_busca:
        filtrados = [node for node in top_level if node_contem_texto(node, filtro_busca)]
        if filtrados:
            candidatos = filtrados

    if not candidatos:
        st.error("Não foi possível montar uma raiz para o organograma com os dados enviados.")
        st.stop()

    candidatos = sorted(
        candidatos,
        key=lambda item: (
            node_level_sort_value(item),
            -contar_nos(item),
            item.get("subtitle", ""),
            item.get("title", ""),
        ),
    )

    for node in candidatos:
        sort_node_children_recursive(node)

    def encaixar_candidatos_por_nivel(nodes):
        roots = []
        stack = []

        for node in nodes:
            node_level = node_level_sort_value(node)
            while stack and node_level_sort_value(stack[-1]) >= node_level:
                stack.pop()

            if stack:
                parent = stack[-1]
                parent.setdefault("children", []).append(node)
                # mark that this child was attached by level-based nesting (not by actual gestor link)
                node.setdefault("info", {})["synthetic_parent"] = True
                sort_node_children_recursive(parent)
            else:
                roots.append(node)

            stack.append(node)

        return roots

    candidatos = encaixar_candidatos_por_nivel(candidatos)

    # If multiple top-level candidates exist, create a synthetic root so all managerless boxes are visible
    if len(candidatos) > 1:
        tree_data = {
            "id": "root_all",
            "name": "ORGANOGRAMA",
            "title": "ORGANOGRAMA",
            "subtitle": "",
            "type": "root",
            "children": candidatos,
        }
    else:
        tree_data = candidatos[0]
        tree_data["type"] = "root"

    tree_json = json.dumps(tree_data, ensure_ascii=False)
    people_json = json.dumps(all_people, ensure_ascii=False)

    html_template = """<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <script src="https://d3js.org/d3.v7.min.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        html, body { width: 100%; height: 100%; font-family: Arial, sans-serif; background: white; }
        .search-bar { position: fixed; top: 0; left: 0; right: 0; background: #2c3e50; padding: 10px 20px; z-index: 100; height: 60px; display: flex; align-items: flex-start; gap: 15px; overflow: visible; }
        .search-panel { position: relative; width: 100%; max-width: 900px; }
        .search-input { width: 100%; padding: 8px 12px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; }
        .search-results { display: none; position: absolute; top: calc(100% + 6px); left: 0; right: 0; background: #34495e; padding: 8px 12px; font-size: 11px; color: #ecf0f1; border-radius: 4px; max-height: 240px; overflow-y: auto; box-shadow: 0 10px 24px rgba(0,0,0,0.22); }
        .search-results.open { display: block; }
        .search-results strong { color: #fff; }
        .search-results .search-result-btn { width: 100%; text-align: left; display: block; background: transparent; border: 0; padding: 6px 0; color: inherit; cursor: pointer; }
        .search-results .search-result-btn:hover { color: #ffffff; text-decoration: underline; }
        .search-results li { margin: 4px 0; }
        .search-results small { color: #bdc3c7; }
        .layout { position: fixed; top: 60px; left: 0; right: 0; bottom: 0; width: 100%; height: calc(100vh - 60px); }
        .left-rail { position: absolute; top: 0; left: 0; bottom: 36px; width: 36px; background: white; border-right: 1px solid #e6e6e6; border-bottom: 1px solid #e6e6e6; display: flex; align-items: center; justify-content: center; cursor: pointer; z-index: 10; }
        .left-rail span { transform: rotate(-90deg); transform-origin: center; font-size: 14px; font-weight: 800; color: #3a3a3a; letter-spacing: .06em; white-space: nowrap; }
        .info-drawer { position: absolute; top: 0; left: 36px; bottom: 36px; width: 0; overflow: hidden; background: #f7f8fa; border-right: 1px solid #d9d9d9; transition: width .18s ease; z-index: 9; }
        .info-drawer.open { width: 320px; }
        .info-drawer-inner { width: 320px; height: 100%; display: flex; flex-direction: column; overflow: auto; }
        .sidebar-header { padding: 14px; background: white; border-bottom: 1px solid #e5e7eb; flex-shrink: 0; }
        .sidebar-title { font-size: 15px; font-weight: 800; color: #111827; }
        .sidebar-subtitle { color: #6b7280; font-size: 12px; margin-top: 4px; }
        .sidebar-section { padding: 8px 10px; border-bottom: 1px solid #e8eaee; }
        .sidebar-footer { padding: 8px 10px 10px; border-top: 1px solid #e8eaee; background: #f7f8fa; flex-shrink: 0; }
        .sidebar-section-title { font-size: 10px; color: #374151; text-transform: uppercase; letter-spacing: .08em; font-weight: 800; margin-bottom: 6px; }
        .sidebar-cards { display: grid; gap: 4px; }
        .sidebar-card { background: white; border: 1px solid #e4e4e4; border-radius: 8px; padding: 6px 8px; }
        .sidebar-card-label { font-size: 9px; color: #6b7280; text-transform: uppercase; letter-spacing: .04em; margin-bottom: 2px; }
        .sidebar-card-value { font-size: 11px; font-weight: 700; color: #111827; word-break: break-word; line-height: 1.15; }
        .sidebar-actions { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 10px; }
        .people-list-panel { display: none; flex-direction: column; gap: 10px; min-height: 0; }
        .people-list-panel.open { display: flex; }
        .people-list-panel.minimized .people-list-body { display: none; }
        .people-list-header { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
        .people-list-title { font-size: 12px; font-weight: 800; color: #111827; }
        .people-list-badge { font-size: 11px; color: #6b7280; background: #eef2f7; border-radius: 999px; padding: 3px 8px; }
        .people-list-body { max-height: calc(100vh - 360px); overflow: auto; background: white; border: 1px solid #e4e4e4; border-radius: 8px; }
        .bottom-rail { position: absolute; left: 0; right: 0; bottom: 0; height: 36px; background: white; border-top: 1px solid #e6e6e6; border-right: 1px solid #e6e6e6; display: flex; align-items: center; padding-left: 12px; z-index: 10; font-size: 14px; font-weight: 800; color: #3a3a3a; letter-spacing: .03em; cursor: pointer; }
        #svg-container { position: absolute; top: 0; left: 36px; right: 0; bottom: 36px; background: white; border: none; overflow: hidden; }
        .people-bottom-panel { position: fixed; left: 36px; right: 0; bottom: 36px; height: 0; overflow: auto; background: white; border-top: 1px solid #e6e6e6; box-shadow: 0 -4px 14px rgba(0,0,0,0.08); transition: height .18s ease; z-index: 9999; }
        .people-bottom-panel.open { height: 160px; }
        .node { cursor: pointer; }
        .node rect { stroke: #333; stroke-width: 1.5px; rx: 8px; ry: 8px; }
        .node text { pointer-events: none; text-anchor: middle; font-weight: 500; }
        .link { fill: none; stroke: #999; stroke-width: 1.5px; opacity: 0.6; }
        .node.root rect { fill: #2c3e50; }
        .node.gestor rect { fill: #3498db; }
        .node.uo rect { fill: #2ecc71; }
        .node.colaborador rect { fill: #e74c3c; }
        .node.root .title { fill: white; font-weight: bold; font-size: 11px; }
        .node.root .subtitle { fill: #dfe6e9; font-size: 9px; }
        .node.gestor .title { fill: white; font-weight: bold; font-size: 10px; }
        .node.gestor .subtitle { fill: white; font-size: 9px; }
        .node.uo .title { fill: white; font-weight: bold; font-size: 9px; }
        .node.uo .subtitle { fill: white; font-size: 8px; }
        .node.colaborador .title { fill: white; font-weight: bold; font-size: 8px; }
        .node.colaborador .subtitle { fill: white; font-size: 7px; }
        .tooltip { position: fixed; background: rgba(0, 0, 0, 0.9); color: white; padding: 8px 12px; border-radius: 4px; font-size: 11px; pointer-events: none; display: none; z-index: 1000; }
        table.people-table { width: 100%; border-collapse: collapse; font-size: 12px; }
        table.people-table th, table.people-table td { padding: 8px 10px; border-bottom: 1px solid #eee; text-align: left; vertical-align: top; white-space: nowrap; }
        table.people-table th { position: sticky; top: 0; background: #f3f5f7; z-index: 1; font-size: 11px; text-transform: uppercase; letter-spacing: .04em; }
        .action-btn { border: 1px solid #7d7d7d; background: white; color: #222; border-radius: 6px; padding: 8px 12px; cursor: pointer; font-weight: 700; font-size: 12px; }
        .action-btn.primary { background: #2c3e50; border-color: #2c3e50; color: white; }
        .action-btn:hover { opacity: 0.92; }
    </style>
</head>
<body>
    <div class="search-bar">
        <div class="search-panel">
            <input type="text" class="search-input" id="search-input" placeholder="Pesquise por nome de colaborador, gestor, UO..." />
            <div class="search-results" id="search-results"></div>
        </div>
    </div>
    <div class="layout">
        <div class="left-rail" id="info-rail" title="Abrir informações"><span>INFORMACOES</span></div>
        <div class="info-drawer" id="info-drawer">
            <div class="info-drawer-inner">
                <div class="sidebar-header">
                    <div class="sidebar-title">INFORMACOES</div>
                    <div class="sidebar-subtitle" id="details-subtitle">Selecione uma caixa para ver os dados</div>
                </div>
                <div class="sidebar-section">
                    <div class="sidebar-section-title">Dados da selecao</div>
                    <div id="details-content" class="sidebar-cards">
                        <div class="sidebar-card"><div class="sidebar-card-value">Nenhum item selecionado.</div></div>
                    </div>
                </div>
                <div class="sidebar-section people-list-panel minimized" id="people-list-panel">
                    <div class="people-list-header">
                        <div class="people-list-title">Lista de Pessoas</div>
                        <div class="people-list-badge" id="people-count-badge">0</div>
                    </div>
                    <div class="people-list-body" id="people-table-slot"></div>
                </div>
                <div class="sidebar-footer">
                    <div class="sidebar-actions">
                        <!-- Botões movidos para o rail inferior -->
                    </div>
                </div>
            </div>
        </div>
        <div id="svg-container"></div>
        <div class="bottom-rail" id="people-rail">LISTA DE PESSOAS</div>
    </div>
    <div class="tooltip" id="tooltip"></div>
    <div id="people-bottom-panel" class="people-bottom-panel"><div id="people-bottom-slot" style="padding:8px 12px 8px;"></div></div>
    <script>
        const data = __TREE_JSON__;
        const allPeople = __PEOPLE_JSON__;
        const width = window.innerWidth - 36;
        const height = window.innerHeight - 96;
        
        const svg = d3.select('#svg-container').append('svg').style('width', '100%').style('height', '100%');
        const g = svg.append('g');
        const zoomBehavior = d3.zoom().on('zoom', (event) => {
            currentTransform = event.transform;
            g.attr('transform', event.transform);
        });

        let selectedNode = null;
        let peopleMode = false;
        let currentPeople = [];
        let currentTransform = d3.zoomIdentity;

        function nodeBox(d) {
            // Responsive sizing based on text length - calculate but don't distort
            const title = (d.data.title || d.data.name || '');
            const subtitle = (d.data.subtitle || '');
            const maxTextLen = Math.max(title.length, subtitle.length);

            // Estimate width needed: ~7px per character + padding
            // Cap max width to prevent huge boxes
            const estimatedWidth = Math.min(460, Math.max(160, maxTextLen * 7 + 40));
            
            // Base heights per type, slightly adjustable for very long subtitles
            let baseHeight = 52;
            if (d.data.type === 'root') baseHeight = 64;
            else if (d.data.type === 'gestor') baseHeight = 64;
            else if (d.data.type === 'uo') baseHeight = 70;
            
            // Add extra height if subtitle is very long
            const height = baseHeight + (subtitle.length > 40 ? 12 : 0);

            return { width: estimatedWidth, height: height };
        }

        function collapse(d) {
            if (d.children) {
                d._children = d.children;
                d._children.forEach(collapse);
                d.children = null;
            }
        }

        function collapseAll(d) {
            if (!d) return;
            if (d.children) {
                d.children.forEach(collapseAll);
                d._children = d.children;
                d.children = null;
            }
            if (d._children) {
                d._children.forEach(collapseAll);
            }
        }

        function expandAncestors(node) {
            let current = node;
            while (current && current.parent) {
                const parent = current.parent;
                if (parent._children) {
                    parent.children = parent._children;
                    parent._children = null;
                }
                current = parent;
            }
        }

        function centerNode(node) {
            if (!node) return;
            const scale = 1.15;
            const x = (width / 2) - (node.x * scale);
            const y = (height / 2) - (node.y * scale);
            const transform = d3.zoomIdentity.translate(x, y).scale(scale);
            svg.transition().duration(450).call(zoomBehavior.transform, transform);
        }

        function resetTreeToNode(node) {
            collapseAll(root);
            expandAncestors(node);
            // Also expand the clicked node itself if it has children
            if (node._children) {
                node.children = node._children;
                node._children = null;
            }
            selectedNode = node;
            renderDetailsForNode(node);
            update(root);
            centerNode(node);
        }

        function normalizeText(value) {
            return String(value || '').toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g, '').trim();
        }

        function renderEmptyDetails(message) {
            document.getElementById('details-content').innerHTML = `<div class="sidebar-card"><div class="sidebar-card-value">${message}</div></div>`;
        }

        function hidePeopleList() {
            peopleMode = false;
            currentPeople = [];
            document.getElementById('people-list-panel').classList.remove('open');
            document.getElementById('people-table-slot').innerHTML = '';
            document.getElementById('people-count-badge').textContent = '0';
        }

        document.getElementById('info-rail').addEventListener('click', () => {
            document.getElementById('info-drawer').classList.toggle('open');
        });

        function renderDetailsForNode(node) {
            selectedNode = node;
            const info = node.data.info || {};
            const title = node.data.title || node.data.name || '';
            const isAggregated = node.data.is_aggregated || false;
            const detailCards = [];

            if (isAggregated) {
                document.getElementById('details-subtitle').textContent = `Colaboradores agregados do gestor: ${node.data.info.gestor || ''}`;
                const aggregatedPeople = (node.data.aggregated_people || []).map(p => ({
                    nome: p.child_name,
                    matricula: p.matricula || '',
                    gestor: p.gestor || info.gestor,
                    uo: p.uo_desc || info.uo,
                    codigo_uo: p.uo_code || info.codigo_uo,
                    nivel: info.nivel,
                    cargo: p.cargo || '',
                }));
                detailCards.push(
                    `<div class="sidebar-card"><div class="sidebar-card-label">Gestor</div><div class="sidebar-card-value">${info.gestor || '-'}</div></div>`,
                    `<div class="sidebar-card"><div class="sidebar-card-label">UO</div><div class="sidebar-card-value">${info.uo || '-'}</div></div>`,
                    `<div class="sidebar-card"><div class="sidebar-card-label">Codigo UO</div><div class="sidebar-card-value">${info.codigo_uo || '-'}</div></div>`,
                    `<div class="sidebar-card"><div class="sidebar-card-label">Nivel</div><div class="sidebar-card-value">${info.nivel || '-'}</div></div>`,
                    `<div class="sidebar-card"><div class="sidebar-card-label">Total de Colaboradores</div><div class="sidebar-card-value" id="uo-count">${aggregatedPeople.length}</div></div>`
                );
                document.getElementById('details-content').innerHTML = detailCards.join('');
                currentPeople = aggregatedPeople;
                document.getElementById('people-count-badge').textContent = String(aggregatedPeople.length);
                // If bottom panel is already open, refresh its contents immediately
                try {
                    const bottomPanel = document.getElementById('people-bottom-panel');
                    if (bottomPanel && bottomPanel.classList.contains('open')) {
                        renderPeopleTable(currentPeople);
                        document.getElementById('people-bottom-slot').innerHTML = document.getElementById('people-table-slot').innerHTML;
                    }
                } catch (err) {
                    console.error('Erro ao atualizar painel inferior', err);
                }
            } else {
                document.getElementById('details-subtitle').textContent = 'Caixa selecionada.';
                // Try to find a matching person row to show richer fields (cargo, codigo uo, etc.)
                const selectedName = node.data.subtitle || '';
                const personMatch = allPeople.find(p => normalizeText(p.nome) === normalizeText(selectedName));

                detailCards.push(
                    `<div class="sidebar-card"><div class="sidebar-card-label">UO</div><div class="sidebar-card-value">${info.uo || title || ''}</div></div>`
                );

                // Codigo UO (prefer person value when available)
                detailCards.push(`<div class="sidebar-card"><div class="sidebar-card-label">Codigo UO</div><div class="sidebar-card-value">${(personMatch && personMatch.codigo_uo) || info.codigo_uo || '-'}</div></div>`);

                // Centro de Custo if available on personMatch
                if (personMatch && (personMatch.cc || personMatch.centro_custo)) {
                    detailCards.push(`<div class="sidebar-card"><div class="sidebar-card-label">Centro de Custo</div><div class="sidebar-card-value">${personMatch.cc || personMatch.centro_custo}</div></div>`);
                }

                // Prepare level code and description with multiple possible source keys
                const codigoNivel = (personMatch && (personMatch.codigo_nivel_uo || personMatch.codigo_nivel)) || info.codigo_nivel_uo || info.codigo_nivel || '-';
                const descricaoNivel = (personMatch && (personMatch.nivel_uo || personMatch.nivel || personMatch.descricao_nivel)) || info.nivel_uo || info.nivel || info.descricao_nivel || '-';

                detailCards.push(
                    `<div class="sidebar-card"><div class="sidebar-card-label">Gestor</div><div class="sidebar-card-value">${info.gestor || '-'}</div></div>`,
                    `<div class="sidebar-card"><div class="sidebar-card-label">Matrícula Gestor</div><div class="sidebar-card-value">${(personMatch && personMatch.matricula_gestor) || '-'}</div></div>`,
                    `<div class="sidebar-card"><div class="sidebar-card-label">Nível (Código)</div><div class="sidebar-card-value">${codigoNivel}</div></div>`,
                    `<div class="sidebar-card"><div class="sidebar-card-label">Nível UO</div><div class="sidebar-card-value">${descricaoNivel}</div></div>`,
                    `<div class="sidebar-card"><div class="sidebar-card-label">Pessoa selecionada</div><div class="sidebar-card-value">${personMatch ? personMatch.nome : (node.data.subtitle || '-')}</div></div>`,
                    `<div class="sidebar-card"><div class="sidebar-card-label">Cargo</div><div class="sidebar-card-value">${(personMatch && personMatch.cargo) || '-'}</div></div>`
                );
                document.getElementById('details-content').innerHTML = detailCards.join('');
                const matchingPeople = getPeopleForNode(node);
                document.getElementById('people-count-badge').textContent = String(matchingPeople.length);
                currentPeople = matchingPeople;
                // Do NOT auto-open bottom panel. User can open it via the bottom rail.
                // If bottom panel is already open, refresh its contents immediately
                try {
                    const bottomPanel = document.getElementById('people-bottom-panel');
                    if (bottomPanel && bottomPanel.classList.contains('open')) {
                        renderPeopleTable(currentPeople);
                        document.getElementById('people-bottom-slot').innerHTML = document.getElementById('people-table-slot').innerHTML;
                    }
                } catch (err) {
                    console.error('Erro ao atualizar painel inferior', err);
                }
            }
        }

        function getPeopleForNode(node) {
            const info = node.data.info || {};
            const nodeUoRaw = info.uo || info.codigo_uo || node.data.title || '';
            const nodeUo = normalizeText(nodeUoRaw);
            const gestorRaw = info.gestor || '';
            const gestor = normalizeText(gestorRaw);

            return allPeople
                .filter(person => {
                    const personUo = normalizeText(person.uo || person.codigo_uo || '');
                    const personUoKey = normalizeText(person.uo_key || '');
                    const personGestor = normalizeText(person.gestor || '');
                            // If the node has a UO, only match by exact normalized UO or its key.
                            if (nodeUo) {
                                if (personUo === nodeUo || personUoKey === nodeUo) return true;
                                return false;
                            }
                            // If node has no UO info, fall back to matching by gestor name (rare case).
                            if (gestor && personGestor === gestor) return true;
                            return false;
                })
                .sort((a, b) => (a.nome || '').localeCompare(b.nome || '', 'pt-BR'));
        }

        function renderPeopleTable(people) {
            const slot = document.getElementById('people-table-slot');
            if (!people.length) {
                slot.innerHTML = '<div style="color: #777; font-size: 13px; padding: 10px 0;">Nenhuma pessoa encontrada.</div>';
                return;
            }
            slot.innerHTML = `<table class="people-table"><thead><tr><th>Colaborador</th><th>Matricula</th><th>Gestor</th><th>UO</th><th>Codigo UO</th><th>Nivel</th><th>Cargo</th></tr></thead><tbody>${people.map(person => `<tr><td>${person.nome || ''}</td><td>${person.matricula || ''}</td><td>${person.gestor || ''}</td><td>${person.uo || ''}</td><td>${person.codigo_uo || ''}</td><td>${person.nivel || ''}</td><td>${person.cargo || ''}</td></tr>`).join('')}</tbody></table>`;
        }

        function walkVisibleAndHiddenNodes(node, visitor, parent = null) {
            if (!node) return;
            visitor(node, parent);
            const children = node.children || node._children || [];
            children.forEach(child => walkVisibleAndHiddenNodes(child, visitor, node));
        }

        function focusNode(targetNode) {
            if (!targetNode) return;
            resetTreeToNode(targetNode);
        }

        function findNodeForPerson(person) {
            const personName = normalizeText(person.nome || '');
            const personUo = normalizeText(person.uo || person.codigo_uo || '');
            const personGestor = normalizeText(person.gestor || '');
            let found = null;

            walkVisibleAndHiddenNodes(root, (node) => {
                if (found) return;
                const info = node.data && node.data.info ? node.data.info : {};
                const nodeUo = normalizeText(info.uo || info.codigo_uo || node.data.title || node.data.name || '');
                const nodeGestor = normalizeText(info.gestor || '');
                const nodeName = normalizeText(node.data.subtitle || node.data.title || node.data.name || '');

                if (personName && nodeName && nodeName === personName) {
                    found = node;
                    console.debug('findNodeForPerson: matched by nodeName', {personName, nodeName, nodeId: node.data.id, nodeType: node.data.type});
                    return;
                }
                if (personUo && nodeUo && nodeUo === personUo) {
                    found = node;
                    console.debug('findNodeForPerson: matched by nodeUo', {personUo, nodeUo, nodeId: node.data.id, nodeType: node.data.type});
                    return;
                }
                if (personUo && (node.data.aggregated_people || []).some((p) => normalizeText(p.uo_desc || p.uo_code || '') === personUo)) {
                    found = node;
                    console.debug('findNodeForPerson: matched in aggregated_people', {personUo, nodeId: node.data.id, nodeType: node.data.type, aggCount: (node.data.aggregated_people || []).length});
                    return;
                }
                if (!personUo && personGestor && nodeGestor && nodeGestor === personGestor) {
                    found = node;
                }
            });

            return found;
        }

        function renderSearchResults(matches) {
            searchResults.classList.add('open');
            if (!matches.length) {
                searchResults.innerHTML = '<strong>Nenhum resultado encontrado</strong>';
                return;
            }

            const limitedMatches = matches.slice(0, 10);
            const resultsHTML = `<strong>Encontrados ${matches.length} resultado(s):</strong><ul style="margin: 8px 0 0 0; padding: 0; list-style: none;">${limitedMatches.map((person, index) => `<li><button type="button" class="search-result-btn" data-search-index="${index}"><strong>${person.nome}</strong> - ${person.cargo || 'N/A'}<br/><small>Gestor: ${person.gestor} | UO: ${person.uo}</small></button></li>`).join('')}${matches.length > 10 ? `<li style="margin: 4px 0; font-style: italic; color: #999;">... e mais ${matches.length - 10}</li>` : ''}</ul>`;
            searchResults.innerHTML = resultsHTML;

            searchResults.querySelectorAll('[data-search-index]').forEach((button) => {
                button.addEventListener('click', () => {
                    const index = Number(button.getAttribute('data-search-index'));
                    const person = limitedMatches[index];
                    const targetNode = findNodeForPerson(person);
                    searchResults.innerHTML = '';
                    searchResults.classList.remove('open');
                    console.debug('search click: targetNode', {targetNodeId: targetNode ? targetNode.data.id : null, targetNodeType: targetNode ? targetNode.data.type : null});
                    if (targetNode) {
                        focusNode(targetNode);
                    } else {
                        console.debug('search click: no exact node found, trying by gestor');
                        const fallback = findNodeForPerson({uo: '', gestor: person.gestor});
                        console.debug('search click: fallback', {fallbackId: fallback ? fallback.data.id : null});
                        if (fallback) focusNode(fallback);
                    }
                });
            });
        }

        // listar-pessoa moved to bottom rail; node click will populate currentPeople and open bottom panel

        // limpar-lista removed from sidebar; use bottom rail to close panel

        document.getElementById('people-rail').addEventListener('click', () => {
            const panel = document.getElementById('people-bottom-panel');
            if (!panel) return;
            const opening = !panel.classList.contains('open');
            panel.classList.toggle('open');
            const bottomSlot = document.getElementById('people-bottom-slot');
            if (opening) {
                if (currentPeople && currentPeople.length) {
                    // render into the sidebar slot then copy to bottom slot
                    renderPeopleTable(currentPeople);
                    bottomSlot.innerHTML = document.getElementById('people-table-slot').innerHTML;
                } else {
                    bottomSlot.innerHTML = '<div style="color:#777;padding:12px;">Nenhuma selecao ativa. Clique em uma caixa para ver colaboradores.</div>';
                }
            } else {
                // closing - clear bottom slot
                bottomSlot.innerHTML = '';
            }
        });

        svg.on('click', () => {
            selectedNode = null;
            hidePeopleList();
            document.getElementById('info-drawer').classList.remove('open');
        });

        svg.call(zoomBehavior);

        g.attr('transform', `translate(${width / 2}, ${height / 2})`);

        const verticalStep = Math.max(240, Math.round(window.innerHeight * 0.24));
        const tree = d3.tree().nodeSize([260, verticalStep]).separation((a, b) => (a.parent === b.parent ? 1.8 : 2.2));
        const root = d3.hierarchy(data);
        if (root.children) root.children.forEach(collapse);

        update(root);

        function update(source) {
            tree(root);
            const nodes_data = root.descendants();

            // PASS 1: Align ALL UO (aggregated) nodes to their parent's X position
            nodes_data.forEach((node) => {
                if (!node || !node.data || node.data.type !== 'uo' || !node.parent) return;
                node.x = node.parent.x;
            });

            // PASS 2: Keep each UO centered under its parent and move the remaining blue siblings to one side.
            nodes_data.forEach((parent) => {
                if (!parent || !parent.children || !parent.children.length) return;

                const uoChildren = parent.children.filter((child) => child && child.data && child.data.type === 'uo');
                if (!uoChildren.length) return;

                const blueSiblings = parent.children.filter((child) => child && child.data && child.data.type === 'gestor');
                const gap = 40;

                // Lay out multiple green boxes side-by-side so they do not overlap.
                const uoBoxes = uoChildren.map((uoChild) => nodeBox(uoChild));
                const totalUoWidth = uoBoxes.reduce((sum, box) => sum + box.width, 0) + (uoChildren.length - 1) * gap;
                let currentX = parent.x - (totalUoWidth / 2);

                uoChildren.forEach((uoChild, index) => {
                    const uoBox = uoBoxes[index];
                    uoChild.x = currentX + (uoBox.width / 2);
                    currentX += uoBox.width + gap;
                });

                if (!blueSiblings.length) return;

                // Push the remaining blue boxes to one side of the whole green block.
                const direction = parent.x <= 0 ? 1 : -1;
                const edgeOffset = (totalUoWidth / 2) + gap;
                let currentBlueX = parent.x + direction * edgeOffset;

                blueSiblings.forEach((child) => {
                    const childBox = nodeBox(child);
                    child.x = currentBlueX + direction * (childBox.width / 2);
                    currentBlueX += direction * (childBox.width + gap);
                });
            });

            const links_data = root.links()
                .filter(d => !(d.target && d.target.data && d.target.data.info && d.target.data.info.synthetic_parent));
            const links = g.selectAll('.link').data(links_data, d => `${d.source.data.id}__${d.target.data.id}`);
            const orthogonalLink = (d) => {
                const midY = (d.source.y + d.target.y) / 2;
                return `M${d.source.x},${d.source.y}V${midY}H${d.target.x}V${d.target.y}`;
            };
            links.enter().insert('path', ':first-child')
                .attr('class', 'link')
                .merge(links)
                .transition().duration(300)
                .attr('d', orthogonalLink);
            links.exit().remove();
            const nodes = g.selectAll('.node').data(nodes_data.filter(d => !d.data.layout_only), d => d.data.id);
            const nodeEnter = nodes.enter()
                .append('g')
                .attr('class', d => `node ${d.data.type}`)
                .attr('transform', d => `translate(${d.x},${d.y})`)
                .on('click', function(event, d) {
                    event.stopPropagation();
                    if (selectedNode && selectedNode.data.id !== d.data.id) hidePeopleList();
                    try {
                        // If clicking the already-selected node, toggle only this node's children
                        if (selectedNode && selectedNode.data.id === d.data.id) {
                            if (d.children) {
                                d._children = d.children;
                                d.children = null;
                            } else if (d._children) {
                                d.children = d._children;
                                d._children = null;
                            }
                            update(root);
                            return;  // Don't recenter on toggle
                        }

                        // Otherwise, collapse all and open only the path to the clicked node
                        collapseAll(root);
                        expandAncestors(d);
                        if (d._children) {
                            d.children = d._children;
                            d._children = null;
                        }
                        selectedNode = d;
                        renderDetailsForNode(d);
                        update(root);
                        centerNode(d);  // Only recenter when selecting a new node
                    } catch (err) {
                        console.error('Error rendering details', err);
                    }
                })
                .on('mouseover', function(event, d) {
                    const tooltip = document.getElementById('tooltip');
                    if (d.data.info) {
                        tooltip.innerHTML = `<strong>${d.data.title || d.data.name}</strong><br/>Gestor: ${d.data.info.gestor || ''}<br/>UO: ${d.data.info.uo || ''}<br/>Nivel: ${d.data.info.nivel || ''}`;
                        tooltip.style.display = 'block';
                        tooltip.style.left = event.pageX + 10 + 'px';
                        tooltip.style.top = event.pageY + 10 + 'px';
                    }
                })
                .on('mouseout', () => {
                    document.getElementById('tooltip').style.display = 'none';
                });

            nodeEnter.append('rect')
                .attr('x', d => -nodeBox(d).width / 2)
                .attr('y', d => -nodeBox(d).height / 2)
                .attr('width', d => nodeBox(d).width)
                .attr('height', d => nodeBox(d).height);

            nodeEnter.append('text')
                .attr('class', 'title')
                .attr('y', d => d.data.subtitle ? -4 : 4)
                .text(d => d.data.title || d.data.name);

            nodeEnter.append('text')
                .attr('class', 'subtitle')
                .attr('y', 12)
                .text(d => d.data.subtitle || '');

            nodes.merge(nodeEnter)
                .transition()
                .duration(300)
                .attr('transform', d => `translate(${d.x},${d.y})`)
                .select('rect')
                .attr('x', d => -nodeBox(d).width / 2)
                .attr('y', d => -nodeBox(d).height / 2)
                .attr('width', d => nodeBox(d).width)
                .attr('height', d => nodeBox(d).height);

            nodes.exit().remove();
        }

        renderEmptyDetails('Nenhum item selecionado.');

        const searchInput = document.getElementById('search-input');
        const searchResults = document.getElementById('search-results');

        searchInput.addEventListener('input', (e) => {
            const query = normalizeText(e.target.value);
            if (!query) {
                searchResults.innerHTML = '';
                searchResults.classList.remove('open');
                return;
            }
            const matches = allPeople.filter(person => {
                return normalizeText(person.nome).includes(query) ||
                       normalizeText(person.gestor).includes(query) ||
                       normalizeText(person.uo).includes(query) ||
                       normalizeText(person.matricula).includes(query) ||
                       normalizeText(person.cargo).includes(query);
            });
            renderSearchResults(matches);
        });
    </script>
</body>
</html>"""

    html_content = html_template.replace('__TREE_JSON__', tree_json).replace('__PEOPLE_JSON__', people_json)

    st.components.v1.html(html_content, height=1200)

    st.download_button(
        label="Baixar Organograma como HTML",
        data=html_content,
        file_name="organograma_interativo.html",
        mime="text/html",
        key="download_html_organograma"
    )

except Exception as e:
    st.error(f"Erro ao processar o arquivo: {e}")
    st.write(traceback.format_exc())
