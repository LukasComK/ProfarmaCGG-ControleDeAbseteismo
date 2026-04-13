import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Relatório Integrado", page_icon="📄", layout="wide")

st.title("📄 Relatório Integrado de Absenteísmo")
st.markdown("""
Esta página cruza dados de 4 planilhas diferentes para gerar um relatório automático 
agrupando os colaboradores com **Faltas Injustificadas (FI)** e **Faltas por Atestado (FA)**.
""")

st.divider()

col1, col2 = st.columns(2)
with col1:
    f_abs = st.file_uploader("1️⃣ Planilha de Absenteísmo (Lê a aba 'Dados')", type=["xlsx"])
    f_med = st.file_uploader("2️⃣ Planilha de Medida Disciplinar", type=["xlsx"])
with col2:
    f_dem = st.file_uploader("3️⃣ Planilha de Demissões", type=["xlsx"])
    f_ent = st.file_uploader("4️⃣ Planilha de Entrevista (Lê aba '2026')", type=["xlsx"])

def limpar_nome(nome):
    if pd.isna(nome):
        return ""
    return str(nome).strip().upper()

if f_abs and f_med and f_dem and f_ent:
    if st.button("🚀 Gerar Relatório Integrado", type="primary", use_container_width=True):
        with st.spinner("Analisando e cruzando as planilhas. Isso pode levar alguns segundos..."):
            try:
                # ---------------------------------------------------------
                # 1. PROCESSAR PLANILHA DE ABSENTEÍSMO
                # ---------------------------------------------------------
                # Tenta ler a aba Dados. Se não achar, lê a primeira
                try:
                    df_abs = pd.read_excel(f_abs, sheet_name="Dados", header=None)
                except:
                    f_abs.seek(0)
                    df_abs = pd.read_excel(f_abs, header=None)
                
                absencias = {} # Formato: { "NOME": {"FI": [], "FA": []} }
                
                # Mapear os cabeçalhos de data (Colunas J até AN -> Índices 9 a 39)
                datas_colunas = {}
                col_fim = min(40, len(df_abs.columns))
                
                # Tenta achar a linha de cabeçalho (que tem as datas) olhando as primeiras linhas
                linha_cabecalho = 0
                for r in range(min(5, len(df_abs))):
                    if pd.notna(df_abs.iloc[r, 9]): # Se tem algo na coluna J
                        linha_cabecalho = r
                        break
                        
                for c in range(9, col_fim):
                    val = df_abs.iloc[linha_cabecalho, c]
                    # Formatar a data para ficar bonitinha caso venha como datetime
                    if isinstance(val, pd.Timestamp):
                        datas_colunas[c] = val.strftime('%d/%m')
                    else:
                        datas_colunas[c] = str(val)[:10] if pd.notna(val) else f"Col {c+1}"

                # Percorrer os colaboradores
                for r in range(linha_cabecalho + 1, len(df_abs)):
                    nome = limpar_nome(df_abs.iloc[r, 0]) # Coluna A (Índice 0)
                    if not nome: continue
                    
                    for c in range(9, col_fim):
                        status = limpar_nome(df_abs.iloc[r, c])
                        if status in ['FI', 'FA']:
                            if nome not in absencias:
                                absencias[nome] = {'FI': [], 'FA': []}
                            absencias[nome][status].append(datas_colunas[c])

                # ---------------------------------------------------------
                # 2. PROCESSAR MEDIDAS DISCIPLINARES
                # ---------------------------------------------------------
                f_med.seek(0)
                df_med = pd.read_excel(f_med, header=None)
                medidas_dict = {}
                
                col_max_med = len(df_med.columns)
                for r in range(len(df_med)):
                    # Col B (1), Col D (3), Col E (4)
                    if col_max_med > 1:
                        nome = limpar_nome(df_med.iloc[r, 1])
                    else: continue
                    
                    tipo = limpar_nome(df_med.iloc[r, 3]) if col_max_med > 3 else ""
                    detalhe_data = str(df_med.iloc[r, 4]).strip() if col_max_med > 4 and pd.notna(df_med.iloc[r, 4]) else ""
                    
                    if nome and "FALTA INJUSTIFICADA" in tipo:
                        medidas_dict[nome] = detalhe_data

                # ---------------------------------------------------------
                # 3. PROCESSAR DEMISSÕES
                # ---------------------------------------------------------
                f_dem.seek(0)
                df_dem = pd.read_excel(f_dem, header=None)
                demissoes_dict = {}
                
                col_max_dem = len(df_dem.columns)
                for r in range(len(df_dem)):
                    # Procura na Col B (1), D (3), F (5)
                    if col_max_dem > 1:
                        nome = limpar_nome(df_dem.iloc[r, 1])
                    else: continue
                    
                    if nome:
                        d_data = str(df_dem.iloc[r, 3])[:10] if col_max_dem > 3 and pd.notna(df_dem.iloc[r, 3]) else "Data Indisponível"
                        if isinstance(df_dem.iloc[r, 3], pd.Timestamp):
                            d_data = df_dem.iloc[r, 3].strftime('%d/%m/%Y')
                            
                        d_tipo = str(df_dem.iloc[r, 5]).strip() if col_max_dem > 5 and pd.notna(df_dem.iloc[r, 5]) else "Tipo Indisponível"
                        demissoes_dict[nome] = {'data': d_data, 'tipo': d_tipo}
                
                # ---------------------------------------------------------
                # 4. PROCESSAR ENTREVISTAS DE ABSENTEÍSMO
                # ---------------------------------------------------------
                f_ent.seek(0)
                try:
                    df_ent = pd.read_excel(f_ent, sheet_name="2026", header=None)
                except:
                    # Fallback para a primeira aba se 2026 não existir
                    f_ent.seek(0)
                    df_ent = pd.read_excel(f_ent, header=None)
                    
                entrevistas_set = set()
                col_max_ent = len(df_ent.columns)
                for r in range(len(df_ent)):
                    if col_max_ent > 1:
                        nome = limpar_nome(df_ent.iloc[r, 1]) # Col B (1)
                        if nome:
                            entrevistas_set.add(nome)

                # =========================================================
                # CONSTRUÇÃO DO RELATÓRIO
                # =========================================================
                st.success("✅ Cruzamento de dados concluído com sucesso!")
                
                relatorio = []
                relatorio.append("# 📄 RELATÓRIO INTEGRADO DE ABSENTEÍSMO\n")
                
                # ----------- BLOCO 1: FALTAS INJUSTIFICADAS (FI) -----------
                relatorio.append("## 🔴 COLABORADORES COM FALTA INJUSTIFICADA (FI)")
                relatorio.append("---\n")
                
                fi_encontrada = False
                for nome, rec in sorted(absencias.items()):
                    if rec['FI']:
                        fi_encontrada = True
                        relatorio.append(f"### 👤 {nome}")
                        relatorio.append(f"- **Datas das Faltas (FI):** {', '.join(rec['FI'])}")
                        
                        # Cruza com Medida Disciplinar
                        med = medidas_dict.get(nome)
                        if med:
                            relatorio.append(f"- **Medida Disciplinar:** ✅ Solicitada _{'('+med+')' if med else ''}_")
                        else:
                            relatorio.append("- **Medida Disciplinar:** ⚠️ NÃO teve medida solicitada")
                            
                        # Cruza com Demissões
                        dem = demissoes_dict.get(nome)
                        if dem:
                            relatorio.append(f"- **Desligamento:** 🛑 Projeção para {dem['data']} - {dem['tipo']}")
                        else:
                            relatorio.append("- **Desligamento:** Sem projeção de desligamento")
                        
                        relatorio.append("\n") # Espaço
                
                if not fi_encontrada:
                    relatorio.append("\n*Nenhum colaborador com Falta Injustificada (FI) encontrado!*\n\n")
                
                # ----------- BLOCO 2: FALTAS POR ATESTADO (FA) -----------
                relatorio.append("\n## 🟡 COLABORADORES COM FALTAS POR ATESTADO (FA)")
                relatorio.append("---\n")
                
                fa_encontrada = False
                for nome, rec in sorted(absencias.items()):
                    if rec['FA']:
                        fa_encontrada = True
                        relatorio.append(f"### 👤 {nome}")
                        relatorio.append(f"- **Datas dos Atestados (FA):** {', '.join(rec['FA'])}")
                        
                        # Cruza com Entrevista
                        if nome in entrevistas_set:
                            relatorio.append("- **Entrevista de Absenteísmo:** ✅ Possui entrevista registrada nas planilhas")
                        else:
                            relatorio.append("- **Entrevista de Absenteísmo:** ⚠️ NÃO possui entrevista registrada")
                            
                        relatorio.append("\n") # Espaço
                        
                if not fa_encontrada:
                    relatorio.append("\n*Nenhum colaborador com Falta por Atestado (FA) encontrado!*\n\n")

                # =========================================================
                # EXIBIÇÃO E EXPORTAÇÃO
                # =========================================================
                texto_final = "\n".join(relatorio)
                
                # Exibe em tela na interface do Streamlit usando um container destacado
                with st.container(border=True):
                    st.markdown(texto_final)
                
                st.divider()
                st.download_button(
                    label="📥 Exportar Relatório como Texto (TXT)",
                    data=texto_final,
                    file_name="Relatorio_Cruzado_Absenteismo.txt",
                    mime="text/plain",
                    use_container_width=True
                )
                        
            except Exception as e:
                st.error(f"❌ Ocorreu um erro durante o processamento das planilhas: {e}")
                st.exception(e)
else:
    st.info("⚠️ Aguardando o envio das 4 planilhas para habilitar o relatório.")
