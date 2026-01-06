"""
Página: Banco de Horas
Descrição: Visualização e geração de relatório de banco de horas por centro de custo
Recebe: Arquivo XLSX com banco de horas
"""

import streamlit as st
import pandas as pd
import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Banco de Horas", layout="wide")

st.title("🏦 Banco de Horas")
st.write("Geração de relatório de banco de horas por centro de custo")

st.divider()

# Upload do arquivo
st.subheader("📥 Selecione o arquivo de Banco de Horas")
file_banco_horas = st.file_uploader(
    "Arquivo XLSX com banco de horas",
    type=["xlsx"],
    key="banco_horas"
)

if file_banco_horas:
    try:
        # Carrega o arquivo
        df = pd.read_excel(file_banco_horas)
        
        st.success("✅ Arquivo carregado com sucesso!")
        
        st.divider()
        
        # Preview dos dados
        with st.expander("👀 Visualizar dados carregados"):
            st.dataframe(df, use_container_width=True)
            st.write(f"**Colunas disponíveis:** {list(df.columns)}")
        
        st.divider()
        
        # Configuração de colunas
        st.subheader("⚙️ Configurar colunas")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            col_centro_custo = st.selectbox(
                "Coluna de Centro de Custo",
                df.columns,
                key="centro_custo"
            )
        
        with col2:
            col_positivo = st.selectbox(
                "Coluna de Horas POSITIVAS",
                df.columns,
                key="positivo",
                index=1 if len(df.columns) > 1 else 0
            )
        
        with col3:
            col_negativo = st.selectbox(
                "Coluna de Horas NEGATIVAS",
                df.columns,
                key="negativo",
                index=2 if len(df.columns) > 2 else 0
            )
        
        st.divider()
        
        # Botão para gerar relatório
        if st.button("📊 Gerar Relatório", use_container_width=True):
            try:
                # Processa os dados
                df_processado = df[[col_centro_custo, col_positivo, col_negativo]].copy()
                df_processado.columns = ['Centro de Custo', 'POSITIVO', 'NEGATIVO']
                
                # Remove linhas onde centro de custo está vazio
                df_processado = df_processado[df_processado['Centro de Custo'].notna()]
                df_processado = df_processado[df_processado['Centro de Custo'] != '']
                df_processado = df_processado[df_processado['Centro de Custo'].astype(str).str.strip() != '']
                
                # Função para converter tempo (HH:MM:SS) para horas decimais
                def tempo_para_horas(valor):
                    if pd.isna(valor) or valor == '' or valor == 0:
                        return 0.0
                    
                    try:
                        if isinstance(valor, str):
                            partes = valor.split(':')
                            if len(partes) == 3:
                                horas = int(partes[0])
                                minutos = int(partes[1])
                                segundos = int(partes[2])
                                return horas + minutos/60 + segundos/3600
                            else:
                                return 0.0
                        else:
                            # Se for datetime.time
                            if hasattr(valor, 'hour'):
                                return valor.hour + valor.minute/60 + valor.second/3600
                            else:
                                return float(valor)
                    except:
                        return 0.0
                
                # Função para converter horas decimais de volta para HH:MM:SS
                def horas_para_tempo(horas):
                    if pd.isna(horas) or horas == 0:
                        return "0:00:00"
                    
                    total_segundos = int(horas * 3600)
                    h = total_segundos // 3600
                    m = (total_segundos % 3600) // 60
                    s = total_segundos % 60
                    return f"{h}:{m:02d}:{s:02d}"
                
                # Converte para número
                df_processado['POSITIVO_num'] = df_processado['POSITIVO'].apply(tempo_para_horas)
                df_processado['NEGATIVO_num'] = df_processado['NEGATIVO'].apply(tempo_para_horas)
                
                # Agrupa por centro de custo e soma
                df_resumo = df_processado.groupby('Centro de Custo')[['POSITIVO_num', 'NEGATIVO_num']].sum().reset_index()
                
                # Converte de volta para tempo
                df_resumo['POSITIVO'] = df_resumo['POSITIVO_num'].apply(horas_para_tempo)
                df_resumo['NEGATIVO'] = df_resumo['NEGATIVO_num'].apply(horas_para_tempo)
                
                # Remove colunas numéricas
                df_resumo = df_resumo[['Centro de Custo', 'POSITIVO', 'NEGATIVO']]
                
                st.success("✅ Relatório gerado com sucesso!")
                
                st.divider()
                
                # Mostra preview
                st.subheader("📋 Resumo por Centro de Custo")
                st.dataframe(df_resumo, use_container_width=True)
                
                # Cria arquivo Excel para download
                wb = Workbook()
                ws = wb.active
                ws.title = "Banco de Horas Total"
                
                # Define estilos
                header_fill = PatternFill(start_color="FF0D4F45", end_color="FF0D4F45", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFFFF", size=11)
                header_alignment = Alignment(horizontal="center", vertical="center")
                
                border_style = Side(style="medium", color="000000")
                border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
                
                # Adiciona título
                ws['A1'] = "Banco de Horas Total"
                ws['A1'].font = Font(bold=True, size=12)
                ws.merge_cells('A1:C1')
                ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
                
                # Adiciona headers
                headers = ['Centro de Custo', 'POSITIVO', 'NEGATIVO']
                for col_idx, header in enumerate(headers, 1):
                    cell = ws.cell(row=3, column=col_idx, value=header)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_alignment
                    cell.border = border
                
                # Adiciona dados
                for row_idx, (_, row) in enumerate(df_resumo.iterrows(), 4):
                    ws.cell(row=row_idx, column=1, value=row['Centro de Custo'])
                    ws.cell(row=row_idx, column=2, value=row['POSITIVO'])
                    ws.cell(row=row_idx, column=3, value=row['NEGATIVO'])
                    
                    for col in [1, 2, 3]:
                        cell = ws.cell(row=row_idx, column=col)
                        cell.border = border
                        if col > 1:
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                
                # Adiciona total
                total_row = len(df_resumo) + 5
                ws.cell(row=total_row, column=1, value="Total Geral")
                ws.cell(row=total_row, column=1).font = Font(bold=True)
                ws.cell(row=total_row, column=1).fill = PatternFill(start_color="FFF0F0F0", end_color="FFF0F0F0", fill_type="solid")
                
                # Soma as colunas
                total_positivo = df_resumo['POSITIVO_num'].sum()
                total_negativo = df_resumo['NEGATIVO_num'].sum()
                
                ws.cell(row=total_row, column=2, value=horas_para_tempo(total_positivo))
                ws.cell(row=total_row, column=3, value=horas_para_tempo(total_negativo))
                
                for col in [1, 2, 3]:
                    cell = ws.cell(row=total_row, column=col)
                    cell.border = border
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="FFF0F0F0", end_color="FFF0F0F0", fill_type="solid")
                    if col > 1:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                
                # Ajusta largura das colunas
                ws.column_dimensions['A'].width = 40
                ws.column_dimensions['B'].width = 20
                ws.column_dimensions['C'].width = 20
                
                # Salva em memória
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                # Botão de download
                st.download_button(
                    label="⬇️ Baixar Relatório (XLSX)",
                    data=output.getvalue(),
                    file_name=f"Banco_de_Horas_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            except Exception as e:
                st.error(f"❌ Erro ao gerar relatório: {str(e)}")
                st.write(f"Detalhes: {e}")
    
    except Exception as e:
        st.error(f"❌ Erro ao carregar arquivo: {str(e)}")

else:
    st.info("""
    ### 📌 Como usar:
    
    1. **Upload**: Carregue o arquivo XLSX com banco de horas
    2. **Configurar**: Selecione as colunas corretas para:
       - Centro de Custo
       - Horas POSITIVAS
       - Horas NEGATIVAS
    3. **Gerar**: Clique em "Gerar Relatório"
    4. **Download**: Baixe o arquivo XLSX com o resumo consolidado
    """)

