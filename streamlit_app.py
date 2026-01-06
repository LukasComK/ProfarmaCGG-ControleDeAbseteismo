# Alias para Controle_de_Absenteismo.py
# O Streamlit Cloud reconhece automaticamente streamlit_app.py e app.py
# Este arquivo garante que o Streamlit use o arquivo principal correto

import sys
import os

# Adiciona o diretório atual ao path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Importa e executa o módulo principal
import Controle_de_Absenteismo
