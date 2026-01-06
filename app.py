"""
Wrapper para executar Controle_de_Absenteismo.py
Mantém compatibilidade com Streamlit Cloud que procura por app.py
"""

import subprocess
import sys

if __name__ == "__main__":
    subprocess.run([sys.executable, "-m", "streamlit", "run", "Controle_de_Absenteismo.py"])











