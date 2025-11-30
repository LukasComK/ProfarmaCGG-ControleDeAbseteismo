import pandas as pd

# Tenta ler com diferentes codificações
encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']

for enc in encodings:
    try:
        df = pd.read_csv('exemplo_colaboradores.csv', encoding=enc)
        print(f"Sucesso com encoding: {enc}")
        # Salva em UTF-8
        df.to_csv('exemplo_colaboradores.csv', encoding='utf-8', index=False)
        print("Arquivo convertido para UTF-8!")
        break
    except Exception as e:
        print(f"Falhou com {enc}: {str(e)[:50]}")
