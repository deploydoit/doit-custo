"""Converte a base DOit de Excel para Parquet.

O app lê a base em Parquet (bem mais rápido que Excel), o que acelera o
"acordar" do app publicado no Streamlit Cloud. Rode este script sempre que
atualizar o arquivo Excel:

    python3 gerar_parquet.py

Ele lê 'ListagemdeProdutos DOit.xlsx' e gera 'ListagemdeProdutos DOit.parquet'
na mesma pasta.
"""

import os
import sys
import pandas as pd

CAMINHO_XLSX = "ListagemdeProdutos DOit.xlsx"
CAMINHO_PARQUET = "ListagemdeProdutos DOit.parquet"


def main():
    if not os.path.exists(CAMINHO_XLSX):
        print(f"ERRO: não encontrei '{CAMINHO_XLSX}' nesta pasta.")
        sys.exit(1)

    print(f"Lendo {CAMINHO_XLSX} ...")
    df = pd.read_excel(CAMINHO_XLSX)
    # Mantemos os mesmos tipos que pd.read_excel produz (o app depende disso:
    # 'Id do Fabricante' numérico, etc.). Só normalizamos nomes de coluna para str.
    df.columns = [str(c) for c in df.columns]

    print(f"Gravando {CAMINHO_PARQUET} ({len(df):,} linhas) ...")
    df.to_parquet(CAMINHO_PARQUET, engine="pyarrow", index=False)

    tam_xlsx = os.path.getsize(CAMINHO_XLSX) / 1e6
    tam_parq = os.path.getsize(CAMINHO_PARQUET) / 1e6
    print(f"OK! Excel: {tam_xlsx:.1f} MB  ->  Parquet: {tam_parq:.1f} MB")


if __name__ == "__main__":
    main()
