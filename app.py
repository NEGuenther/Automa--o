import pandas as pd
from inserir_codigos_de_itens import gerar_planilha_com_codigos


gerar_planilha_com_codigos(
        caminho_planilha_modelo=r"planilhas/planilhaPadrao.xlsx",
        caminho_csv_codigos=r"planilhas/dados_teste.csv",
        caminho_saida=r"planilhas/planilha_atualizada.xlsx",
)

