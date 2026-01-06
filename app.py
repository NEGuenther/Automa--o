import pandas as pd
from inserir_codigos_de_itens import gerar_planilha_com_codigos
from inserir_internal_comment import inserir_dados


# 1) Gera a planilha_atualizada.xlsx com os códigos dos itens
gerar_planilha_com_codigos(
	caminho_planilha_modelo=r"planilhas/planilhaPadrao.xlsx",
	caminho_csv_codigos=r"planilhas/dados_teste.csv",
	caminho_saida=r"planilhas/planilha_atualizada.xlsx",
)
print("Planilha atualizada gerada com sucesso.")
# 2) Após gerar a planilha com os códigos, insere a coluna de narrativa (SAP123)
inserir_dados()

