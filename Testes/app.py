import pandas as pd
from inserir_codigos_de_itens import gerar_planilha_com_codigos
from inserir_internal_comment import inserir_internal_coments
from inserir_traducoes import Traducoes


# 1) Gera a planilha com códigos
gerar_planilha_com_codigos(
	caminho_planilha_modelo=r"planilhas/planilhaPadrao.xlsx",
	caminho_csv_codigos=r"planilhas/dados_teste.csv",
	caminho_saida=r"planilhas/planilha_atualizada.xlsx",
)
print("Planilha atualizada gerada com sucesso.")

# 2) Insere dados (narrativa SAP123, etc.)
inserir_internal_coments(
	caminho_planilha_atualizada=r"planilhas/planilha_atualizada.xlsx",
	caminho_base_totvs=r"planilhas/base_dados_TOTVS.xlsx",
)

# 3) Usa a classe Traducoes, referenciando a coluna "SAP123" como base
trad = Traducoes(
	caminho_narrativas=r"planilhas/planilha_atualizada.xlsx",   # deve ter coluna 'SAP123'
	caminho_traducoes=r"dados/dicionario.xlsx",       # deve ter coluna base em português + colunas de idiomas
	coluna_narrativa="SAP123",
	coluna_portugues="MAKTX(PT)",  # nome real da coluna em português na planilha de dicionário
)
trad.processar_narrativas()
trad.salvar_tabela(r"planilhas/planilha_atualizada.xlsx")
print("Traduções inseridas com sucesso.")

