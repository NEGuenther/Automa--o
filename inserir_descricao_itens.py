import pandas as pd


def iserir_dados():
	df_planilha_atualizada = pd.read_excel(r"planilhas/planilha_atualizada.xlsx")
	print("Planilha Atualizada carregada com sucesso.")

	df_base_dados_TOTVS = pd.read_excel(r"planilhas/base_dados_TOTVS.xlsx")
	print("Base de Dados TOTVS carregada com sucesso.")
	
    
