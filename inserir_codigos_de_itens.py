import pandas as pd

def gerar_planilha_com_codigos(
	caminho_planilha_modelo: str,
	caminho_csv_codigos: str,
	caminho_saida: str,
	nome_coluna_codigo: str = "item(table) + it-codigo(field)",
) -> None:
	"""Gera uma nova planilha Excel preenchendo apenas a coluna de códigos.

	- Mantém a primeira linha igual à planilha modelo.
	- A partir da segunda linha, preenche só a coluna de código com os valores do CSV.
	"""

	# 1) Ler planilha padrão (modelo) e CSV de códigos
	df_planilha = pd.read_excel(caminho_planilha_modelo)
	print("Planilha Padrão carregada com sucesso.")

	df_codigos = pd.read_csv(caminho_csv_codigos, header=None, names=["CODIGO"])
	print("CSV de Códigos carregado com sucesso.")

	# 2) Usar a primeira linha da planilha como modelo
	linha_modelo = df_planilha.iloc[0]

	# 3) Criar nova planilha: título na primeira linha, dados a partir da segunda
	quantidade_codigos = len(df_codigos)
	df_novo = pd.concat([linha_modelo.to_frame().T] * (quantidade_codigos + 1), ignore_index=True)

	# 4) Deixar todas as outras colunas vazias a partir da segunda linha
	coluna_codigo = nome_coluna_codigo
	colunas_outros = [c for c in df_novo.columns if c != coluna_codigo]
	df_novo.loc[1:, colunas_outros] = ""

	# 5) Preencher a coluna de código a partir da segunda linha
	df_novo.loc[1:, coluna_codigo] = df_codigos["CODIGO"].values

	# 6) Salvar planilha atualizada
	df_novo.to_excel(caminho_saida, index=False)
	print("Planilha atualizada salva com sucesso.")

