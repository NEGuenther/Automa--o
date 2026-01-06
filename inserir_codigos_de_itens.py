import pandas as pd

def gerar_planilha_com_codigos(
	caminho_planilha_modelo: str,
	caminho_csv_codigos: str,
	caminho_saida: str,
	nome_coluna_codigo: str = "item(table) + it-codigo(field)",
) -> None:
	"""Gera uma nova planilha Excel preenchendo apenas a coluna de códigos.

	- Mantém as DUAS primeiras linhas iguais à planilha modelo
	  (ex.: linha 1 = "SAP123", linha 2 = "Internal comment (narrative)").
	- A partir da TERCEIRA linha, preenche só a coluna de código com os valores do CSV.
	"""

	# 1) Ler planilha padrão (modelo) e CSV de códigos
	df_planilha = pd.read_excel(caminho_planilha_modelo, header=None)
	df_codigos = pd.read_csv(caminho_csv_codigos, header=None, names=["CODIGO"])

	# 2) Separar as duas primeiras linhas (títulos) e a linha-modelo de dados
	#    df_planilha tem exatamente 2 linhas no seu modelo atual
	linhas_titulo = df_planilha.iloc[:2].copy()
	linha_modelo_dados = df_planilha.iloc[1].copy()  # usa a segunda linha como modelo de dados

	# 3) Criar as linhas de dados (a partir da terceira linha da nova planilha)
	quantidade_codigos = len(df_codigos)
	df_dados = pd.concat(
		[linha_modelo_dados.to_frame().T] * quantidade_codigos,
		ignore_index=True,
	)

	# 4) Montar o novo DataFrame: 2 linhas de título + linhas de dados
	df_novo = pd.concat([linhas_titulo, df_dados], ignore_index=True)

	# 5) Deixar todas as outras colunas vazias a partir da terceira linha
	coluna_codigo = nome_coluna_codigo
	colunas_outros = [c for c in df_novo.columns if c != 0]  # coluna 0 é a de código no modelo
	df_novo.iloc[2:, colunas_outros] = ""

	# 6) Preencher a coluna de código a partir da terceira linha
	df_novo.iloc[2:, 0] = df_codigos["CODIGO"].values

	# 7) Definir os nomes das colunas a partir da primeira linha (linha de cabeçalho técnica)
	df_novo.columns = df_novo.iloc[0]
	df_novo = df_novo[1:]  # remove a linha de nomes agora transformada em cabeçalho

	# 8) Salvar planilha atualizada
	df_novo.to_excel(caminho_saida, index=False, header=True, index_label=False)

