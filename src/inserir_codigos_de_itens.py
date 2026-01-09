import pandas as pd

def gerar_planilha_com_codigos(
	caminho_planilha_modelo: str,
	caminho_csv_codigos: str,
	caminho_saida: str,
	nome_coluna_codigo: str = "item(table) + it-codigo(field)",
) -> None:
	"""Gera uma nova planilha Excel preenchendo apenas a coluna de códigos.

	Layout esperado do modelo:
	- Excel linha 1: cabeçalho técnico (vira o header do arquivo de saída)
	- Excel linha 2: linha descritiva (vira a primeira linha de dados / df index 0)

	Saída:
	- Mantém a linha descritiva.
	- A partir da Excel linha 3 (df index 1), preenche a primeira coluna com os códigos do CSV.
	"""
	# 1) Ler planilha padrão (modelo) e CSV de códigos
	# Modelo esperado:
	# - linha 1: cabeçalhos técnicos (será o header do Excel)
	# - linha 2: descrições/linha descritiva (fica como primeira linha de dados)
	df_planilha = pd.read_excel(caminho_planilha_modelo, header=None)
	df_codigos = pd.read_csv(caminho_csv_codigos, header=None, names=["CODIGO"])

	if len(df_planilha) < 2:
		raise ValueError("A planilha modelo precisa ter pelo menos 2 linhas (header + linha descritiva).")

	headers = df_planilha.iloc[0].tolist()
	linha_descritiva = df_planilha.iloc[1].copy()

	# 2) Criar as linhas de dados (a partir da terceira linha da planilha Excel:
	#    1ª linha = header, 2ª linha = descritiva)
	quantidade_codigos = len(df_codigos)
	df_dados = pd.concat([linha_descritiva.to_frame().T] * quantidade_codigos, ignore_index=True)

	# 3) Limpar todas as colunas exceto a de código, e preencher os códigos
	colunas_outros = [c for c in df_dados.columns if c != 0]  # coluna 0 é a de código no modelo
	df_dados.loc[:, colunas_outros] = ""
	df_dados.iloc[:, 0] = df_codigos["CODIGO"].values

	# 4) Montar o DataFrame final: 1ª linha de dados = descritiva + linhas com códigos
	df_final = pd.concat([linha_descritiva.to_frame().T, df_dados], ignore_index=True)
	# Usar a primeira linha do modelo como cabeçalho
	df_final.columns = headers

	# 5) Salvar a planilha sem duplicar o header como linha de dados
	df_final.to_excel(caminho_saida, index=False, header=True)

