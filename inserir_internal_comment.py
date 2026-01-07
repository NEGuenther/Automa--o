import pandas as pd


def inserir_internal_coments(
	caminho_planilha_atualizada: str,
	caminho_base_totvs: str,
):
	"""Compara códigos e insere a narrativa na planilha atualizada.

	- Lê a planilha atualizada gerada no passo anterior (caminho informado).
	- Lê a base TOTVS (caminho informado) com códigos e coluna de narrativa.
	- Usa a primeira coluna da planilha atualizada como código e cruza com a coluna "item" (ou similar) da base TOTVS.
	- Preenche a coluna "SAP123" da planilha atualizada com a narrativa correspondente.
	"""

	# Ler arquivos de entrada
	df_planilha_atualizada = pd.read_excel(caminho_planilha_atualizada)
	print("Planilha Atualizada carregada com sucesso.")

	# A planilha base_dados_TOTVS.xlsx possui o cabeçalho real na 5ª linha (índice 4)
	df_base_dados_TOTVS = pd.read_excel(caminho_base_totvs, header=4)
	print("Base de Dados TOTVS carregada com sucesso.")

	# Identificar colunas de código
	# Na planilha atualizada, usamos a primeira coluna (códigos dos itens)
	col_codigo_atualizada = df_planilha_atualizada.columns[0]

	# Na base TOTVS, procuramos uma coluna cujo nome (ignorando maiúsculas/minúsculas)
	# seja exatamente "item"
	col_codigo_base = None
	for c in df_base_dados_TOTVS.columns:
		if str(c).strip().lower() == "item":
			col_codigo_base = c
			break

	if col_codigo_base is None:
		# fallback: mantém a lógica anterior caso o nome mude
		possiveis_codigos_base = [
			c
			for c in df_base_dados_TOTVS.columns
			if "codigo" in str(c).lower() or "código" in str(c).lower()
		]
		col_codigo_base = (
			possiveis_codigos_base[0]
			if possiveis_codigos_base
			else df_base_dados_TOTVS.columns[0]
		)

	# Identificar coluna de narrativa
	possiveis_narrativas = [
		c for c in df_base_dados_TOTVS.columns if "narrativa" in c.lower()
	]
	if not possiveis_narrativas:
		raise ValueError(
			"Não foi encontrada nenhuma coluna de 'narrativa' na baseDadosTOTVS.xlsx. "
			"Verifique o nome das colunas."
		)
	col_narrativa = possiveis_narrativas[0]

	print(f"Usando coluna de código na planilha atualizada: {col_codigo_atualizada}")
	print(f"Usando coluna de código na base TOTVS: {col_codigo_base}")
	print(f"Usando coluna de narrativa na base TOTVS: {col_narrativa}")

	# Criar um mapa codigo -> narrativa a partir da base TOTVS
	serie_narrativa = (
		df_base_dados_TOTVS[[col_codigo_base, col_narrativa]]
		.drop_duplicates(subset=[col_codigo_base])
		.set_index(col_codigo_base)[col_narrativa]
	)

	# Preencher/atualizar a coluna "SAP123" na planilha atualizada
	# Mantendo as duas primeiras linhas de título intactas
	col_destino_narrativa = "SAP123"
	primeira_linha_dados = 2  # índices 0 e 1 são títulos
	df_planilha_atualizada.loc[
		primeira_linha_dados:,
		col_destino_narrativa,
	] = df_planilha_atualizada.loc[
		primeira_linha_dados:,
		col_codigo_atualizada,
	].map(serie_narrativa)

	# Salvar a própria planilha atualizada com a coluna SAP123 preenchida
	df_planilha_atualizada.to_excel(caminho_planilha_atualizada, index=False)
	print("Coluna SAP123 preenchida com a narrativa na planilha atualizada.")


