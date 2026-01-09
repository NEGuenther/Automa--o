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

	print("Inserindo internal comment (SAP123)...")

	# Ler arquivos de entrada
	df_planilha_atualizada = pd.read_excel(caminho_planilha_atualizada)
	df_base_dados_TOTVS = pd.read_excel(caminho_base_totvs, header=4)

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

	print(
		f"Mapeando narrativas -> codigo_atualizada: '{col_codigo_atualizada}', codigo_base: '{col_codigo_base}', narrativa: '{col_narrativa}'"
	)

	# Criar um mapa codigo -> narrativa a partir da base TOTVS
	serie_narrativa = (
		df_base_dados_TOTVS[[col_codigo_base, col_narrativa]]
		.drop_duplicates(subset=[col_codigo_base])
		.set_index(col_codigo_base)[col_narrativa]
	)

	# Preencher/atualizar a coluna "SAP123" na planilha atualizada
	# Planilha gerada: linha 1 = cabeçalho; primeira linha de dados (índice 0) é descritiva.
	# Itens começam a partir do índice 1.
	primeira_linha_dados = 1

	# Destino principal: SAP123 (internal comment)
	col_destino_sap123 = "SAP123" if "SAP123" in df_planilha_atualizada.columns else None
	if col_destino_sap123 is None:
		raise ValueError("A coluna 'SAP123' não foi encontrada na planilha atualizada.")

	valores_narrativa = df_planilha_atualizada.loc[primeira_linha_dados:, col_codigo_atualizada].map(serie_narrativa)
	df_planilha_atualizada.loc[primeira_linha_dados:, col_destino_sap123] = valores_narrativa

	# Se existir alguma coluna de narrativa (com quebras de linha/espacos), preencher também
	colunas_narrativa = [c for c in df_planilha_atualizada.columns if str(c).strip().lower() == "narrativa"]
	for col in colunas_narrativa:
		# Garante dtype compatível para strings
		if df_planilha_atualizada[col].dtype != object:
			df_planilha_atualizada[col] = df_planilha_atualizada[col].astype("object")
		df_planilha_atualizada.loc[primeira_linha_dados:, col] = valores_narrativa

	# Estatísticas (considera vazio e 'nan' como não preenchido)
	serie = df_planilha_atualizada.loc[primeira_linha_dados:, col_destino_sap123]
	serie_txt = serie.astype(str).str.strip().str.lower()
	preenchidas = int((serie.notna() & (serie_txt != "") & (serie_txt != "nan")).sum())
	total_linhas = int(len(df_planilha_atualizada.index) - primeira_linha_dados)

	# Garante que colunas auxiliares não fiquem no arquivo final
	if "Num_Chars" in df_planilha_atualizada.columns:
		df_planilha_atualizada = df_planilha_atualizada.drop(columns=["Num_Chars"])

	# Salvar a própria planilha atualizada com a coluna SAP123 preenchida
	df_planilha_atualizada.to_excel(caminho_planilha_atualizada, index=False)
	print(f"Internal comment inserido: {preenchidas}/{total_linhas} linhas com narrativa")


