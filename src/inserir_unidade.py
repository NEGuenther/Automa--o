import pandas as pd

def inserir_unidade(
	caminho_planilha_atualizada: str,
	caminho_base_totvs: str,
):
	"""Compara códigos e insere o product group (SAP5) na planilha atualizada."""

	print("Inserindo product group (SAP5)...")

	# Ler arquivos de entrada
	df_planilha_atualizada = pd.read_excel(caminho_planilha_atualizada)
	df_base_dados_TOTVS = pd.read_excel(caminho_base_totvs, header=4)

	# Identificar colunas de código
	# Na planilha atualizada, usamos a primeira coluna (códigos dos itens)
	col_codigo_atualizada = df_planilha_atualizada.columns[0]

	# Na base TOTVS, a coluna de código costuma ser "item" (mesmo usada para SAP123)
	col_codigo_base = None
	for c in df_base_dados_TOTVS.columns:
		if str(c).strip().lower() == "item":
			col_codigo_base = c
			break

	if col_codigo_base is None:
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

	# Identificar coluna de unidade
	col_unidade = None
	for c in df_base_dados_TOTVS.columns:
		if str(c).strip().lower() == "un":
			col_unidade = c
			break
	
	if col_unidade is None:
		# fallback: tenta procurar por "unidade"
		possiveis_unidades = [
			c for c in df_base_dados_TOTVS.columns if "unidade" in c.lower()
		]
		if possiveis_unidades:
			col_unidade = possiveis_unidades[0]
		else:
			raise ValueError(
				"Não foi encontrada nenhuma coluna de 'Unidade' na baseDadosTOTVS.xlsx. "
				"Verifique o nome das colunas."
			)

	# Criar um mapa codigo -> unidade a partir da base TOTVS
	serie_unidade = (
		df_base_dados_TOTVS[[col_codigo_base, col_unidade]]
		.drop_duplicates(subset=[col_codigo_base])
		.set_index(col_codigo_base)[col_unidade]
	)

	# Preencher/atualizar a coluna "SAP5" na planilha atualizada
	# Mantendo as duas primeiras linhas de título intactas
	col_destino_unidade = "SAP5"
	primeira_linha_dados = 1  # índices 0 e 1 são títulos
	df_planilha_atualizada.loc[
		primeira_linha_dados:,
		col_destino_unidade,
	] = df_planilha_atualizada.loc[
		primeira_linha_dados:,
		col_codigo_atualizada,
	].map(serie_unidade)

	# Estatísticas
	preenchidas = int(df_planilha_atualizada.loc[primeira_linha_dados:, col_destino_unidade].notna().sum())
	total_linhas = int(len(df_planilha_atualizada.index) - primeira_linha_dados)

	# Salvar a própria planilha atualizada com a coluna SAP5 preenchida
	df_planilha_atualizada.to_excel(caminho_planilha_atualizada, index=False)
	print(f"Unidade inserida: {preenchidas}/{total_linhas} linhas com SAP5")