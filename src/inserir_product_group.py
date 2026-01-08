import pandas as pd
from unidecode import unidecode
import re

def inserir_internal_coments(
	caminho_planilha_atualizada: str,
	caminho_base_totvs: str,
):
	"""Compara códigos e insere a product group na planilha atualizada.

	- Lê a planilha atualizada gerada.
	- Lê a base TOTVS (caminho informado) com códigos e coluna de product group.
	- Usa a primeira coluna da planilha atualizada como código e cruza com a coluna "item" (ou similar) da base TOTVS.
	- Preenche a coluna "SAP6" da planilha atualizada com o product group correspondente.
	"""

	print("Inserindo product group (SAP6)...")

	# Ler arquivos de entrada
	df_planilha_atualizada = pd.read_excel(caminho_planilha_atualizada)
	df_base_dados_TOTVS = pd.read_excel(caminho_base_totvs, header=4)

	# Identificar colunas de código
	# Na planilha atualizada, usamos a primeira coluna (códigos dos itens)
	col_codigo_atualizada = df_planilha_atualizada.columns[0]

	# Na base TOTVS, procuramos uma coluna cujo nome (ignorando maiúsculas/minúsculas)
	# seja exatamente "SAP6"
	col_codigo_base = None
	for c in df_base_dados_TOTVS.columns:
		if str(c).strip().lower() == "sap6":
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

	# Identificar coluna de product group (Fam Coml)
	col_product_group = None
	for c in df_base_dados_TOTVS.columns:
		if str(c).strip().lower() == "fam coml":
			col_product_group = c
			break
	
	if col_product_group is None:
		# fallback: tenta procurar por "product group"
		possiveis_product_groups = [
			c for c in df_base_dados_TOTVS.columns if "product group" in c.lower() or "fam" in c.lower()
		]
		if possiveis_product_groups:
			col_product_group = possiveis_product_groups[0]
		else:
			raise ValueError(
				"Não foi encontrada nenhuma coluna de 'Fam Coml' ou 'product group' na baseDadosTOTVS.xlsx. "
				"Verifique o nome das colunas."
			)

	print(
		f"Mapeando product groups -> codigo_atualizada: '{col_codigo_atualizada}', codigo_base: '{col_codigo_base}', product group: '{col_product_group}'"
	)

	# Criar um mapa codigo -> product group a partir da base TOTVS
	serie_product_group = (
		df_base_dados_TOTVS[[col_codigo_base, col_product_group]]
		.drop_duplicates(subset=[col_codigo_base])
		.set_index(col_codigo_base)[col_product_group]
	)

	# Preencher/atualizar a coluna "SAP6" na planilha atualizada
	# Mantendo as duas primeiras linhas de título intactas
	col_destino_product_group = "SAP6"
	primeira_linha_dados = 2  # índices 0 e 1 são títulos
	df_planilha_atualizada.loc[
		primeira_linha_dados:,
		col_destino_product_group,
	] = df_planilha_atualizada.loc[
		primeira_linha_dados:,
		col_codigo_atualizada,
	].map(serie_product_group)

	# Estatísticas
	preenchidas = int(df_planilha_atualizada.loc[primeira_linha_dados:, col_destino_product_group].notna().sum())
	total_linhas = int(len(df_planilha_atualizada.index) - primeira_linha_dados)

	# Salvar a própria planilha atualizada com a coluna SAP6 preenchida
	df_planilha_atualizada.to_excel(caminho_planilha_atualizada, index=False)
	print(f"Product group inserido: {preenchidas}/{total_linhas} linhas com SAP6")