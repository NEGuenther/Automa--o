import pandas as pd
from unidecode import unidecode
import re


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
	# Mantendo as duas primeiras linhas de título intactas
	col_destino_narrativa = "Narrativa"
	primeira_linha_dados = 2  # índices 0 e 1 são títulos
	df_planilha_atualizada.loc[
		primeira_linha_dados:,
		col_destino_narrativa,
	] = df_planilha_atualizada.loc[
		primeira_linha_dados:,
		col_codigo_atualizada,
	].map(serie_narrativa)

	# Verifica se o numero de caracteres ultrapassa o limite de 141 
	if col_destino_narrativa in df_planilha_atualizada.columns:
		limite_caracteres = 141 # limite de caracteres
		num_chars = df_planilha_atualizada[col_destino_narrativa].astype(str).str.len()
		ultrapassam_limite = df_planilha_atualizada[num_chars > limite_caracteres]
		if not ultrapassam_limite.empty:
			col_destino_narrativa141 = "SAP123"
			df_planilha_atualizada.loc[
			primeira_linha_dados:,
			col_destino_narrativa141,
			] = df_planilha_atualizada.loc[
				primeira_linha_dados:,
				col_codigo_atualizada,
			].map(serie_narrativa)
		else:
			print("Todas as entradas estão dentro do limite de caracteres.")

	# Estatísticas
	preenchidas = int(df_planilha_atualizada.loc[primeira_linha_dados:, col_destino_narrativa].notna().sum())
	total_linhas = int(len(df_planilha_atualizada.index) - primeira_linha_dados)

	# Garante que colunas auxiliares não fiquem no arquivo final
	if "Num_Chars" in df_planilha_atualizada.columns:
		df_planilha_atualizada = df_planilha_atualizada.drop(columns=["Num_Chars"])

	# Salvar a própria planilha atualizada com a coluna SAP123 preenchida
	df_planilha_atualizada.to_excel(caminho_planilha_atualizada, index=False)
	print(f"Internal comment inserido: {preenchidas}/{total_linhas} linhas com narrativa")


def extrair_campos_narrativa(
	caminho_planilha_atualizada: str,
	caminho_dicionario_excel: str,
	caminho_txt_materiais: str,
	caminho_txt_normas: str,
	caminho_txt_nome_pt: str,
	coluna_narrativa: str | None = None,
):
	"""Extrai 3 campos a partir da narrativa usando dicionários.

	Campos gerados:
	- "Basic material" (lista de termos encontrados nos materiais)
	- "Norma" (lista de normas encontradas)
	- "MAKTX(PT)" (primeiro nome de material encontrado)

	Parâmetros:
	- caminho_planilha_atualizada: Excel já gerado no fluxo
	- caminho_dicionario_excel: Excel com traduções (usado aqui só para normalização opcional)
	- caminho_txt_materiais: arquivo TXT com termos de materiais (linhas; pode conter abreviações no formato "ABV=EXPANSO")
	- caminho_txt_normas: arquivo TXT com termos de normas
	- caminho_txt_nome_pt: arquivo TXT com nomes de material em PT
	- coluna_narrativa: nome da coluna com a narrativa; se None, tenta detectar entre ["Internal Comments", "Internal comment (narrative)", "SAP123"]
	"""

	# Carrega planilha
	df = pd.read_excel(caminho_planilha_atualizada)

	# Detecta coluna de narrativa, se não informada
	if coluna_narrativa is None:
		possiveis = ["Internal Comments", "Internal comment (narrative)", "SAP123"]
		coluna_narrativa = next((c for c in possiveis if c in df.columns), None)
		if coluna_narrativa is None:
			raise ValueError(
				"Não foi possível detectar a coluna de narrativa. Informe explicitamente via 'coluna_narrativa' ou garanta que exista uma das colunas: "
				+ ", ".join(possiveis)
			)

	# Carrega dicionário Excel (mantém como referência; não obrigatório para extração)
	try:
		df_tr = pd.read_excel(caminho_dicionario_excel)
	except Exception:
		df_tr = None

	def limpar_texto(txt):
		if not isinstance(txt, str):
			txt = str(txt)
		return (
			txt.replace("\xa0", " ")
			   .replace("\ufeff", "")
			   .replace("\u200b", "")
			   .strip()
		)

	# Carrega e prepara listas/abreviações dos TXT
	def carregar_txt_categorias(path):
		buscas = []
		abrevs = {}
		with open(path, "r", encoding="utf-8") as f:
			for linha in f:
				linha = limpar_texto(linha)
				if not linha:
					continue
				if "=" in linha:
					chave, valor = linha.split("=", 1)
					abrevs[unidecode(chave.upper())] = unidecode(valor.upper())
				else:
					buscas.append(unidecode(linha.upper()))
		buscas.sort(key=len, reverse=True)
		return buscas, abrevs

	buscas_material, abrevs_material = carregar_txt_categorias(caminho_txt_materiais)
	buscas_normas, abrevs_normas = carregar_txt_categorias(caminho_txt_normas)
	buscas_nome_pt, abrevs_nome_pt = carregar_txt_categorias(caminho_txt_nome_pt)

	# Normaliza comentário e substitui abreviações
	def normalizar_comment(comment):
		comment_norm = unidecode(limpar_texto(comment).upper())
		for abreviacao, completo in {**abrevs_material, **abrevs_normas, **abrevs_nome_pt}.items():
			pattern = r"\b" + re.escape(abreviacao) + r"\b"
			comment_norm = re.sub(pattern, completo, comment_norm)
		return comment_norm

	df["Comments_norm"] = df[coluna_narrativa].astype(str).apply(normalizar_comment)

	# Busca termos usando correspondência de palavra inteira
	def extrair_termos(comment, buscas, apenas_um=False):
		encontrados = []
		for termo in buscas:
			if re.search(rf"\b{re.escape(termo)}\b", comment):
				if apenas_um:
					return termo
				if termo not in encontrados:
					encontrados.append(termo)
		return ", ".join(encontrados) if encontrados else "Verificar"

	# Aplica extrações
	df["Basic material"] = df["Comments_norm"].apply(lambda c: extrair_termos(c, buscas_material))
	df["Norma"] = df["Comments_norm"].apply(lambda c: extrair_termos(c, buscas_normas))
	df["MAKTX(PT)"] = df["Comments_norm"].apply(lambda c: extrair_termos(c, buscas_nome_pt, apenas_um=True))

	# Remove coluna auxiliar
	df = df.drop(columns=["Comments_norm"])

	# Resumo
	n_basic_ok = int((df["Basic material"] != "Verificar").sum()) if "Basic material" in df.columns else 0
	n_norma_ok = int((df["Norma"] != "Verificar").sum()) if "Norma" in df.columns else 0
	n_maktx_ok = int((df["MAKTX(PT)"] != "Verificar").sum()) if "MAKTX(PT)" in df.columns else 0

	# Salva de volta na mesma planilha
	df.to_excel(caminho_planilha_atualizada, index=False)
	print(f"Campos extraídos -> Basic material={n_basic_ok}, Norma={n_norma_ok}, MAKTX(PT)={n_maktx_ok}")


