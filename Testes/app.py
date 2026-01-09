"""Pipeline de preparação de planilhas para cargas SAP/TOTVS.

Fluxo principal:
- Gera planilha base a partir do modelo e CSV de códigos.
- Enriquecimento com comentários internos, product group e unidade.
- Preenchimento de colunas derivadas por narrativa: materiais, normas e size dimension.
- Aplicação de valores fixos e ajustes em narrativas longas.
"""

import sys
from pathlib import Path

import pandas as pd

# Caminhos base usados em todo o pipeline de planilhas
BASE_DIR = Path(__file__).resolve().parent.parent
SRC_DIR = BASE_DIR / "src"
PLANILHA_MODELO = BASE_DIR / "planilhas/planilhaPadrao.xlsx"
CSV_CODIGOS = BASE_DIR / "planilhas/dados_teste.csv"
PLANILHA_SAIDA = BASE_DIR / "planilhas/planilha_atualizada.xlsx"
BASE_TOTVS = BASE_DIR / "planilhas/base_dados_TOTVS.xlsx"
DICIONARIO_MATERIAIS = BASE_DIR / "dados/dicionario_materiais.csv"
DICIONARIO_NORMAS = BASE_DIR / "dados/dicionario_normas.csv"
DICIONARIO_SIZE_DIMENSION = BASE_DIR / "dados/dicionario_size_dimension.csv"
DICIONARIO_TRADUCOES = BASE_DIR / "dados/dicionario.xlsx"

# Garante que os módulos em src sejam encontrados
if str(SRC_DIR) not in sys.path:
	sys.path.insert(0, str(SRC_DIR))

from inserir_codigos_de_itens import gerar_planilha_com_codigos
from inserir_internal_comment import inserir_internal_coments
from inserir_unidade import inserir_unidade
from inserir_traducoes import inserir_traducoes
from inserir_material import carregar_dicionario, encontrar_material
from inserir_valores_fixos import inserir_valores_fixos
from inserir_narrativas import inserir_narrativa
from inserir_product_group import inserir_product_group
from inserir_normas import encontrar_normas, carregar_dicionario_normas
from inserir_size_dimension import carregar_dicionario_size_dimension, encontrar_size_dimension


def gerar_planilha_base(modelo: Path, csv_codigos: Path, saida: Path) -> Path:
	"""Gera a planilha inicial a partir do modelo e do CSV de códigos, se ambos existirem.

	Args:
		modelo: caminho do arquivo xlsx modelo.
		csv_codigos: caminho do CSV com códigos para preencher o modelo.
		saida: caminho desejado para salvar o resultado.

	Returns:
		Path para a planilha (existente ou recém-gerada).
	"""
	if not (modelo.exists() and csv_codigos.exists()):
		print("Pulando geração da planilha base")
		return saida

	gerar_planilha_com_codigos(
		caminho_planilha_modelo=str(modelo),
		caminho_csv_codigos=str(csv_codigos),
		caminho_saida=str(saida),
	)
	print(f"Planilha base gerada: {saida}")
	return saida


def garantir_planilha_saida(saida: Path) -> Path:
	"""Garante que há um arquivo de trabalho retornando o caminho válido.

	Args:
		saida: caminho esperado para a planilha.

	Returns:
		Caminho existente para a planilha de trabalho.

	Raises:
		SystemExit: quando nenhuma planilha válida é encontrada.
	"""
	if saida.exists():
		return saida

	# Fallback para o nome padrão esperado pelo restante das funções
	fallback = BASE_DIR / "planilhas/planilha_atualizada.xlsx"
	if fallback.exists():
		return fallback

	print("Erro: arquivo de trabalho inexistente: planilhas/planilha_atualizada.xlsx")
	raise SystemExit(1)


def atualizar_coluna_por_narrativa(df: pd.DataFrame, coluna_destino: str, linha_inicial: int, busca_fn) -> int:
	"""Preenche uma coluna baseada na narrativa SAP123 usando a função de busca fornecida.

	Args:
		df: dataframe carregado da planilha.
		coluna_destino: coluna a ser preenchida/atualizada.
		linha_inicial: índice inicial para processamento (pula cabeçalhos ou linhas fixas).
		busca_fn: função que recebe a narrativa e retorna o valor a ser escrito.

	Returns:
		Quantidade de valores preenchidos (não nulos) a partir da linha inicial.
	"""
	if "SAP123" not in df.columns:
		print("Aviso: coluna 'SAP123' não encontrada na planilha.")
		return 0

	if coluna_destino not in df.columns:
		df[coluna_destino] = None

	for idx in range(linha_inicial, len(df)):
		narrativa = df.loc[idx, "SAP123"]
		df.loc[idx, coluna_destino] = busca_fn(narrativa)

	encontrados = df.loc[linha_inicial:, coluna_destino].notna().sum()
	return int(encontrados)


def processar_materiais(saida: Path) -> None:
	"""Preenche Coluna4 com materiais correspondentes às narrativas."""
	print("Processando materiais (matching por narrativa)...")
	materiais = carregar_dicionario(str(DICIONARIO_MATERIAIS))
	print(f"Materiais carregados: {len(materiais)} entradas")

	df = pd.read_excel(str(saida))
	encontrados = atualizar_coluna_por_narrativa(
		df,
		coluna_destino="Coluna4",
		linha_inicial=2,
		busca_fn=lambda narrativa: encontrar_material(narrativa, materiais),
	)
	print(f"Materiais encontrados: {encontrados}")
	df.to_excel(str(saida), index=False)
	print("Coluna4 atualizada e salva na planilha.")


def processar_normas(saida: Path) -> None:
	"""Preenche SAP17 com normas vinculadas às narrativas."""
	print("Processando normas (matching por narrativa)...")
	normas = carregar_dicionario_normas(str(DICIONARIO_NORMAS))
	print(f"Normas carregadas: {len(normas)} entradas")

	df = pd.read_excel(str(saida))
	encontrados = atualizar_coluna_por_narrativa(
		df,
		coluna_destino="SAP17",
		linha_inicial=1,
		busca_fn=lambda narrativa: encontrar_normas(narrativa, normas),
	)
	print(f"Normas encontradas: {encontrados}")
	df.to_excel(str(saida), index=False)
	print("SAP17 atualizada e salva na planilha.")


def processar_size_dimension(saida: Path) -> None:
	"""Preenche SAP15 com size dimensions encontradas por narrativa."""
	print("Processando size dimensions (matching por narrativa)...")
	size_dimensions = carregar_dicionario_size_dimension(str(DICIONARIO_SIZE_DIMENSION))
	print(f"Size dimensions carregadas: {len(size_dimensions)} entradas")

	df = pd.read_excel(str(saida))
	encontrados = atualizar_coluna_por_narrativa(
		df,
		coluna_destino="SAP15",
		linha_inicial=1,
		busca_fn=lambda narrativa: encontrar_size_dimension(narrativa, size_dimensions),
	)
	print(f"Size dimensions encontradas: {encontrados}")
	df.to_excel(str(saida), index=False)
	print("SAP15 atualizada e salva na planilha.")


def inserir_valores_fixos_planilha(saida: Path) -> None:
	"""Aplica valores fixos nas colunas SAP10 e SAP14 da planilha de trabalho."""
	print("Aplicando valores fixos em SAP10 e SAP14...")
	inserir_valores_fixos(
		caminho_planilha_modelo=str(saida),
		caminho_saida=str(saida),
	)


def ajustar_narrativas(saida: Path) -> None:
	"""Atualiza SAP15 quando a narrativa SAP123 excede 141 caracteres."""
	print("Ajustando SAP15 para narrativas maior que 141 caracteres...")
	inserir_narrativa(
		caminho_planilha_modelo=str(saida),
		caminho_saida=str(saida),
	)
	print("Atualização de SAP15 concluída.")


def processar_traducoes(saida: Path) -> None:
	"""Processa traduções das descrições de produtos.
	
	Chama a função do módulo inserir_traducoes para preencher as colunas
	SAP1, SAP2, SAP3 e Coluna32 com traduções em português, inglês, espanhol e alemão.
	"""
	inserir_traducoes(
		caminho_planilha_atualizada=str(saida),
		caminho_base_totvs=str(BASE_TOTVS),
		caminho_dicionario_traducoes=str(DICIONARIO_TRADUCOES),
	)


def main() -> None:
	"""Orquestra o pipeline de geração, enriquecimento e ajustes da planilha."""
	saida = gerar_planilha_base(PLANILHA_MODELO, CSV_CODIGOS, PLANILHA_SAIDA)
	saida = garantir_planilha_saida(saida)

	inserir_internal_coments(
		caminho_planilha_atualizada=str(saida),
		caminho_base_totvs=str(BASE_TOTVS),
	)

	print("Preenchendo product group...")
	inserir_product_group(
		caminho_planilha_atualizada=str(saida),
		caminho_base_totvs=str(BASE_TOTVS),
	)
	print("Product group preenchido.")

	print("Inserindo unidade...")
	inserir_unidade(
		caminho_planilha_atualizada=str(saida),
		caminho_base_totvs=str(BASE_TOTVS),
	)
	print("Unidade inserida.")

	processar_materiais(saida)
	processar_normas(saida)
	processar_size_dimension(saida)
	processar_traducoes(saida)
	inserir_valores_fixos_planilha(saida)
	ajustar_narrativas(saida)


if __name__ == "__main__":
	main()
