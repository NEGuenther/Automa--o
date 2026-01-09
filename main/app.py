"""Pipeline de preparação de planilhas para cargas SAP/TOTVS.

Fluxo principal:
- Gera planilha base a partir do modelo e CSV de códigos.
- Enriquecimento com comentários internos, product group e unidade.
- Preenchimento de colunas derivadas por narrativa: materiais, normas e size dimension.
- Aplicação de valores fixos e ajustes em narrativas longas.
"""

import json
import platform
import sys
import time
import traceback
from datetime import datetime
from pathlib import Path

import pandas as pd

# Caminhos base usados em todo o pipeline de planilhas
BASE_DIR = Path(__file__).resolve().parent.parent
SRC_DIR = BASE_DIR / "src"
PLANILHA_MODELO = BASE_DIR / "planilhas/planilha_padrao.xlsx"
CSV_CODIGOS = BASE_DIR / "dados/dados_teste.csv"
PLANILHA_SAIDA = BASE_DIR / "planilhas/planilha_atualizada.xlsx"
BASE_TOTVS = BASE_DIR / "planilhas/base_dados_TOTVS.xlsx"
DICIONARIO_MATERIAIS = BASE_DIR / "dados/dicionario_materiais.csv"
DICIONARIO_NORMAS = BASE_DIR / "dados/dicionario_normas.csv"
DICIONARIO_SIZE_DIMENSION = BASE_DIR / "dados/dicionario_size_dimension.csv"
DICIONARIO_TRADUCOES = BASE_DIR / "dados/dicionario.xlsx"
LOGS_DIR = BASE_DIR / "logs"
RELATORIO_EXECUCAO = LOGS_DIR / "relatorio_execucao.json"

# A planilha gerada tem:
# - linha 1: header
# - linha 2: linha descritiva (vira df index 0)
# - itens a partir da linha 3 (vira df index 1)
PRIMEIRA_LINHA_ITENS_DF = 1

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


def _now_iso() -> str:
	return datetime.now().astimezone().isoformat(timespec="seconds")


def _norm_col_name(value: object) -> str:
	return "".join(str(value or "").split()).upper()


def _find_col(df: pd.DataFrame, wanted: str) -> str | None:
	wanted_n = _norm_col_name(wanted)
	for c in df.columns:
		if _norm_col_name(c) == wanted_n:
			return c
	# fallback: alguns arquivos vêm com sufixos tipo _X000D_
	for c in df.columns:
		c_n = _norm_col_name(c)
		if c_n.startswith(wanted_n):
			return c
	return None


def _count_nonempty_column(excel_path: Path, column_name: str, start_idx: int) -> int:
	df = pd.read_excel(str(excel_path))
	col = _find_col(df, column_name)
	if col is None:
		return 0
	serie = df.loc[start_idx:, col]
	as_text = serie.astype(str).str.strip()
	return int((serie.notna() & (as_text != "") & (as_text.str.lower() != "nan")).sum())


def _count_equals(excel_path: Path, column_name: str, value: str, start_idx: int) -> int:
	df = pd.read_excel(str(excel_path))
	col = _find_col(df, column_name)
	if col is None:
		return 0
	serie = df.loc[start_idx:, col].astype(str).str.strip()
	return int((serie == value).sum())


def _write_report(report: dict) -> None:
	LOGS_DIR.mkdir(parents=True, exist_ok=True)
	RELATORIO_EXECUCAO.write_text(
		json.dumps(report, ensure_ascii=False, indent=2),
		encoding="utf-8",
	)


def gerar_planilha_base(modelo: Path, csv_codigos: Path, saida: Path) -> Path:
	"""Gera a planilha inicial a partir do modelo e do CSV de códigos, se ambos existirem."""
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
	"""Garante que há um arquivo de trabalho retornando o caminho válido."""
	if saida.exists():
		return saida

	fallback = BASE_DIR / "planilhas/planilha_atualizada.xlsx"
	if fallback.exists():
		return fallback

	print("Erro: arquivo de trabalho inexistente: planilhas/planilha_atualizada.xlsx")
	raise SystemExit(1)


def atualizar_coluna_por_narrativa(df: pd.DataFrame, coluna_destino: str, linha_inicial: int, busca_fn) -> int:
	"""Preenche uma coluna baseada na narrativa SAP123 usando a função de busca fornecida."""
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
		linha_inicial=PRIMEIRA_LINHA_ITENS_DF,
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
		linha_inicial=PRIMEIRA_LINHA_ITENS_DF,
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
		linha_inicial=PRIMEIRA_LINHA_ITENS_DF,
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
	"""Marca a coluna 'Narrativa' quando SAP123 excede 141 caracteres."""
	print("Ajustando coluna 'Narrativa' para SAP123 > 141 caracteres...")
	inserir_narrativa(
		caminho_planilha_modelo=str(saida),
		caminho_saida=str(saida),
	)
	print("Atualização de 'Narrativa' concluída.")


def processar_traducoes(saida: Path) -> None:
	"""Processa traduções das descrições de produtos."""
	inserir_traducoes(
		caminho_planilha_atualizada=str(saida),
		caminho_base_totvs=str(BASE_TOTVS),
		caminho_dicionario_traducoes=str(DICIONARIO_TRADUCOES),
	)


def main() -> None:
	"""Orquestra o pipeline de geração, enriquecimento e ajustes da planilha."""
	report: dict = {
		"run_started_at": _now_iso(),
		"environment": {
			"python": sys.version.split()[0],
			"platform": platform.platform(),
		},
		"paths": {
			"modelo": str(PLANILHA_MODELO),
			"csv_codigos": str(CSV_CODIGOS),
			"base_totvs": str(BASE_TOTVS),
			"saida": str(PLANILHA_SAIDA),
			"relatorio": str(RELATORIO_EXECUCAO),
		},
		"steps": [],
		"status": "in_progress",
	}

	def run_step(name: str, fn, metrics_fn=None) -> None:
		step = {"name": name, "started_at": _now_iso()}
		t0 = time.perf_counter()
		try:
			fn()
			step["status"] = "ok"
			if metrics_fn is not None:
				step["metrics"] = metrics_fn() or {}
		except Exception as exc:
			step["status"] = "error"
			step["error"] = {
				"type": type(exc).__name__,
				"message": str(exc),
				"traceback": traceback.format_exc(),
			}
			report["status"] = "error"
			raise
		finally:
			step["duration_seconds"] = round(time.perf_counter() - t0, 3)
			step["finished_at"] = _now_iso()
			report["steps"].append(step)
			_write_report(report)

	saida: Path | None = None

	def _step_gerar_planilha_base() -> None:
		nonlocal saida
		saida = gerar_planilha_base(PLANILHA_MODELO, CSV_CODIGOS, PLANILHA_SAIDA)
		saida = garantir_planilha_saida(saida)

	run_step(
		"gerar_planilha_base",
		_step_gerar_planilha_base,
		metrics_fn=lambda: {
			"saida_existe": bool(saida and Path(saida).exists()),
			"linhas_csv_codigos": int(pd.read_csv(CSV_CODIGOS, header=None).shape[0]) if CSV_CODIGOS.exists() else 0,
		},
	)
	assert saida is not None
	# atualiza caminho de saída efetivo (caso fallback seja usado)
	report["paths"]["saida"] = str(saida)
	_write_report(report)

	run_step(
		"inserir_internal_comment",
		lambda: inserir_internal_coments(
			caminho_planilha_atualizada=str(saida),
			caminho_base_totvs=str(BASE_TOTVS),
		),
		metrics_fn=lambda: {
			"sap123_preenchidos": _count_nonempty_column(saida, "SAP123", PRIMEIRA_LINHA_ITENS_DF),
		},
	)

	run_step(
		"inserir_product_group",
		lambda: inserir_product_group(
			caminho_planilha_atualizada=str(saida),
			caminho_base_totvs=str(BASE_TOTVS),
		),
		metrics_fn=lambda: {
			"sap6_preenchidos": _count_nonempty_column(saida, "SAP6", PRIMEIRA_LINHA_ITENS_DF),
		},
	)

	run_step(
		"inserir_unidade",
		lambda: inserir_unidade(
			caminho_planilha_atualizada=str(saida),
			caminho_base_totvs=str(BASE_TOTVS),
		),
		metrics_fn=lambda: {
			"sap5_preenchidos": _count_nonempty_column(saida, "SAP5", PRIMEIRA_LINHA_ITENS_DF),
		},
	)

	run_step(
		"processar_materiais",
		lambda: processar_materiais(saida),
		metrics_fn=lambda: {
			"coluna4_preenchidos": _count_nonempty_column(saida, "Coluna4", PRIMEIRA_LINHA_ITENS_DF),
		},
	)

	run_step(
		"processar_normas",
		lambda: processar_normas(saida),
		metrics_fn=lambda: {
			"sap17_preenchidos": _count_nonempty_column(saida, "SAP17", PRIMEIRA_LINHA_ITENS_DF),
		},
	)

	run_step(
		"processar_size_dimension",
		lambda: processar_size_dimension(saida),
		metrics_fn=lambda: {
			"sap15_preenchidos": _count_nonempty_column(saida, "SAP15", PRIMEIRA_LINHA_ITENS_DF),
		},
	)

	run_step(
		"processar_traducoes",
		lambda: processar_traducoes(saida),
		metrics_fn=lambda: {
			"sap1_preenchidos": _count_nonempty_column(saida, "SAP1", PRIMEIRA_LINHA_ITENS_DF),
			"sap2_preenchidos": _count_nonempty_column(saida, "SAP2", PRIMEIRA_LINHA_ITENS_DF),
			"sap3_preenchidos": _count_nonempty_column(saida, "SAP3", PRIMEIRA_LINHA_ITENS_DF),
			"coluna32_preenchidos": _count_nonempty_column(saida, "Coluna32", PRIMEIRA_LINHA_ITENS_DF),
		},
	)

	run_step(
		"inserir_valores_fixos",
		lambda: inserir_valores_fixos_planilha(saida),
		metrics_fn=lambda: {
			"sap10_igual_10": _count_equals(saida, "SAP10", "10", PRIMEIRA_LINHA_ITENS_DF),
			"sap14_igual_NDB": _count_equals(saida, "SAP14", "NDB", PRIMEIRA_LINHA_ITENS_DF),
		},
	)

	run_step(
		"ajustar_narrativas",
		lambda: ajustar_narrativas(saida),
		metrics_fn=lambda: {
			"narrativa_marcada": _count_equals(
				saida,
				"Narrativa",
				"verificar internal comment",
				PRIMEIRA_LINHA_ITENS_DF,
			),
		},
	)

	report["status"] = "ok"
	report["run_finished_at"] = _now_iso()
	# duração total aproximada: soma das etapas
	report["duration_seconds"] = round(
		sum(float(s.get("duration_seconds", 0.0)) for s in report.get("steps", [])),
		3,
	)
	_write_report(report)


if __name__ == "__main__":
	main()
