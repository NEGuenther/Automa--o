import sys
from pathlib import Path

# Adicionar a pasta src ao caminho de importação
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

import pandas as pd
from inserir_codigos_de_itens import gerar_planilha_com_codigos
from inserir_internal_comment import inserir_internal_coments
from inserir_traducoes import Traducoes
from inserir_material import carregar_dicionario, encontrar_material
from inserir_valores_fixos import inserir_valores_fixos
from inserir_narrativas import inserir_narrativa
from inserir_product_group import inserir_internal_coments as inserir_product_group
from inserir_normas import encontrar_normas, carregar_dicionario_normas
from inserir_size_dimension import carregar_dicionario_size_dimension, encontrar_size_dimension


# 1) Gera a planilha com códigos

modelo = Path("planilhas/planilhaPadrao.xlsx")
csv_codigos = Path("planilhas/dados_teste.csv")
saida = Path("planilhas/planilha_atualizada.xlsx")

if modelo.exists() and csv_codigos.exists():
    gerar_planilha_com_codigos(
    	caminho_planilha_modelo=str(modelo),
    	caminho_csv_codigos=str(csv_codigos),
    	caminho_saida=str(saida),
    )
    print(f"Planilha base gerada: {saida}")
else:
    print("Pulando geração da planilha base")

# 2) Insere dados (narrativa SAP123, etc.)
inserir_internal_coments(
	caminho_planilha_atualizada=r"planilhas/planilha_atualizada.xlsx",
	caminho_base_totvs=r"planilhas/base_dados_TOTVS.xlsx",
)

# 2.1) Preenche product group (SAP6)
print("Preenchendo product group (SAP6)...")
inserir_product_group(
	caminho_planilha_atualizada=r"planilhas/planilha_atualizada.xlsx",
	caminho_base_totvs=r"planilhas/base_dados_TOTVS.xlsx",
)
print("Product group (SAP6) preenchido.")

# 2.5) Carregar dicionário de materiais e encontrar materiais correspondentes
print("Processando materiais (matching por narrativa)...")
materiais = carregar_dicionario(r"dados/dicionario_materiais.csv")
print(f"Materiais carregados: {len(materiais)} entradas")

# Carregar a planilha atualizada
if not saida.exists():
	if Path("planilhas/planilha_atualizada.xlsx").exists():
		saida = Path("planilhas/planilha_atualizada.xlsx")
	else:
		print("Erro: arquivo de trabalho inexistente: planilhas/planilha_atualizada.xlsx")
		raise SystemExit(1)

df = pd.read_excel(str(saida))

# Verificar se a coluna SAP123 existe
if "SAP123" in df.columns:
	print("Processando correspondências na coluna SAP123...")
	
	# Criar a coluna Coluna4 se não existir, inicializar com None
	if "Coluna4" not in df.columns:
		df["Coluna4"] = None
	
	# Processar apenas da linha 2 em diante (índice 2)
	for idx in range(2, len(df)):
		narrativa = df.loc[idx, "SAP123"]
		material = encontrar_material(narrativa, materiais)
		df.loc[idx, "Coluna4"] = material
	
	# Contar quantos materiais foram encontrados (a partir da linha 2)
	materiais_encontrados = df.loc[2:, "Coluna4"].notna().sum()
	print(f"Materiais encontrados: {materiais_encontrados}")
	
	# Salvar a planilha atualizada
	df.to_excel(str(saida), index=False)
	print("Coluna4 atualizada e salva na planilha.")
else:
	print("Aviso: coluna 'SAP123' não encontrada na planilha.")

# 2.6) Carregar dicionário de normas e encontrar normas correspondentes
print("Processando normas (matching por narrativa)...")
normas = carregar_dicionario_normas(r"dados/dicionario_normas.csv")
print(f"Normas carregadas: {len(normas)} entradas")

# Recarregar a planilha
df = pd.read_excel(str(saida))

# Verificar se a coluna SAP123 existe
if "SAP123" in df.columns:
	print("Processando correspondências de normas na coluna SAP123...")
	
	# Criar a coluna SAP17 se não existir, inicializar com None
	if "SAP17" not in df.columns:
		df["SAP17"] = None
	
	# Processar apenas da linha 1 em diante (índice 1)
	for idx in range(1, len(df)):
		narrativa = df.loc[idx, "SAP123"]
		norma = encontrar_normas(narrativa, normas)
		df.loc[idx, "SAP17"] = norma
	
	# Contar quantas normas foram encontradas (a partir da linha 2)
	normas_encontradas = df.loc[2:, "SAP17"].notna().sum()
	print(f"Normas encontradas: {normas_encontradas}")
	
	# Salvar a planilha atualizada
	df.to_excel(str(saida), index=False)
	print("SAP17 atualizada e salva na planilha.")
# 2.7) Carregar dicionário de size dimension e encontrar size dimensions correspondentes
print("Processando size dimensions (matching por narrativa)...")
size_dimensions = carregar_dicionario_size_dimension(r"dados/dicionario_size_dimension.csv")
print(f"Size dimensions carregadas: {len(size_dimensions)} entradas")

# Recarregar a planilha
df = pd.read_excel(str(saida))

# Verificar se a coluna SAP123 existe
if "SAP123" in df.columns:
	print("Processando correspondências de size dimension na coluna SAP123...")
	
	# Criar a coluna SAP15 se não existir, inicializar com None
	if "SAP15" not in df.columns:
		df["SAP15"] = None
	
	# Processar apenas da linha 1 em diante (índice 1)
	for idx in range(1, len(df)):
		narrativa = df.loc[idx, "SAP123"]
		size_dim = encontrar_size_dimension(narrativa, size_dimensions)
		df.loc[idx, "SAP15"] = size_dim
	
	# Contar quantas size dimensions foram encontradas (a partir da linha 2)
	size_dims_encontradas = df.loc[2:, "SAP15"].notna().sum()
	print(f"Size dimensions encontradas: {size_dims_encontradas}")
	
	# Salvar a planilha atualizada
	df.to_excel(str(saida), index=False)
	print("SAP15 atualizada e salva na planilha.")
else:
	print("Aviso: coluna 'SAP123' não encontrada na planilha.")


# 3) Insere valores fixos nas colunas SAP10 e SAP14
print("Aplicando valores fixos em SAP10 e SAP14...")
inserir_valores_fixos(
	caminho_planilha_modelo=str(saida),
	caminho_saida=str(saida),
)

# 4) Verifica tamanho de SAP123 e atualiza SAP15
print("Ajustando SAP15 para narrativas maior que 144 caracteres...")
inserir_narrativa(
	caminho_planilha_modelo=str(saida),
	caminho_saida=str(saida),
)
print("Atualização de SAP15 concluída.")
