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
from inserir_size_dimension import inserir_size_dimension


# 1) Gera a planilha com códigos
'''gerar_planilha_com_codigos(
	caminho_planilha_modelo=r"planilhas/planilhaPadrao.xlsx",
	caminho_csv_codigos=r"planilhas/dados_teste.csv",
	caminho_saida=r"planilhas/planilha_atualizada.xlsx",
)
print("Planilha atualizada gerada com sucesso.")

# 2) Insere dados (narrativa SAP123, etc.)
inserir_internal_coments(
	caminho_planilha_atualizada=r"planilhas/planilha_atualizada.xlsx",
	caminho_base_totvs=r"planilhas/base_dados_TOTVS.xlsx",
)
'''

# 2.5) Carregar dicionário de materiais e encontrar materiais correspondentes
print("Processando materiais...")
materiais = carregar_dicionario(r"dados/dicionario_materiais.csv")
print(f"Total de materiais carregados: {len(materiais)}")

# Carregar a planilha atualizada
df = pd.read_excel(r"planilhas/planilha_atualizada.xlsx")

# Verificar se a coluna SAP123 existe
if "SAP123" in df.columns:
	print(f"Processando coluna SAP123...")
	
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
	print(f"Total de materiais encontrados: {materiais_encontrados}")
	print(f"Primeiros 5 materiais encontrados (a partir da linha 3): {df.loc[2:6, 'Coluna4'].tolist()}")
	
	# Salvar a planilha atualizada
	df.to_excel(r"planilhas/planilha_atualizada.xlsx", index=False)
	print("Materiais encontrados e inseridos na Coluna4 com sucesso.")
else:
	print("Aviso: A coluna 'SAP123' não foi encontrada na planilha.")

# 3) Insere valores fixos nas colunas SAP10 e SAP1
inserir_valores_fixos(
	caminho_planilha_modelo=r"planilhas/planilha_atualizada.xlsx",
	caminho_saida=r"planilhas/planilha_atualizada.xlsx",
)
print("Valores fixos inseridos com sucesso.")

# 4) Verifica tamanho de SAP123 e atualiza SAP15
inserir_size_dimension(
	caminho_planilha_modelo=r"planilhas/planilha_atualizada.xlsx",
	caminho_saida=r"planilhas/planilha_atualizada.xlsx",
)
print("Atualização de SAP15 por tamanho aplicada com sucesso.")
