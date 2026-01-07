import pandas as pd
from thefuzz import process, fuzz

# Função para carregar o dicionário de materiais do arquivo
def carregar_dicionario(caminho_dicionario):
    """
    Carrega o dicionário de materiais a partir de um arquivo de texto.
    :param caminho_dicionario: Caminho para o arquivo de texto contendo os materiais.
    :return: Um conjunto (set) com os materiais.
    """
    with open(caminho_dicionario, 'r', encoding='utf-8') as arquivo:
        materiais = {linha.strip() for linha in arquivo if linha.strip()}  # Remove linhas vazias e espaços extras
    return materiais

# Função para encontrar o melhor material correspondente
def encontrar_material(narrativa, materiais):
    """
    Encontra o material que melhor corresponde à narrativa.
    :param narrativa: A narrativa a ser comparada.
    :param materiais: O conjunto de materiais disponíveis.
    :return: O material correspondente ou None se a pontuação for baixa.
    """
    melhor_material, pontuacao = process.extractOne(narrativa, materiais, scorer=fuzz.ratio)
    return melhor_material if pontuacao > 80 else None  # Retorna o material se a pontuação for maior que 80

# Caminhos dos arquivos
caminho_dicionario = "DADOS/Dicionario_Materiais.txt"
caminho_planilha = "planilha/base_dados_TOTVS.xlsx"
caminho_saida = "planilha/planilha_atualizada.xlsx"

# Carregar o dicionário de materiais
materiais = carregar_dicionario(caminho_dicionario)

# Carregar a planilha
df = pd.read_excel(caminho_planilha)

# Verificar se a coluna WRKST existe
if "WRKST" not in df.columns:
    raise ValueError("A coluna 'WRKST' não foi encontrada na planilha.")

# Processar cada narrativa e encontrar o material correspondente
df["material_encontrado"] = df["WRKST"].apply(lambda narrativa: encontrar_material(narrativa, materiais))

# Salvar a planilha atualizada
df.to_excel(caminho_saida, index=False)

print(f"Planilha atualizada salva em: {caminho_saida}")
