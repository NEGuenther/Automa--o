import pandas as pd
from thefuzz import process, fuzz

# Função para carregar o dicionário de materiais do arquivo
def carregar_dicionario_size_dimension(caminho_dicionario):
    """
    Carrega o dicionário de size dimension a partir de um arquivo de texto.
    :param caminho_dicionario: Caminho para o arquivo de texto contendo os materiais.
    :return: Um conjunto (set) com os materiais.
    """
    with open(caminho_dicionario, 'r', encoding='utf-8') as arquivo:
        materiais = {linha.strip() for linha in arquivo if linha.strip()}  # Remove linhas vazias e espaços extras
    return materiais

    # Função para encontrar a melhor norma correspondente
def encontrar_size_dimension(narrativa, size_dimension):
    """
    Encontra o size dimension que melhor corresponde à narrativa.
    :param narrativa: A narrativa a ser comparada.
    :param size_dimension: O conjunto de size dimensions disponíveis.
    :return: O material correspondente ou None se a pontuação for baixa.
    """
    # Validar se narrativa é string válida
    if not isinstance(narrativa, str) or not narrativa.strip():
        return None

    # Converter narrativa para maiúsculas para comparação
    narrativa_upper = narrativa.upper()
    
    # Primeiro, tentar encontrar norma que aparecem como substring na narrativa
    materiais_encontrados = []
    for material in size_dimension:
        material_upper = material.upper()
        if material_upper in narrativa_upper:
            materiais_encontrados.append(material)
    
    # Se encontrou size_dimension por substring, retornar o mais longo (mais específico)
    if materiais_encontrados:
        selecionado = max(materiais_encontrados, key=len)
        return selecionado
    
    # Se não encontrou por substring, usar fuzzy matching
    melhor_material, pontuacao = process.extractOne(narrativa, size_dimension, scorer=fuzz.ratio)
    if melhor_material and melhor_material.upper():
        return "material nao informado"
    return melhor_material if pontuacao > 80 else None  # Retorna o material se a pontuação for maior que 80

if __name__ == "__main__":
    # Caminhos dos arquivos
    caminho_dicionario = "dados/dicionario_size_dimension.csv"
    caminho_planilha = "planilhas/base_dados_TOTVS.xlsx"
    caminho_saida = "planilhas/planilha_atualizada.xlsx"

    # Carregar o dicionário de materiais
    materiais = carregar_dicionario_size_dimension(caminho_dicionario)

    # Carregar a planilha
    df = pd.read_excel(caminho_planilha)

    # Verificar se a coluna SAP15 existe
    if "SAP15" not in df.columns:
        raise ValueError("A coluna 'SAP15' não foi encontrada na planilha.")

    # Processar cada narrativa e encontrar o material correspondente
    #df["material_encontrado"] = df["Coluna4"].apply(lambda narrativa: encontrar_material(narrativa, materiais))
    # Salvar a planilha atualizada
    df.to_excel(caminho_saida, index=False)

    print(f"Planilha atualizada salva em: {caminho_saida}")
