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
    # Validar se narrativa é string válida
    if not isinstance(narrativa, str) or not narrativa.strip():
        return None

    materiais_bloqueados = {"MOTOR", "SPECIAL"}
    
    # Converter narrativa para maiúsculas para comparação
    narrativa_upper = narrativa.upper()
    
    # Primeiro, tentar encontrar materiais que aparecem como substring na narrativa
    materiais_encontrados = []
    for material in materiais:
        material_upper = material.upper()
        if material_upper in materiais_bloqueados:
            continue
        if material_upper in narrativa_upper:
            materiais_encontrados.append(material)
    
    # Se encontrou materiais por substring, retornar o mais longo (mais específico)
    if materiais_encontrados:
        selecionado = max(materiais_encontrados, key=len)
        return selecionado
    
    # Se não encontrou por substring, usar fuzzy matching
    melhor_material, pontuacao = process.extractOne(narrativa, materiais, scorer=fuzz.ratio)
    if melhor_material and melhor_material.upper() in materiais_bloqueados:
        return "material nao informado"
    return melhor_material if pontuacao > 80 else None  # Retorna o material se a pontuação for maior que 80

