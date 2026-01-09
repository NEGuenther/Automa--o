from pathlib import Path

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


def preencher_materiais(caminho_planilha: str | Path, caminho_dicionario: str | Path) -> None:
    """Preenche a coluna Coluna4 com materiais encontrados a partir de SAP123."""
    caminho_planilha = Path(caminho_planilha)
    materiais = carregar_dicionario(str(caminho_dicionario))

    df = pd.read_excel(caminho_planilha)
    if "SAP123" not in df.columns:
        print("Aviso: coluna 'SAP123' não encontrada na planilha.")
        return

    if "Coluna4" not in df.columns:
        df["Coluna4"] = None

    for idx in range(1, len(df)):
        narrativa = df.loc[idx, "SAP123"]
        df.loc[idx, "Coluna4"] = encontrar_material(narrativa, materiais)

    encontrados = df.loc[2:, "Coluna4"].notna().sum()
    print(f"Materiais encontrados: {encontrados}")

    df.to_excel(caminho_planilha, index=False)
    print("Coluna4 atualizada e salva na planilha.")

if __name__ == "__main__":
    # Caminhos dos arquivos
    caminho_dicionario = "dados/dicionario_materiais.csv"
    caminho_planilha = "planilhas/base_dados_TOTVS.xlsx"
    caminho_saida = "planilha/planilha_atualizada.xlsx"

    # Carregar o dicionário de materiais
    materiais = carregar_dicionario(caminho_dicionario)

    # Carregar a planilha
    df = pd.read_excel(caminho_planilha)

    # Verificar se a coluna Coluna4 existe
    if "Coluna4" not in df.columns:
        raise ValueError("A coluna 'Coluna4' não foi encontrada na planilha.")

    # Processar cada narrativa e encontrar o material correspondente
    #df["material_encontrado"] = df["Coluna4"].apply(lambda narrativa: encontrar_material(narrativa, materiais))
    # Salvar a planilha atualizada
    df.to_excel(caminho_saida, index=False)

    print(f"Planilha atualizada salva em: {caminho_saida}")

