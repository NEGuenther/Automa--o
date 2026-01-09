import pandas as pd

def LerPlanilhaPadrao():
    return pd.read_excel(r"planilhas/planilha_padrao.xlsx")

def LerPlanilhaAtualizada():
    return pd.read_excel(r"planilhas/planilha_atualizada.xlsx")


