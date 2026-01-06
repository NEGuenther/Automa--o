import pandas as pd

# 1) Ler planilha padrão (modelo) e CSV de códigos
Df_Planilha_Padrao = pd.read_excel(r"planilhas/planilhaPadrao.xlsx")
print("Planilha Padrão carregada com sucesso.")

Df_Codigos_CSV = pd.read_csv(r"planilhas/test.csv", header=None, names=["CODIGO"])
print("CSV de Códigos carregado com sucesso.")

# 2) Usar a primeira linha da planilha como modelo
linha_modelo = Df_Planilha_Padrao.iloc[0]

# 3) Criar nova planilha: título na primeira linha, dados a partir da segunda
quantidade_codigos = len(Df_Codigos_CSV)
Df_Novo = pd.concat([linha_modelo.to_frame().T] * (quantidade_codigos + 1), ignore_index=True)

# 4) Deixar todas as outras colunas vazias a partir da segunda linha
coluna_codigo = "item(table) + it-codigo(field)"
colunas_outros = [c for c in Df_Novo.columns if c != coluna_codigo]
Df_Novo.loc[1:, colunas_outros] = ""

# 5) Preencher a coluna de código a partir da segunda linha
Df_Novo.loc[1:, coluna_codigo] = Df_Codigos_CSV["CODIGO"].values

print(Df_Codigos_CSV.head(20))
print(Df_Novo.head(20))

# 6) Salvar planilha atualizada
Df_Novo.to_excel(r"planilhas/planilha_atualizada.xlsx", index=False)
print("Planilha atualizada salva com sucesso.")