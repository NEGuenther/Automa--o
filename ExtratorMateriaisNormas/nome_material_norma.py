import pandas as pd
from unidecode import unidecode
import tkinter as tk
from tkinter import filedialog, messagebox
import re

# Função para limpar caracteres invisíveis
def limpar_texto(txt):
    if not isinstance(txt, str):
        txt = str(txt)
    return (
        txt.replace("\xa0", " ")  
           .replace("\ufeff", "") 
           .replace("\u200b", "") 
           .strip()
    )

def processar_arquivos(excel_path, dict_path, txt_material_path, txt_normas_path, txt_nome_mat_path, sheet_name):
    try:
        # identifica a extensão do arquivo principal
        if excel_path.lower().endswith((".xlsx", ".xls")):
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
        elif excel_path.lower().endswith(".csv"):
            df = pd.read_csv(excel_path, encoding="utf-8", sep=";")  # ajuste sep se necessário
        else:
            raise ValueError("Formato de arquivo principal não suportado. Use Excel ou CSV.")

        # lê dicionário informado pelo usuário (mantém Excel)
        df_tr = pd.read_excel(dict_path, sheet_name="Plan1")

        # normaliza chave no dicionário
        df_tr["MAKTX_norm"] = df_tr["MAKTX(PT)"].astype(str).map(lambda x: unidecode(limpar_texto(x).upper()))
        df_tr = df_tr.drop_duplicates(subset=["MAKTX_norm"], keep="first")

        # ================ MATERIAL ================
        buscas_material = []
        abreviacoes_material = {}
        with open(txt_material_path, "r", encoding="utf-8") as f:
            for linha in f:
                linha = limpar_texto(linha)
                if not linha:
                    continue
                if "=" in linha:
                    chave, valor = linha.split("=", 1)
                    abreviacoes_material[unidecode(chave.upper())] = unidecode(valor.upper())
                else:
                    buscas_material.append(unidecode(linha.upper()))
        buscas_material.sort(key=len, reverse=True)

        print("\n=== LISTA DE BUSCAS MATERIAL ===")
        for termo in buscas_material:
            print(repr(termo))

        # ================ NORMAS ================
        buscas_normas = []
        abreviacoes_normas = {}
        with open(txt_normas_path, "r", encoding="utf-8") as f:
            for linha in f:
                linha = limpar_texto(linha)
                if not linha:
                    continue
                if "=" in linha:
                    chave, valor = linha.split("=", 1)
                    abreviacoes_normas[unidecode(chave.upper())] = unidecode(valor.upper())
                else:
                    buscas_normas.append(unidecode(linha.upper()))
        buscas_normas.sort(key=len, reverse=True)

        print("\n=== LISTA DE BUSCAS NORMAS ===")
        for termo in buscas_normas:
            print(repr(termo))

        # ================ NOME DO MATERIAL ================
        buscas_nome_mat = []
        abreviacoes_nome_mat = {}
        with open(txt_nome_mat_path, "r", encoding="utf-8") as f:
            for linha in f:
                linha = limpar_texto(linha)
                if not linha:
                    continue
                if "=" in linha:
                    chave, valor = linha.split("=", 1)
                    abreviacoes_nome_mat[unidecode(chave.upper())] = unidecode(valor.upper())
                else:
                    buscas_nome_mat.append(unidecode(linha.upper()))
        buscas_nome_mat.sort(key=len, reverse=True)

        print("\n=== LISTA DE BUSCAS NOME MATERIAL ===")
        for termo in buscas_nome_mat:
            print(repr(termo))

        # função para normalizar comentário e substituir abreviações
        def normalizar_comment(comment, abreviacoes):
            comment_norm = unidecode(limpar_texto(comment).upper())
            for abreviacao, completo in abreviacoes.items():
                # substitui apenas a palavra inteira
                pattern = r'\b' + re.escape(abreviacao) + r'\b'
                comment_norm = re.sub(pattern, completo, comment_norm)
            return comment_norm

        # normaliza comentários
        df['Comments_norm'] = df['Internal Comments'].astype(str).apply(
            lambda c: normalizar_comment(c, {**abreviacoes_material, **abreviacoes_normas, **abreviacoes_nome_mat})
        )

        print("\n=== EXEMPLOS DE COMMENTS NORMALIZADOS ===")
        for i, c in enumerate(df['Comments_norm'].head(10)):  # mostra só os 10 primeiros
            print(i, repr(c))

        # Essa função extrai os termos encontrados
        def extrair_termos(comment, buscas, apenas_um=False):
            encontrados = []
            for termo in buscas:
                if re.search(rf"\b{re.escape(termo)}\b", comment):
                    if apenas_um:
                        return termo
                    if termo not in encontrados:
                        encontrados.append(termo)
            return ", ".join(encontrados) if encontrados else "Verificar"

        # aplica extração
        df['Basic material '] = df['Comments_norm'].apply(lambda c: extrair_termos(c, buscas_material))
        df['Norma'] = df['Comments_norm'].apply(lambda c: extrair_termos(c, buscas_normas))
        df['MAKTX(PT)'] = df['Comments_norm'].apply(lambda c: extrair_termos(c, buscas_nome_mat, apenas_um=True))

        # cria DataFrame parcial
        df_resultados = df[['Código', 'Material', 'Norma', 'MAKTX(PT)', 'Internal Comments']]

        # normaliza chave antes do merge
        df_resultados["MAKTX_norm"] = df_resultados["MAKTX(PT)"].astype(str).map(lambda x: unidecode(limpar_texto(x).upper()))

        # merge com traduções (PT -> EN, ES, DE)
        df_final = pd.merge(
            df_resultados,
            df_tr[["MAKTX_norm", "MAKTX(EN)", "MAKTX(ES)", "MAKTX(DE)"]],
            on="MAKTX_norm",
            how="left"
        )

        # remove coluna auxiliar
        df_final = df_final.drop(columns=["MAKTX_norm"])

        # salva em Excel
        df_final.to_excel("AncoraT.xlsx", index=False)

        messagebox.showinfo("Sucesso", "Arquivo 'Resultados.xlsx' gerado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", str(e))


# Funções para selecionar arquivos
def selecionar_excel():
    path = filedialog.askopenfilename(
        filetypes=[("Planilhas", "*.xlsx;*.xls;*.csv"), ("Todos os arquivos", "*.*")]
    )
    entry_excel.delete(0, tk.END)
    entry_excel.insert(0, path)

def selecionar_dicionario():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    entry_dict.delete(0, tk.END)
    entry_dict.insert(0, path)

def selecionar_txt_material():
    path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    entry_txt_material.delete(0, tk.END)
    entry_txt_material.insert(0, path)

def selecionar_txt_norma():
    path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    entry_txt_norma.delete(0, tk.END)
    entry_txt_norma.insert(0, path)

def selecionar_txt_nome_mat():
    path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    entry_txt_nome_mat.delete(0, tk.END)
    entry_txt_nome_mat.insert(0, path)


#================ INTERFACE ================
root = tk.Tk() # 
root.title("Extração de Materiais e Normas")

# Excel principal
tk.Label(root, text="Arquivo Excel/CSV Principal:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_excel = tk.Entry(root, width=50)
entry_excel.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Selecionar", command=selecionar_excel).grid(row=0, column=2, padx=5, pady=5)

# Excel Dicionário
tk.Label(root, text="Arquivo Dicionário:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
entry_dict = tk.Entry(root, width=50)
entry_dict.grid(row=1, column=1, padx=5, pady=5)
tk.Button(root, text="Selecionar", command=selecionar_dicionario).grid(row=1, column=2, padx=5, pady=5)

# TXT Materiais
tk.Label(root, text="Dicionário Materiais:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry_txt_material = tk.Entry(root, width=50)
entry_txt_material.grid(row=2, column=1, padx=5, pady=5)
tk.Button(root, text="Selecionar", command=selecionar_txt_material).grid(row=2, column=2, padx=5, pady=5)

# TXT Normas
tk.Label(root, text="Dicionário Normas:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
entry_txt_norma = tk.Entry(root, width=50)
entry_txt_norma.grid(row=3, column=1, padx=5, pady=5)
tk.Button(root, text="Selecionar", command=selecionar_txt_norma).grid(row=3, column=2, padx=5, pady=5)

# TXT Nome Material
tk.Label(root, text="Dicionário Nome Material:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
entry_txt_nome_mat = tk.Entry(root, width=50)
entry_txt_nome_mat.grid(row=4, column=1, padx=5, pady=5)
tk.Button(root, text="Selecionar", command=selecionar_txt_nome_mat).grid(row=4, column=2, padx=5, pady=5)

# Sheet
tk.Label(root, text="Nome da aba (sheet):").grid(row=5, column=0, padx=5, pady=5, sticky="e")
entry_sheet = tk.Entry(root, width=50)
entry_sheet.grid(row=5, column=1, padx=5, pady=5)
entry_sheet.insert(0, "E")  # valor padrão

# Botão Processar
tk.Button(root, text="Processar", width=20,
          command=lambda: processar_arquivos(
              entry_excel.get(),
              entry_dict.get(),
              entry_txt_material.get(),
              entry_txt_norma.get(),
              entry_txt_nome_mat.get(),
              entry_sheet.get()
          )).grid(row=6, column=0, columnspan=3, pady=20)

root.mainloop()
