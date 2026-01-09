# Automação de Planilha SAP (a partir de base TOTVS)

Este projeto automatiza a montagem de uma planilha final (`planilhas/planilha_atualizada.xlsx`) usando:

- um **modelo** (template) de planilha SAP;
- uma lista de **códigos de item** (CSV);
- uma **base TOTVS** (`base_dados_TOTVS.xlsx`) contendo dados (descrição, unidade, família comercial, narrativa etc.);
- dicionários auxiliares para **materiais**, **normas**, **size dimension** e **traduções**.

O runner principal do pipeline está em `main/app.py`.

---

## Requisitos

- Python 3.x
- Dependências Python:
  - `pandas`
  - `openpyxl`
  - `thefuzz`

Instalação rápida:

```bash
pip install -r requirements.txt
```

> Observação: o projeto usa `pandas` para leitura/escrita em Excel e `openpyxl` para alterações célula-a-célula em passos específicos.

---

## Entradas e Saída

### Entradas

- `planilhas/planilha_padrao.xlsx`
  - Modelo com **14 colunas** (ver seção “Colunas e Regras de Preenchimento”).
  - Estrutura esperada:
    - **Linha 1 (Excel)**: cabeçalho técnico (nomes das colunas)
    - **Linha 2 (Excel)**: linha descritiva (permanece no arquivo final)

- `dados/dados_teste.csv`
  - CSV com uma coluna (sem header) contendo **apenas os códigos** dos itens.

- `planilhas/base_dados_TOTVS.xlsx`
  - Base TOTVS usada para enriquecer a planilha.
  - Leitura feita com `pd.read_excel(..., header=4)`, ou seja:
    - o cabeçalho “real” deve começar na **linha 5** do Excel.

- `dados/dicionario_materiais.csv`
  - Lista (uma por linha) de materiais que serão buscados na narrativa.

- `dados/dicionario_normas.csv`
  - Lista (uma por linha) de normas que serão buscadas na narrativa.

- `dados/dicionario_size_dimension.csv`
  - Lista (uma por linha) de size dimensions que serão buscadas na narrativa.

- `dados/dicionario.xlsx`
  - Dicionário de traduções com colunas:
    - `PORTUGUÊS`, `INGLÊS`, `ESPANHOL` e `ALEMÂO` (alguns arquivos podem vir como `ALEMÂO_x000d_` — o código normaliza isso).

### Saída

- `planilhas/planilha_atualizada.xlsx`
  - Resultado final após todos os enriquecimentos.

---

## Como executar

No diretório raiz do projeto:

```bash
python main/app.py
```

---

## Layout da planilha (importante)

A planilha gerada e processada pelo pipeline segue esta convenção:

- **Excel linha 1**: cabeçalhos (nomes técnicos)
- **Excel linha 2**: linha descritiva (explicação dos campos)
- **Excel linha 3 em diante**: itens (dados reais)

No `pandas` isso significa:

- `df.index == 0` → linha descritiva
- `df.index >= 1` → itens

---

## Processo (pipeline)

Ordem executada em `main/app.py`:

1. **Gerar planilha base** (`src/inserir_codigos_de_itens.py`)
   - Lê o modelo e cria `planilhas/planilha_atualizada.xlsx`.
   - Mantém a **linha descritiva**.
   - Preenche a primeira coluna com os códigos do CSV.

2. **Internal comment / narrativa** (`src/inserir_internal_comment.py`)
   - Cruza o código do item (primeira coluna) com a coluna `Item` da base TOTVS.
   - Preenche `SAP123` com o texto da coluna de narrativa do TOTVS.

3. **Product group** (`src/inserir_product_group.py`)
   - Preenche `SAP6` a partir da coluna `Fam Coml` (ou fallback por nome parecido) da base TOTVS.

4. **Unidade** (`src/inserir_unidade.py`)
   - Preenche `SAP5` a partir da coluna `UN` (unidade) da base TOTVS.

5. **Materiais** (`src/inserir_material.py`)
   - Lê `dados/dicionario_materiais.csv`.
   - Faz match na narrativa `SAP123` e preenche `Coluna4`.
   - Estratégia: substring (preferindo o termo mais longo) e fallback fuzzy.

6. **Normas** (`src/inserir_normas.py`)
   - Lê `dados/dicionario_normas.csv`.
   - Faz match na narrativa `SAP123` e preenche `SAP17`.

7. **Size dimension** (`src/inserir_size_dimension.py`)
   - Lê `dados/dicionario_size_dimension.csv`.
   - Faz match na narrativa `SAP123` e preenche `SAP15`.

8. **Traduções** (`src/inserir_traducoes.py`)
   - Preenche `SAP1`/`SAP2`/`SAP3`/`Coluna32` (PT/EN/ES/DE).
   - Fonte do texto para match:
     1) tenta `Descrição` do TOTVS
     2) se não houver match (descrição curta), faz fallback para o texto longo de `SAP123`

9. **Valores fixos** (`src/inserir_valores_fixos.py`)
   - Para cada item (Excel linha 3+):
     - `SAP10 = "10"`
     - `SAP14 = "NDB"`

10. **Ajuste por tamanho de narrativa** (`src/inserir_narrativas.py`)
   - Se `SAP123` tiver mais que 141 caracteres, escreve:
     - `Narrativa = "verificar internal comment"`

---

## Colunas e Regras de Preenchimento (coluna-a-coluna)

O modelo (`planilhas/planilha_padrao.xlsx`) contém estas colunas:

| Coluna | Como é preenchida | Fonte / Regra |
|---|---|---|
| `item(table) + it-codigo(field)` | Código do item | `dados/dados_teste.csv` via `src/inserir_codigos_de_itens.py` |
| `SAP10` | Valor fixo | `"10"` para linhas com código via `src/inserir_valores_fixos.py` |
| `SAP5` | Unidade | Da base TOTVS (`UN`) via `src/inserir_unidade.py` |
| `SAP14` | Valor fixo | `"NDB"` para linhas com código via `src/inserir_valores_fixos.py` |
| `SAP1` | Tradução PT | Dicionário `dados/dicionario.xlsx` via `src/inserir_traducoes.py` |
| `SAP2` | Tradução EN | Dicionário `dados/dicionario.xlsx` via `src/inserir_traducoes.py` |
| `SAP3` | Tradução ES | Dicionário `dados/dicionario.xlsx` via `src/inserir_traducoes.py` |
| `Coluna32` | Tradução DE | Dicionário `dados/dicionario.xlsx` via `src/inserir_traducoes.py` |
| `SAP6` | Product group | Da base TOTVS (`Fam Coml`) via `src/inserir_product_group.py` |
| `SAP15` | Size dimension | Match em `SAP123` usando `dados/dicionario_size_dimension.csv` |
| `Coluna4` | Material | Match em `SAP123` usando `dados/dicionario_materiais.csv` |
| `SAP17` | Norma | Match em `SAP123` usando `dados/dicionario_normas.csv` |
| `SAP123` | Internal comment (narrative) | Texto de narrativa da base TOTVS via `src/inserir_internal_comment.py` |
| `Narrativa` | Flag para revisão | Se `len(SAP123) > 141` → `"verificar internal comment"` via `src/inserir_narrativas.py` |

---

## Troubleshooting rápido

- **Erro: arquivo de trabalho inexistente**
  - Confirme se o modelo existe em `planilhas/planilha_padrao.xlsx` e o CSV existe em `dados/dados_teste.csv`.

- **Colunas esperadas não encontradas**
  - O modelo precisa conter todas as colunas listadas acima.
  - A base TOTVS precisa conter pelo menos: `Item`, `Descrição` (ou algo parecido), `UN`, `Fam Coml` e alguma coluna contendo “narrativa”.

- **Tradução não preencheu**
  - A descrição do TOTVS pode ser muito curta; o código faz fallback para `SAP123`.
  - Verifique se o dicionário tem as colunas corretas (`PORTUGUÊS`, `INGLÊS`, `ESPANHOL`, `ALEMÂO`).

---

## Estrutura do projeto

- `src/` módulos de enriquecimento e geração
- `dados/` dicionários e CSV de códigos
- `planilhas/` modelo e arquivos Excel (entrada e saída)
- `main/app.py` runner principal
