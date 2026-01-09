import pandas as pd
import re


def inserir_traducoes(
    caminho_planilha_atualizada: str,
    caminho_base_totvs: str,
    caminho_dicionario_traducoes: str,
) -> None:
    """Preenche traduções (SAP1/SAP2/SAP3/Coluna32) a partir da Descrição (TOTVS).

    Regras:
    - Usa a coluna "Item" da base TOTVS para achar a "Descrição" do item.
    - Compara a descrição com o dicionário (dados/dicionario.xlsx) por substring.
    - Escreve as traduções em:
        - SAP1 = português
        - SAP2 = inglês
        - SAP3 = espanhol
        - Coluna32 = alemão
    - NÃO cria colunas novas: se alguma dessas colunas não existir na planilha, lança erro.
    """
    print("Processando traduções das descrições de produtos...")

    # Não criar colunas novas: valida que as colunas de destino existem na planilha
    colunas_destino = {
        "PORTUGUÊS": "SAP1",
        "INGLÊS": "SAP2",
        "ESPANHOL": "SAP3",
        "ALEMÂO": "Coluna32",
    }

    try:
        df_planilha = pd.read_excel(caminho_planilha_atualizada)
        df_totvs = pd.read_excel(caminho_base_totvs, header=4)
        df_dicionario = pd.read_excel(caminho_dicionario_traducoes)

        def _norm_col_name(value: object) -> str:
            return re.sub(r"\s+", "", str(value or "")).strip().upper()

        def _find_col(df: pd.DataFrame, wanted: str) -> str | None:
            wanted_n = _norm_col_name(wanted)
            for c in df.columns:
                c_n = _norm_col_name(c)
                if c_n == wanted_n:
                    return c
            # fallback: alguns arquivos vêm com sufixos tipo _X000D_
            for c in df.columns:
                c_n = _norm_col_name(c)
                if c_n.startswith(wanted_n):
                    return c
            return None

        faltando = [col for col in colunas_destino.values() if col not in df_planilha.columns]
        if faltando:
            raise ValueError(
                "A planilha atualizada não contém as colunas de tradução esperadas. "
                f"Faltando: {faltando}. "
                "Ajuste o modelo para incluir SAP1, SAP2, SAP3 e Coluna32 (não serão criadas automaticamente)."
            )

        # item -> descrição (TOTVS)
        col_item_totvs = _find_col(df_totvs, "Item")
        col_desc_totvs = _find_col(df_totvs, "Descrição")
        if col_item_totvs is None:
            raise ValueError("Coluna 'Item' não encontrada na base TOTVS.")
        if col_desc_totvs is None:
            # fallback: tenta achar alguma coluna que pareça descrição
            candidatas = [c for c in df_totvs.columns if "descr" in str(c).lower()]
            col_desc_totvs = candidatas[0] if candidatas else None
        mapa_descricoes: dict[str, str] = {}
        if col_desc_totvs is not None:
            mapa_descricoes = dict(
                zip(
                    df_totvs[col_item_totvs].astype(str).str.strip(),
                    df_totvs[col_desc_totvs].astype(str),
                )
            )

        # português -> traduções
        dicionario_traducoes: dict[str, dict[str, object]] = {}

        col_pt = _find_col(df_dicionario, "PORTUGUÊS")
        col_en = _find_col(df_dicionario, "INGLÊS")
        col_es = _find_col(df_dicionario, "ESPANHOL")
        col_de = _find_col(df_dicionario, "ALEMÂO") or _find_col(df_dicionario, "ALEMAO")
        if col_pt is None:
            raise ValueError("Coluna 'PORTUGUÊS' não encontrada no dicionário de traduções.")

        for _, row in df_dicionario.iterrows():
            palavra_pt = str(row[col_pt]).replace("\u00a0", " ").strip().lower()
            palavra_pt = re.sub(r"\s+", " ", palavra_pt)
            if not palavra_pt or palavra_pt == "nan":
                continue
            if palavra_pt in dicionario_traducoes:
                continue
            dicionario_traducoes[palavra_pt] = {
                "PORTUGUÊS": row.get(col_pt),
                "INGLÊS": row.get(col_en) if col_en is not None else None,
                "ESPANHOL": row.get(col_es) if col_es is not None else None,
                "ALEMÂO": row.get(col_de) if col_de is not None else None,
            }

        # Ordena por tamanho (mais específico primeiro)
        termos_ordenados = sorted(
            dicionario_traducoes.items(),
            key=lambda kv: len(kv[0]),
            reverse=True,
        )

        col_codigo = df_planilha.columns[0]
        col_sap123 = "SAP123" if "SAP123" in df_planilha.columns else None

        for idx in range(1, len(df_planilha)):
            codigo = str(df_planilha.loc[idx, col_codigo]).strip()
            if not codigo or codigo.lower() == "nan":
                continue

            # Fonte do match: tenta Descrição (TOTVS) e faz fallback para SAP123 (texto longo)
            candidatos_texto: list[str] = []
            descricao = mapa_descricoes.get(codigo, "") if mapa_descricoes else ""
            if descricao and str(descricao).lower() != "nan":
                candidatos_texto.append(str(descricao))
            if col_sap123 is not None:
                sap123 = df_planilha.loc[idx, col_sap123]
                if isinstance(sap123, str) and sap123.strip():
                    candidatos_texto.append(sap123)

            if not candidatos_texto:
                continue

            traducoes_encontradas = None
            for texto in candidatos_texto:
                texto_lower = re.sub(r"\s+", " ", str(texto).lower())
                for palavra_pt, traducoes in termos_ordenados:
                    if len(palavra_pt) > 5 and palavra_pt in texto_lower:
                        traducoes_encontradas = traducoes
                        break
                if traducoes_encontradas:
                    break

            if not traducoes_encontradas:
                continue

            for idioma, coluna in colunas_destino.items():
                df_planilha.at[idx, coluna] = traducoes_encontradas.get(idioma)

        # Não contar a linha descritiva (índice 0)
        contadores = {
            idioma: int(df_planilha.loc[1:, col].notna().sum())
            for idioma, col in colunas_destino.items()
        }
        df_planilha.to_excel(caminho_planilha_atualizada, index=False)

        print(f"Traduções preenchidas: {contadores}")
        print("Traduções processadas e salvas na planilha.")
    except Exception as e:
        print(f"Aviso: erro ao processar traduções: {e}")

        print(f"Aviso: erro ao processar traduções: {e}")
