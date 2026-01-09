import pandas as pd
from thefuzz import process, fuzz
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


class Traducoes:
    def __init__(self, caminho_narrativas, caminho_traducoes, coluna_narrativa="Descrição", coluna_portugues="SAP1"):
        """
        Inicializa a classe com as tabelas de narrativas e traduções.
        :param caminho_narrativas: Caminho para o arquivo da tabela de narrativas (Excel ou CSV).
        :param caminho_traducoes: Caminho para o arquivo da tabela de traduções (Excel ou CSV).
        :param coluna_narrativa: Nome da coluna na planilha de narrativas que será usada como base
                     para comparar e gerar traduções (ex.: "narrativa" ou "SAP123").
        :param coluna_portugues: Nome da coluna em português na tabela de traduções
                     (ex.: "portugues", "Português", "PT").
        """
        self.df_narrativas = pd.read_excel(caminho_narrativas) if caminho_narrativas.endswith('.xlsx') else pd.read_csv(caminho_narrativas)
        self.df_traducoes = pd.read_excel(caminho_traducoes) if caminho_traducoes.endswith('.xlsx') else pd.read_csv(caminho_traducoes)
        self.coluna_narrativa = coluna_narrativa
        self.coluna_portugues = coluna_portugues

        if self.coluna_narrativa not in self.df_narrativas.columns:
            raise ValueError(f"A tabela de narrativas deve conter a coluna '{self.coluna_narrativa}'.")
        if self.coluna_portugues not in self.df_traducoes.columns:
            raise ValueError(f"A tabela de traduções deve conter a coluna '{self.coluna_portugues}'.")

        self.linguas = [col for col in self.df_traducoes.columns if col != self.coluna_portugues]

    def comparar_palavra(self, palavra_pt):
        """
        Compara uma palavra em português com as traduções disponíveis e retorna as melhores traduções para cada idioma.
        Tenta match exato primeiro, depois fuzzy, depois substring.
        :param palavra_pt: A palavra em português a ser comparada.
        :return: Um dicionário com as melhores traduções para cada idioma.
        """
        # garante que estamos comparando strings e removendo valores nulos
        base_portugues = (
            self.df_traducoes[self.coluna_portugues]
            .dropna()
            .astype(str)
            .tolist()
        )

        if not base_portugues:
            return {lingua: None for lingua in self.linguas}

        traducoes_por_idioma = {lingua: None for lingua in self.linguas}
        palavra_pt_str = str(palavra_pt).lower()

        # Primeiro tenta buscar por substring (palavra_pt contém termo do dicionário)
        for termo_dict in base_portugues:
            termo_dict_lower = termo_dict.lower()
            if termo_dict_lower in palavra_pt_str and len(termo_dict_lower) > 5:
                # Encontrou um match por substring
                linha_match = self.df_traducoes[
                    self.df_traducoes[self.coluna_portugues].astype(str) == termo_dict
                ]
                if not linha_match.empty:
                    linha_match = linha_match.iloc[0]
                    for lingua in self.linguas:
                        valor = linha_match.get(lingua)
                        traducoes_por_idioma[lingua] = valor if pd.notna(valor) else None
                    return traducoes_por_idioma

        # Se não achou por substring, tenta fuzzy matching
        melhor_valor, pontuacao = process.extractOne(
            palavra_pt_str,
            base_portugues,
            scorer=fuzz.ratio,
        )

        if pontuacao <= 75:
            return traducoes_por_idioma

        # pega a primeira linha que bate com o melhor valor encontrado em português
        linha_match = self.df_traducoes[
            self.df_traducoes[self.coluna_portugues].astype(str) == melhor_valor
        ]

        if linha_match.empty:
            return traducoes_por_idioma

        linha_match = linha_match.iloc[0]
        for lingua in self.linguas:
            valor = linha_match.get(lingua)
            traducoes_por_idioma[lingua] = valor if pd.notna(valor) else None

        return traducoes_por_idioma

    def processar_descricao(self):
        """
        Processa todas as narrativas e adiciona as melhores traduções para cada idioma.
        """
        for idx, narrativa in self.df_narrativas[self.coluna_narrativa].items():
            melhor_traducao = self.comparar_palavra(narrativa)
            for lingua, traducao in melhor_traducao.items():
                # agora usamos exatamente o mesmo nome de coluna do dicionário
                coluna = lingua
                if coluna not in self.df_narrativas.columns:
                    self.df_narrativas[coluna] = None
                self.df_narrativas.at[idx, coluna] = traducao

    def salvar_tabela(self, caminho_saida):
        """
        Salva a tabela de narrativas atualizada com as traduções em um arquivo.
        :param caminho_saida: Caminho para salvar a tabela (Excel ou CSV).
        """
        if caminho_saida.endswith('.xlsx'):
            self.df_narrativas.to_excel(caminho_saida, index=False)
        else:
            self.df_narrativas.to_csv(caminho_saida, index=False)
 