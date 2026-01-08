import pandas as pd
from thefuzz import process, fuzz

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

        melhor_valor, pontuacao = process.extractOne(
            str(palavra_pt),
            base_portugues,
            scorer=fuzz.ratio,
        )

        traducoes_por_idioma = {lingua: None for lingua in self.linguas}

        if pontuacao <= 80:
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
 