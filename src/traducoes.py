import pandas as pd
from thefuzz import process, fuzz

class Traducoes:
    def __init__(self, caminho_narrativas, caminho_traducoes):
        """
        Inicializa a classe com as tabelas de narrativas e traduções.
        :param caminho_narrativas: Caminho para o arquivo da tabela de narrativas (Excel ou CSV).
        :param caminho_traducoes: Caminho para o arquivo da tabela de traduções (Excel ou CSV).
        """
        self.df_narrativas = pd.read_excel(caminho_narrativas) if caminho_narrativas.endswith('.xlsx') else pd.read_csv(caminho_narrativas)
        self.df_traducoes = pd.read_excel(caminho_traducoes) if caminho_traducoes.endswith('.xlsx') else pd.read_csv(caminho_traducoes)
        
        if 'narrativa' not in self.df_narrativas.columns:
            raise ValueError("A tabela de narrativas deve conter a coluna 'narrativa'.")
        if 'portugues' not in self.df_traducoes.columns:
            raise ValueError("A tabela de traduções deve conter a coluna 'portugues'.")
        
        self.linguas = [col for col in self.df_traducoes.columns if col != 'portugues']

    def comparar_palavra(self, palavra_pt):
        """
        Compara uma palavra em português com as traduções disponíveis e retorna as melhores traduções para cada idioma.
        :param palavra_pt: A palavra em português a ser comparada.
        :return: Um dicionário com as melhores traduções para cada idioma.
        """
        traducoes_por_idioma = {}
        for lingua in self.linguas:
            traducoes = self.df_traducoes[['portugues', lingua]].dropna().to_dict('records')
            melhor_traducao, pontuacao = process.extractOne(
                palavra_pt,
                traducoes,
                scorer=fuzz.ratio,
                processor=lambda x: x['portugues']
            )
            traducoes_por_idioma[lingua] = melhor_traducao[lingua] if pontuacao > 80 else None
        return traducoes_por_idioma

    def processar_narrativas(self):
        """
        Processa todas as narrativas e adiciona as melhores traduções para cada idioma.
        """
        for idx, narrativa in self.df_narrativas['narrativa'].iteritems():
            melhor_traducao = self.comparar_palavra(narrativa)
            for lingua, traducao in melhor_traducao.items():
                coluna = f'traducao_{lingua}'
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