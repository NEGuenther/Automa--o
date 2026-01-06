# Inicializar a classe com as tabelas de palavras em português e traduções
tradutor = Traducoes("palavras_portugues.xlsx", "tabela_traducoes.xlsx")

# Processar as palavras em português e adicionar as traduções
tradutor.processar_palavras()

# Salvar a tabela de palavras atualizada
tradutor.salvar_tabela("palavras_traduzidas.xlsx")