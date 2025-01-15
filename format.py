import pandas as pd
import re

# Carregar o arquivo Excel
df = pd.read_excel('dados_partes.xlsx')  # Substitua 'seu_arquivo.xlsx' pelo nome do arquivo

# Função para extrair apenas o nome da parte
def extrair_nome(texto):
    # Usar expressão regular para capturar apenas o nome no início da string
    nome = re.match(r'^[A-Z\s]+', texto)
    return nome.group(0).strip() if nome else texto

# Aplicar a função na coluna 'nome da parte'
df['Nome da Parte'] = df['Nome da Parte'].apply(extrair_nome)

# Salvar o resultado em um novo arquivo Excel
df.to_excel('nome_formatado.xlsx', index=False)
print("Nomes extraídos e salvos em 'nome_formatado.xlsx'")
