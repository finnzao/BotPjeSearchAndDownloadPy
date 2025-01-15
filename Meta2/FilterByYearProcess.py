import pandas as pd
import glob
import re

# Defina o caminho para a planilha original e para a pasta com as outras planilhas
caminho_planilha_original = 'relatorio-completo.xlsx'
caminho_outras_planilhas = 'processosPje/*.xlsx'

# Função para formatar o número do processo
def formatar_numero_processo(numero):
    # Remover caracteres não numéricos
    numero = re.sub(r'\D', '', str(numero))

    # Tentar extrair o ano do processo (assumindo que é '20XX')
    match_ano = re.search(r'20\d{2}', numero)
    if match_ano:
        ano = match_ano.group()
        pos_ano = match_ano.start()
    else:
        # Se não encontrar o ano, não é possível formatar corretamente
        # Retornar o número original com zeros à esquerda para completar 20 dígitos
        numero = numero.zfill(20)
        # Substituir os últimos 7 dígitos por '8050216'
        numero = numero[:-7] + '8050216'
        return numero

    # Verificar se o número começa com '0' ou '8' conforme o ano
    if int(ano) < 2015:
        prefixo = '0'
    else:
        prefixo = '8'

    # Número sequencial com dígito verificador (antes do ano)
    sequencial_com_dv = numero[:pos_ano]
    # Ajustar para 9 dígitos (7 do sequencial + 2 do dígito verificador)
    sequencial_com_dv = sequencial_com_dv.zfill(9)

    # Montar o número completo com os últimos 7 dígitos fixos em '8050216'
    numero_formatado = prefixo + sequencial_com_dv[1:] + ano + '8050216'

    # Garantir que tenha exatamente 20 dígitos
    numero_formatado = numero_formatado[:20]

    return numero_formatado

# Leia a planilha original
df_original = pd.read_excel(caminho_planilha_original)

# Verifique se as colunas necessárias existem na planilha original
coluna_processo = 'numeroProcesso'  # Nome da coluna do número do processo na planilha original
colunas_a_preencher = ['fila', 'orgaoJulgador']

for coluna in colunas_a_preencher:
    if coluna not in df_original.columns:
        df_original[coluna] = None  # Cria a coluna se não existir

# Aplicar a função de formatação ao número do processo na planilha original
df_original[coluna_processo] = df_original[coluna_processo].apply(formatar_numero_processo)

# Leia todas as outras planilhas e compile em um único DataFrame
arquivos_planilhas = glob.glob(caminho_outras_planilhas)
df_compilado = pd.DataFrame()

for arquivo in arquivos_planilhas:
    df_temp = pd.read_excel(arquivo)
    # Normalizar nomes das colunas
    df_temp.columns = df_temp.columns.str.strip()
    # Verificar se a coluna do número do processo existe
    if coluna_processo in df_temp.columns:
        # Aplicar a função de formatação ao número do processo
        df_temp[coluna_processo] = df_temp[coluna_processo].apply(formatar_numero_processo)
    else:
        # Se a coluna não existir, criar uma coluna vazia
        df_temp[coluna_processo] = None
    df_compilado = pd.concat([df_compilado, df_temp], ignore_index=True)

# Verifique se as colunas necessárias existem nas outras planilhas
colunas_necessarias = [coluna_processo] + colunas_a_preencher
for coluna in colunas_necessarias:
    if coluna not in df_compilado.columns:
        df_compilado[coluna] = None  # Cria a coluna se não existir

df_compilado = df_compilado[colunas_necessarias]

# Mescle a planilha original com o DataFrame compilado com base no número do processo
df_resultado = pd.merge(df_original, df_compilado, on=coluna_processo, how='left', suffixes=('', '_novo'))

# Atualize as colunas 'fila' e 'orgaoJulgador' com os dados das outras planilhas
for coluna in colunas_a_preencher:
    df_resultado[coluna] = df_resultado[coluna].combine_first(df_resultado[coluna + '_novo'])

# Remova as colunas extras
colunas_para_remover = [coluna + '_novo' for coluna in colunas_a_preencher]
df_resultado.drop(columns=colunas_para_remover, inplace=True)

# Salve o resultado em uma nova planilha
df_resultado.to_excel('original_atualizado.xlsx', index=False)

print("Processo concluído! A planilha 'original_atualizado.xlsx' foi criada com os dados atualizados.")
