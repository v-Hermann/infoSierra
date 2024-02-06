import glob
import pandas as pd
import os

# Definir o caminho para os arquivos TXT
caminho_arquivos = 'C:/Users/seuUsuario/Downloads/Files/*.txt'

# Lista para armazenar os dados de cada arquivo
dados = []

# Lista de todos os arquivos no diretório
todos_arquivos = glob.glob(caminho_arquivos)

# Lista para armazenar os nomes dos arquivos processados com sucesso
arquivos_processados = []

# Lista de possíveis codificações de idioma
codificacoes = ['utf-8', 'utf-16-le', 'latin-1']


# Função para tentar ler o arquivo com diferentes codificações
def ler_arquivo_com_codificacoes(arquivo, codificacoes):
    for cod in codificacoes:
        try:
            with open(arquivo, 'r', encoding=cod) as f:
                return f.readlines(), cod
        except UnicodeDecodeError:
            continue
    return None, None


# Processar cada arquivo TXT no diretório
for arquivo in todos_arquivos:
    linhas, cod_usada = ler_arquivo_com_codificacoes(arquivo, codificacoes)
    if linhas is None:
        print(f"Não foi possível ler o arquivo {arquivo} com as codificações fornecidas.")
        continue

    # Dicionário para armazenar os dados do arquivo atual
    dados_arquivo = {'nome do arquivo': os.path.basename(arquivo)}  # Adiciona o nome do arquivo no dicionário
    for linha in linhas:
        linha_limpa = linha.strip()
        if ':' in linha_limpa:
            chave, valor = linha_limpa.split(':', 1)
            dados_arquivo[chave.strip()] = valor.strip()
    if dados_arquivo:  # Certificar-se de que o dicionário não está vazio além do nome do arquivo
        dados.append(dados_arquivo)
        arquivos_processados.append(os.path.basename(arquivo))

# Criar um DataFrame com os dados coletados
df = pd.DataFrame(dados)

# Especificar o nome do arquivo Excel onde os dados serão salvos
nome_planilha = 'dados_importados.xlsx'

# Salvar o DataFrame em uma planilha Excel
df.to_excel(nome_planilha, index=False)

print('Dados importados com sucesso para o Excel.')

# Determinar quais arquivos não foram processados
arquivos_nao_processados = [os.path.basename(arquivo) for arquivo in todos_arquivos if
                            os.path.basename(arquivo) not in arquivos_processados]

print("Arquivos não processados:")
for arquivo in arquivos_nao_processados:
    print(arquivo)
