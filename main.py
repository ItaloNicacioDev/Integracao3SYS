#Integração de sistemas - BETA VER. 0.04
#Ultima atualização : 21/04/2024
# © 2024. Todos os direitos reservados. Este código é propriedade intelectual de Italo Nicacio Eufrasio.
#-----------------------------------------------------------------------------------------------------------------------
#Bibliotecas

import openpyxl
import requests
import time

# Função para ler os dados da tabela do Excel
def ler_tabela_excel(nome_arquivo, nome_planilha):
    try:
        workbook = openpyxl.load_workbook(nome_arquivo)
        sheet = workbook[nome_planilha]

        # Ler os dados da tabela do Excel
        dados = []
        for row in sheet.iter_rows(values_only=True):
            dados.append(row)

        return dados
    except Exception as e:
        print(f"Erro ao ler a tabela do Excel: {e}")
        return None

# Função para fazer uma requisição GET para uma API
def fazer_requisicao_get(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Erro ao fazer requisição GET para {url}. Código de status: {response.status_code}")
            return None
    except Exception as e:
        print(f"Erro ao fazer requisição GET para {url}: {e}")
        return None

# Função para fazer uma requisição POST para uma API
def fazer_requisicao_post(url, dados):
    try:
        response = requests.post(url, json=dados)
        if response.status_code == 200:
            print("Dados atualizados com sucesso!")
        else:
            print(f"Erro ao fazer requisição POST para {url}. Código de status: {response.status_code}")
    except Exception as e:
        print(f"Erro ao fazer requisição POST para {url}: {e}")

# Função para integrar os sistemas
def integrar_sistemas():
    # Nome do arquivo Excel e da planilha
    nome_arquivo_excel = "dados.xlsx"
    nome_planilha_excel = "Planilha1"

    # URL da API para atualização da tabela de Excel
    url_api_excel = "https://api.example.com/atualizar-excel"

    # URL do sistema externo
    url_sistema_externo = "https://api.example.com/dados"

    while True:
        # Ler dados da tabela do Excel
        dados_excel = ler_tabela_excel(nome_arquivo_excel, nome_planilha_excel)
        if dados_excel:
            print("Dados da tabela do Excel:")
            print(dados_excel)

        # Fazer requisição para o sistema externo
        dados_sistema_externo = fazer_requisicao_get(url_sistema_externo)
        if dados_sistema_externo:
            print("Dados do sistema externo:")
            print(dados_sistema_externo)

            # Atualizar os dados da tabela do Excel com os dados recebidos do sistema externo
            fazer_requisicao_post(url_api_excel, dados_sistema_externo)

        # Aguardar 1 minuto antes de fazer a próxima verificação
        time.sleep(60)

if __name__ == "__main__":
    integrar_sistemas()






