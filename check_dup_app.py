import pandas as pd
import logging
import os
import json
from tqdm import tqdm

# Configurar o logger com codificação UTF-8
logging.basicConfig(
    filename="check_dup_app.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    encoding="utf-8"
)


def criar_diretorio(caminho_diretorio):
    try:
        # Verificar se o diretório já existe
        if not os.path.exists(caminho_diretorio):
            # Criar o diretório se não existir
            os.makedirs(caminho_diretorio)
            logging.info(f"Diretório '{caminho_diretorio}' criado com sucesso.")
        else:
            logging.info(f"O diretório '{caminho_diretorio}' já existe.")
    except Exception as e:
        logging.error(f"Erro ao criar o diretório: {str(e)}")


def ler_caminhos_do_arquivo_json(caminho_arquivo_json):
    try:
        with open(caminho_arquivo_json, "r", encoding="utf-8") as file:
            dados_json = json.load(file)
        return dados_json.get("arquivos", [])
    except FileNotFoundError:
        logging.error(
            "Arquivo JSON não encontrado. Verifique o caminho e tente novamente."
        )
        return []


def analisar_arquivo_xlsx(arquivo):
    diretorio_arquivo = arquivo.get("diretorio_arquivo", "")
    nome_arquivo = f'{diretorio_arquivo}/{arquivo.get("nome_arquivo", "")}'
    colunas_analisadas = arquivo.get("colunas", [])
    diretorio_saida = arquivo.get("diretorio_saida", "")
    nome_saida_excel = f'{diretorio_saida}/{arquivo.get("nome_saida_excel", "")}'

    criar_diretorio(diretorio_arquivo)
    criar_diretorio(diretorio_saida)

    try:
        # Log: Tentativa de carregar o arquivo
        logging.info(f"Tentativa de carregar o arquivo: {nome_arquivo}")
        df = pd.read_excel(nome_arquivo)
    except FileNotFoundError:
        # Log: Arquivo não encontrado
        logging.error(
            f"Arquivo não encontrado: {nome_arquivo}. Verifique o caminho e tente novamente."
        )
        return f"Arquivo não encontrado: {nome_arquivo}. Verifique o caminho e tente novamente."
    except Exception as e:
        # Log: Erro ao carregar o arquivo
        logging.error(f"Erro ao carregar o arquivo {nome_arquivo}: {str(e)}")
        return f"Erro ao carregar o arquivo {nome_arquivo}: {str(e)}"

    # Verificar se as colunas fornecidas pelo usuário existem no DataFrame
    colunas_invalidas = [
        coluna for coluna in colunas_analisadas if coluna not in df.columns
    ]
    if colunas_invalidas:
        # Log: Colunas inválidas
        logging.error(f"Colunas inválidas: {', '.join(colunas_invalidas)}")
        return f"Colunas inválidas: {', '.join(colunas_invalidas)}"

    # Identificar linhas duplicadas com base nas colunas fornecidas
    linhas_duplicadas = df[df.duplicated(subset=colunas_analisadas, keep=False)]

    # Exibir as linhas duplicadas encontradas
    if not linhas_duplicadas.empty:
        # Log: Linhas duplicadas encontradas
        logging.info(
            f"{len(linhas_duplicadas)} linhas duplicadas encontradas nas colunas especificadas."
        )

        # Exportar resultado para um novo arquivo XLSX
        linhas_duplicadas.to_excel(nome_saida_excel, index=False)

        logging.info(f"Resultado exportado para: {nome_saida_excel}")
        return f"Resultado exportado para: {nome_saida_excel}"
    else:
        # Log: Nenhuma linha duplicada encontrada
        logging.info("Nenhuma linha duplicada encontrada nas colunas especificadas.")
        return "Nenhuma linha duplicada encontrada nas colunas especificadas."


# Modificado para aceitar caminho do arquivo JSON
caminho_arquivo_json = "./infoArquivo.json"
arquivos = ler_caminhos_do_arquivo_json(caminho_arquivo_json)

# Exemplo de uso
for arquivo in arquivos:
    tqdm.write('Executando o check_dup_app. Por favor aguarde...')
    resultado_analise = analisar_arquivo_xlsx(arquivo)
    tqdm.write(resultado_analise)

# Garantir que todas as mensagens do logger sejam gravadas imediatamente
logging.getLogger().handlers[0].flush()
