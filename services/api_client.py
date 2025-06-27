import requests # Importa a biblioteca requests para fazer requisições HTTP
from config import API_TIMEOUT, URLS

# Função para buscar dados de uma api
def fetch_api(url):
    try:
        response = requests.get(url, timeout=API_TIMEOUT)
        response.raise_for_status()
        return response.json()
    except requests.RequestException as e:
        print(f"Erro na API: {e}")
        return None

# Função para obter as cotações atuais de Dólar, Euro e Bitcoin
def get_quotes():
    return {
        "Dólar": fetch_api(URLS["dolar"]),
        "Euro": fetch_api(URLS["euro"]),
        "Bitcoin": fetch_api(URLS["bitcoin"])
    }

# Função para obter os dados da taxa SELIC
def get_selic():
    return fetch_api(URLS["selic"])

# Função para obter o histórico diário da cotação do Dólar
def get_dollar_history():
    return fetch_api(URLS["historico_dolar"])
