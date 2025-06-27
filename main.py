from services.api_client import get_quotes, get_selic, get_dollar_history
from utils.excel_writer import save_excel
from utils.logger import log

def main():
    log("Iniciando coleta de dados...")

    quotes = get_quotes()
    selic = get_selic()
    history = get_dollar_history()

    if not quotes or not selic or not history:
        log("Erro na coleta de dados. Verifique as APIs.")
        return

    save_excel(quotes, selic, history)

    log("Relatório Excel gerado com sucesso.")

if __name__ == "__main__":
    main()
