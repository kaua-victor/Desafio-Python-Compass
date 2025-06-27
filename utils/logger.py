import logging

# Configuração básica do sistema de logging
logging.basicConfig(
    filename='execution.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Função para registrar mensagens no log e também exibir no terminal
def log(message):
    print(message)
    logging.info(message)
