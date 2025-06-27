# Projeto: Coleta de Dados de Moedas e Relatório em Excel

## Descrição
Script Python que coleta cotações de moedas, taxas de juros SELIC e histórico de dólar. Salva tudo em um Excel com múltiplas abas.

## Requisitos utilizados
- Python 3.12
- Bibliotecas:
  - requests
  - pandas
  - openpyxl


## Execução
```
# Clone este repositório
$ git clone linkrepo

# Acesse a pasta do projeto no seu terminal
$ cd projeto_moedas

# Instale dependências
$ pip install requests pandas openpyxl

# Execute o arquivo principal
$ python main.py

# A aplicação irá gerar o arquivo relatorio_moedas.xlsx
# Após isso, abra o Microsoft Excel, clique em abrir > procurar e navegue até a pasta do projeto
# Selecione o arquivo relatorio_moedas.xlsx

```

## Importante
- Antes de rodar o código novamente, feche a aba do excel
- Logs podem ser vistos após execução do código no arquivo execution.log

## Saídas geradas:
- relatorio_moedas.xlsx
- execution.log

## Fontes de API:
- AwesomeAPI
- Banco Central do Brasil
