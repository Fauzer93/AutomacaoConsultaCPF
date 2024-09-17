# Consulta de Status de Pagamento de Clientes

Este projeto automatiza a consulta de status de pagamento de clientes utilizando o CPF presente em uma planilha Excel. A automação acessa o site [Consulta CPF](https://consultcpf-devaprender.netlify.app/), insere o CPF e extrai as informações de pagamento, registrando-as em uma nova planilha Excel.

## Funcionalidades

- Extrai CPFs de clientes de uma planilha `dados_clientes.xlsx`.
- Realiza a consulta automática de cada CPF no site [Consulta CPF](https://consultcpf-devaprender.netlify.app/).
- Verifica se o status do pagamento está "em dia" ou "atrasado".
- Se o pagamento estiver "em dia", o script registra a data e o método de pagamento.
- Registra os dados em uma nova planilha `planilha fechamento.xlsx` com informações sobre o status do pagamento.

## Requisitos

- Python 3.x
- Bibliotecas:
  - `openpyxl`
  - `selenium`
  - `chromedriver` (ou outro WebDriver compatível com o navegador utilizado)
  
## Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/Fauzer93/AutomacaoConsultaCPF/

----------------------------------------------------------------------------------------------------------------------------------------
Instale as bibliotecas necessárias:

Copiar código
	pip install openpyxl selenium
	Certifique-se de que o ChromeDriver está instalado e disponível no seu PATH ou coloque-o na mesma pasta do seu script.

Execução
	Certifique-se de que possui a planilha dados_clientes.xlsx no mesmo diretório do script. A planilha deve conter os seguintes campos:

		Nome
		Valor
		CPF
		Vencimento


Execute o script:
	python consulta_pagamento.py

O script abrirá o navegador, fará a consulta do CPF e salvará os resultados na planilha planilha fechamento.xlsx.

Detalhes Técnicos
	O script utiliza a biblioteca openpyxl para leitura e escrita de planilhas Excel.
	A automação web é realizada com o Selenium WebDriver, acessando o site de consulta de CPF e extraindo as informações relevantes.


