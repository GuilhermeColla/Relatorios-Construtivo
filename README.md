## Descrição

Com esse projeto eu queria mostrar meu conhecimento de programação e de Python em meu trabalho atual. Ele extrai relatórios de um sistema de fluxo de projetos para uma posterior análise gerencial.

## Ideias
1- Realizar o dowload dos arquivos brutos do sistema usando Selenium (em andamento).

## Bugs conhecidos.

## Funcionamento
Utilizando **Selenium** fazemos o download dos relatórios em CSV brutos fornecidos pelo sistema de gestão de documentos Construtivo.
>O ideal seria utilizar a biblioteca Requests para uma solução mais ágil e simples, porém seja por falta de conhecimento, ou pela arquitetura do sistema Construtivo, não foi possível implementar esses downloads com um método GET.

Posteriormente, os novos arquivos são movidos para um diretório que ficará no mesmo nível que o ```main.py``` e utilizamos **Pandas** para processar os arquivos CSV utilizando os DataFrames do pacote. Com os dados tratados, os DataFrames são exportados para uma planilha de Excel.