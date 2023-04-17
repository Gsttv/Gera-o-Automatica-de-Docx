# *Gerador Automático de Recibos*

Este código em Python tem como objetivo gerar recibos de forma automática a partir de informações contidas em uma planilha Excel, usando um modelo de recibo em formato Word como base.
Bibliotecas Utilizadas

    num2words: biblioteca utilizada para converter valores numéricos em palavras.
    docx: biblioteca utilizada para manipular documentos Word (.docx).
    pandas: biblioteca utilizada para trabalhar com planilhas Excel.

## Funcionamento

O código lê uma planilha Excel contendo informações de funcionários (nome e salário) e, a partir dessas informações, preenche automaticamente um modelo de recibo em formato Word, gerando um arquivo individual para cada funcionário.

## Como utilizar

    Certifique-se de ter as bibliotecas necessárias instaladas.
    Coloque o modelo de recibo em formato Word na mesma pasta que o arquivo Python.
    Coloque a planilha Excel contendo as informações dos funcionários na mesma pasta que o arquivo Python.
    No código Python, ajuste as referências no modelo de recibo para que correspondam às referências usadas no código (no exemplo, XXXX para o nome do funcionário, YYYY para o valor do salário e ZZZZ para o valor do salário por extenso).
    Execute o código Python.

