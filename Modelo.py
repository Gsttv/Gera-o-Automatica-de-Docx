from num2words import num2words
from docx import Document
import pandas as pd

tabela = pd.read_excel("Informacoes.xlsx")


for linha in tabela.index:

    documento = Document("ReciboModelo.docx") # Ele voltar para o recibo modelo cada vez que reinicia o loop
                                              # porque senão ele vai apenas executar uma vez e ficar reutilizadno
                                              # o modelo ja editado
    nome = tabela.loc[linha,"FUNCIONARIO"]
    salario = str(tabela.loc[linha,"SALARIO"])
    salarioEscrito =num2words(str(salario), lang='pt-br')

    referencias = {
        "XXXX":nome,
        "YYYY": salario,
        "ZZZZ": salarioEscrito
    }
    for paragrafo in documento.paragraphs:
        for chave in referencias:
            print(chave)
            valorDaChave = referencias[chave]
            print(valorDaChave)
            paragrafo.text = paragrafo.text.replace(chave,valorDaChave)

    documento.save(f"Recibo - {nome}.docx")
