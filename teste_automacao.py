"""
Script relativo ao teste de automação do candidato André França
Este scriot realiza o preenchimento de um formulário do Google.
Os dados estão contidos em uma planilha de excel.


Dependências:
$ pip install openpyxl
$ pip install requests
"""

from tkinter import *
from tkinter.filedialog import askopenfilename
import os
from openpyxl import load_workbook
import requests

window = Tk()
window.title('Seja bem vindo!')
window.geometry('400x200')

label2 = Label(window, text="", font=("Arial", 10))
label2.grid(column=0, row=0, pady=20)


def click1():
    global file
    file = askopenfilename(initialdir=os.getcwd(), title="Selecione o arquivo.",
                           filetype=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
    label2.configure(text=str(file))
    return file


btn1 = Button(window, text='Escolher o arquivo', command=click1)
btn1.grid(column=0, row=1, padx=150, pady=20)


def click2():
    global file
    wb = load_workbook(file)  # abrindo o Workbook test.xlsx
    ws = wb['Planilha1']  # selecionando a planilha 'Planilha1' dentro do Workbook test.xlsx

    linha = 0
    for line in ws:  # iterando em todas as linhas da 'Plan1'
        linha = linha + 1

    # Escrita no formulário online
    ### OBSERVAÇÃO ~> retire '/viewform' e add '/formResponse' ao fim do link ###

    url = "https://docs.google.com/forms/d/e/1FAIpQLSd3dOHHwAQ7JtxBc-Yyfu2B9jBT834Z6OLv_STMPm9qqYdyjg/formResponse"

    # Laço for com número de iterações = número de linhas do documento
    for x in range(1, linha + 1):
        nome = 'André José de França'
        Item = ws['A' + str(x)].value
        CEP = ws['B' + str(x)].value
        print(x)

        Endereco = ws['C' + str(x)].value
        Complemento = ws['D' + str(x)].value
        form_data = {'entry.366340186': nome, 'entry.1613519622': Item, 'entry.2045809580': CEP,
                     'entry.1334556551': Endereco, 'entry.1114882091': Complemento, 'draftResponse': [],
                     'pageHistory': 0}

        user_agent = {'Referer': url,
                      'User-Agent': "Mozilla/5.0 (X11; Linux i686) AppleWebKit/537.36 (KHTML, like Gecko) "
                                    "Chrome/28.0.1500.52 Safari/537.36"}
        r = requests.post(url, data=form_data, headers=user_agent)


btn2 = Button(window, text='Executar', command=click2, bg='green', fg='yellow')
btn2.grid(column=0, row=2, pady=10)


window.mainloop()
