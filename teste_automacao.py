from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import requests
from openpyxl import load_workbook

class Root(Tk):
    def __init__(self):
        super(Root, self).__init__()
        self.title("Python Tkinter Dialog Widget")
        self.minsize(600, 300)
        # self.wm_iconbitmap('icon.ico')

        self.labelFrame = ttk.LabelFrame(self, text="Open File")
        self.labelFrame.grid(column=0, row=1, padx=20, pady=20)

        self.button()

    def button(self):
        self.button = ttk.Button(self.labelFrame, text="Browse A File", command=self.fileDialog)
        self.button.grid(column=1, row=1)

    def fileDialog(self):
        self.filename = filedialog.askopenfilename(initialdir="/", title="Select A File", filetype=
        (("xlsx files", "*.xlsx"), ("all files", "*.*")))
        self.label = ttk.Label(self.labelFrame, text="")
        self.label.grid(column=3, row=10)
        self.label.configure(text=self.filename)

        #print(self.filename)
        wb = load_workbook(self.filename)  # abrindo o Workbook test.xlsx
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
            Endereco = ws['C' + str(x)].value
            Complemento = ws['D' + str(x)].value
            form_data = {'entry.366340186': nome, 'entry.1613519622': Item, 'entry.2045809580': CEP,
                         'entry.1334556551': Endereco, 'entry.1114882091': Complemento, 'draftResponse': [],
                         'pageHistory': 0}

            user_agent = {'Referer': url,
                          'User-Agent': "Mozilla/5.0 (X11; Linux i686) AppleWebKit/537.36 (KHTML, like Gecko) "
                                        "Chrome/28.0.1500.52 Safari/537.36"}
            r = requests.post(url, data=form_data, headers=user_agent)


root = Root()
root.mainloop()
