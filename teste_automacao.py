"""
Script relativo ao teste de automação do candidato André França
Este scriot realiza o preenchimento de um formulário do Google.
Os dados estão contidos em uma planilha de excel.


Dependências:
$ pip install openpyxl
$ pip install requests
"""

import requests
from openpyxl import load_workbook

# Leitura da planilha
wb = load_workbook('test.xlsx') # abrindo o Workbook test.xlsx
ws = wb['Planilha1'] # selecionando a planilha 'Planilha1' dentro do Workbook test.xlsx
linha=0

for line in ws:  # iterando em todas as linhas da 'Plan1'
    #print(line[0]) # print a primeira célula da linha
    linha = linha + 1
    #print(linha) # print a qtdd de linhas

"""
# Utililizando em índices o nome das células como em um dicionário
a1 = ws['A1']
# Imprime o valor da célula C1
print(a1.value)
"""

# Escrita no formulário online
### OBSERVAÇÃO ~> retire '/viewform' e add '/formResponse' ao fim do link ###
url = 'https://docs.google.com/forms/d/e/1FAIpQLSd3dOHHwAQ7JtxBc-Yyfu2B9jBT834Z6OLv_STMPm9qqYdyjg/formResponse'

# Laço for com número de iterações = número de linhas do documento
for x in range(2,linha + 1):
    nome='André José de França'
    Item= ws['A'+str(x)].value
    CEP= ws['B'+str(x)].value
    Endereco= ws['C'+str(x)].value
    Complemento= ws['D'+str(x)].value
    form_data = {'entry.366340186': nome, 'entry.1613519622': Item, 'entry.2045809580': CEP, 'entry.1334556551': Endereco, 'entry.1114882091': Complemento , 'draftResponse':[], 'pageHistory':0}

    user_agent = {'Referer':url,'User-Agent': "Mozilla/5.0 (X11; Linux i686) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/28.0.1500.52 Safari/537.36"}
    r = requests.post(url, data=form_data, headers=user_agent)
