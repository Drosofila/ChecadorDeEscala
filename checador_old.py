from email import header
from sqlite3 import Row
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
from pandas import DataFrame as df
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

column=("VISITA","PRE-NATAL MANHA" , "PRE-NATAL TARDE", "PLANTAO DIURNO PRO-MATRE", "PLANTAO DIURNO SERRA", "PLANTAO NOTURNO PRO-MATRE", "PLANTAO NOTURNO SERRA")
lines=range(0, 42, 1)
'''for line in lines:
    print(line)'''


col_ref=0
col_comp=4
li_ref=3
li_comp=6

#Cellula referencia para comparacao
df = pd.read_excel("escala.xlsx", sheet_name="divisao")
nomes = df.loc[li_ref, column[col_ref]]
nomeStrRef=(nomes.split(','))

#celula a ser comparada
nomes2 = df.loc[li_comp, column[col_comp]]
nomesStrComp = (nomes2.split(','))

print('Celula ref:', nomeStrRef)
print('Celula comp:', nomesStrComp)


if any(item in nomeStrRef for item in nomesStrComp):
    print('****OPA ALGUEM ESTA TRABALHANDO DEMAIS****')
    print('Verifique a coluna', column[col_ref], 'linha', li_ref+2, 'e a coluna', column[col_comp], 'linha', li_comp+2)
else:
    print('O ESTUDANTE NÃO ESTÁ SENDO EXPLORADO')


