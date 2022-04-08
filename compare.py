from email import header
from sqlite3 import Row
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
from pandas import DataFrame as df
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


def comparar(i, a , b = 7, c = 7, d = 7, e = 7):
    
    column=(
        "VISITA", #C 0
        "PRE-NATAL MANHA" , #D 1 
        "PRE-NATAL TARDE", #E 2
        "PLANTAO DIURNO PRO-MATRE", #F 3 
        "PLANTAO DIURNO SERRA", #G 4
        "PLANTAO NOTURNO PRO-MATRE", #H 5
        "PLANTAO NOTURNO SERRA", #I 6
        'NULL') #J 7
    
    lines=range(0, 42, 1)

    col_ref = column[i]
    comparacao_1 = (column[a], column[b], column[c], column[d], column[e] )

    for li_ref in lines:
        df = pd.read_excel("escala.xlsx", sheet_name="divisao")
        nomes = str(df.loc[li_ref, col_ref])
        nomeStrRef=(nomes.split(','))
        #print(nomeStrRef)
        for col_comp in comparacao_1:
            nomes2 = str(df.loc[li_ref, col_comp])
            nomesStrComp = (nomes2.split(','))
            if any(item in nomeStrRef for item in nomesStrComp):
                print('****OPA ALGUEM ESTA TRABALHANDO DEMAIS****')
                print('Verifique a coluna', col_ref, 'linha', li_ref+2, 'e a coluna', col_comp, 'linha', li_ref+2, '\n')
                #print('-----------------------------------------------------------------------------------------------------\n')

def comp_tomorrow(i, a , b = 7, c = 7, d = 7, e = 7):
    
    column=(
        "VISITA", #C 0
        "PRE-NATAL MANHA" , #D 1 
        "PRE-NATAL TARDE", #E 2
        "PLANTAO DIURNO PRO-MATRE", #F 3 
        "PLANTAO DIURNO SERRA", #G 4
        "PLANTAO NOTURNO PRO-MATRE", #H 5
        "PLANTAO NOTURNO SERRA", #I 6
        'NULL') #J 7
    
    lines=range(0, 42, 1)
    lines_2=range(1, 43, 1)

    col_ref = column[i]
    comparacao_1 = (column[a], column[b], column[c], column[d], column[e] )
    for li_ref in lines:
        df = pd.read_excel("escala.xlsx", sheet_name="divisao")
        nomes = str(df.loc[li_ref, col_ref])
        nomeStrRef=(nomes.split(','))
        #print(nomeStrRef)
        for col_comp in comparacao_1:
                nomes2 = str(df.loc[li_ref+1, col_comp])
                nomesStrComp = (nomes2.split(','))
                if any(item in nomeStrRef for item in nomesStrComp):
                    print('****OPA ALGUEM ESTA TRABALHANDO DEMAIS****')
                    print('Verifique a coluna', col_ref, 'linha', li_ref+2, 'e a coluna', col_comp, 'linha', li_ref+3, '\n')