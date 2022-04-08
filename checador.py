from email import header
from sqlite3 import Row
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import numpy as np
from pandas import DataFrame as df
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import compare

compare.comparar(0, 1, 3, 4) #C - F, G, D

compare.comparar(1, 3, 4) #D - F, G

compare.comparar(2, 3, 4) #E - F, G

compare.comparar(3, 5, 6) #F - (G), H, I, (C), (D), (E)

compare.comparar(4, 5, 6) #G - (F), H, I, (C), (D), (E)

compare.comparar(5, 6) #H - I, (F), (G)

compare.comp_tomorrow(5, 0, 1, 3, 4) #H - F+1, G+1, C+1, D+1

compare.comp_tomorrow(6, 0, 1, 3, 4) #H - F+1, G+1, C+1, D+1



print('---------------------------------------------------------------')
print('Pronto! Se algum problema foi encontrado, sua respectiva linha\ne coluna estará listada acima. Corrija o erro, salve a planilha\ne rode o script novamente. Caso o script não tenha retornado \nnenhum erro, todos estão escalados dentro da regulamentação.')
print('---------------------------------------------------------------\n')