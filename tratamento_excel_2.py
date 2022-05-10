
#ESTOQUE - LINX UX DIL

import pyautogui as pg
import pyperclip as pc
import time

tic=timeit.default_timer()
pg.hotkey('win')
pg.write("estoquelinx2.xlsx")
pg.press("enter")

time.sleep(10)


pg.hotkey("ctrl","b")
time.sleep(5)
pg.hotkey("alt","f4")


origem = xl.load_workbook(r"diretorioarquivobase")
stk = origem.worksheets[0]
                         
# Abrindo o arquivo que vai pro DW

inventory_ux = xl.load_workbook(r"diretorioarquivoDW")
dz = inventory_ux.active

#Limpando o arquivo (exceto os títulos das colunas)

for row in dz['A2:W100000']:
    for cell in row:
        cell.value = None
                                 
#Definindo número de linhas e colunas

linha = stk.max_row
col = stk.max_column 
                          
# Copiando as células do arquivo do Linx UX para o do DW

for i in range (3, linha+1):
    for j in range (1, col + 1):
        c = stk.cell(row = i, column = j)
        dz.cell(row = i-1, column = j+1).value = c.value
        
for i in range(2, linha):
    c = dz.cell(row = i, column = 1)
    c.value = datetime.date.today()
    c.number_format = 'dd/mmm/yyyy'

for i in range(2, linha):
    c = dz.cell(row = i, column = 2)
    c.value = float(c.value)

for i in range(2, linha):
    c = dz.cell(row = i, column = 20)
    c.number_format = 'dd/mmm/yy'
    
for i in range(2, linha):
    c = dz.cell(row = i, column = 21)
    c.number_format = 'dd/mmm/yy'
    
for i in range(2, linha):
    c = dz.cell(row = i, column = 22)
    c.number_format = 'dd/mmm/yy'
    
for i in range(2, linha):
    c = dz.cell(row = i, column = 23)
    c.number_format = 'dd/mmm/yy'  
    
inventory_ux.save(filename = r"diretorioarquivofinal")

toc=timeit.default_timer()

tempototal=toc - tic
print(tempototal)
