import xlwt
import xlrd
import random
from xlutils.copy import copy

file = "C:/Users/ASUS/Desktop/Siparişler.xlsm"
rb = xlrd.open_workbook(file)
sheet = rb.sheet_by_index(0)
data = [[sheet.cell_value(r ,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
listeuzunlugu = data[4][1]
listeuzunlugu = int(listeuzunlugu)

file = "C:/Users/ASUS/Desktop/a.xls"
rb = xlrd.open_workbook(file)
sheet = rb.sheet_by_index(0)
wb = copy(rb)
w_sheet = wb.get_sheet(0)
data = [[sheet.cell_value(r ,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]


g = 0
while True:
    w_sheet.write(g, 0, )
    w_sheet.write(g, 1, )
    w_sheet.write(g, 2, )
    w_sheet.write(g, 3, )
    w_sheet.write(g, 4, )
    w_sheet.write(g, 5, )
    w_sheet.write(g, 6, )
    g += 1
    if(g == 20):
        break

liste = [3,11,23,24,28]
g = 0
while True:
    if(listeuzunlugu < 6):
        if(g < 3):
            rmahallekod = random.choice(liste)
            w_sheet.write(g, 0, rmahallekod)
        else:
            rmahallekod = random.randint(3, 28)
            w_sheet.write(g, 0, rmahallekod)
    elif(listeuzunlugu > 5 and listeuzunlugu < 10):
        if(g < 5):
            rmahallekod = random.choice(liste)
            w_sheet.write(g, 0, rmahallekod)
        else:
            rmahallekod = random.randint(3, 28)
            w_sheet.write(g, 0, rmahallekod)
    else:
        if(g < 8):
            rmahallekod = random.choice(liste)
            w_sheet.write(g, 0, rmahallekod)
        else:
            rmahallekod = random.randint(3, 28)
            w_sheet.write(g, 0, rmahallekod)

    g += 1
    if(g == listeuzunlugu ):
        break
wb.save(file)

file = "C:/Users/ASUS/Desktop/a.xls"
rb = xlrd.open_workbook(file)
wb = copy(rb)
w_sheet = wb.get_sheet(0)

g = 0
while True:
    sheet = rb.sheet_by_index(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    mahallekod = data[g] [0]
    mahallekod = int(mahallekod)
    sheet = rb.sheet_by_index(1)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    adres = data [mahallekod] [1]
    w_sheet.write(g, 1, adres)
    g += 1
    sheet = rb.sheet_by_index(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    if (g == listeuzunlugu ):
        break

g = 0
while True:
    rsiparissayisi = random.randint(1, 3)
    w_sheet.write(g, 2, rsiparissayisi)
    g += 1
    if (g == listeuzunlugu ):
        break

g = 0
while True:
    rkafekod = random.randint(3, 22)
    w_sheet.write(g, 3, rkafekod)
    g += 1
    if (g == listeuzunlugu ):
        break

wb.save(file)

file = "C:/Users/ASUS/Desktop/a.xls"
rb = xlrd.open_workbook(file)
wb = copy(rb)
w_sheet = wb.get_sheet(0)

g = 0
while True:
    sheet = rb.sheet_by_index(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    kafekod = data[g] [3]
    kafekod = int(kafekod)
    sheet = rb.sheet_by_index(1)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    adres = data [kafekod] [3]
    w_sheet.write(g, 4, adres)
    g += 1
    sheet = rb.sheet_by_index(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    if (g == listeuzunlugu ):
        break

g = 0
while True:
    kafekod = data[g] [3]
    kafekod = int(kafekod)
    sheet = rb.sheet_by_index(1)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    shazirlamasuresi = data [kafekod] [4]
    w_sheet.write(g, 5, shazirlamasuresi)
    g += 1
    sheet = rb.sheet_by_index(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    if (g == listeuzunlugu ):
        break


wb.save(file)