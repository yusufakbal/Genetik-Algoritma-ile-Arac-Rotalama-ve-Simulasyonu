liste1 = []
liste1mahalle = []

kafekodlistesi = []
mahallekodlistesi = []
analiste = []
kafeisim = []
analisteisim = []
mahallelistesi = []



import xlrd
file = "C:/Users/ASUS/Desktop/Siparişler.xlsm"
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_index(0)

for row in range(2 , sheet.nrows):
    analiste.append(sheet.cell_value(row , 6))
for row in range(2 , sheet.nrows):
    mahallelistesi.append(sheet.cell_value(row , 3))


data = [[sheet.cell_value(r ,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
a = data [2] [6]
a = int(a)
m = data [2] [3]
m = int(m)

file = "C:/Users/ASUS/Desktop/a.xls"
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_index(5)
data = [[sheet.cell_value(r ,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
gecensure = data [2] [a]
gecensure = int(gecensure)
liste1.append(a)
liste1mahalle.append(m)

kisit = 45
index = 0
kafekodlistesi.append(liste1)
mahallekodlistesi.append(liste1mahalle)

g = 1
h = 0
k = 0
while True:
    d = analiste[h]
    d = int(d)
    e = analiste[g]
    e = int(e)

    sheet = workbook.sheet_by_index(3)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    b = data[d][e]
    b = int(b)
    gecensure += b

    sheet = workbook.sheet_by_index(4)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    t = mahallelistesi[k]
    t = int(t)
    y = data[t][e]
    y = int(y)
    gecensure += y
    gecensure += 3


    r = mahallelistesi[h]
    r = int(r)
    w = mahallelistesi[g]
    w = int(w)
    sheet = workbook.sheet_by_index(2)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    p = data[r][w]
    p = int(p)
    gecensure += p
    gecensure += 3


    if ( d != e):
        gecensure = 0
        gecensure += y
        gecensure += 3
        k = g
        sheet = workbook.sheet_by_index(5)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        a = analiste[g]
        a = int(a)
        ik = data[2][a]
        ik = int(ik)
        gecensure += ik
        index += 1
        listea = []
        listeb = []
        kafekodlistesi.append(listea)
        mahallekodlistesi.append(listeb)


    elif (gecensure > kisit):
        gecensure = 0
        gecensure += y
        gecensure += 3
        k = g
        sheet = workbook.sheet_by_index(5)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        a = analiste[g]
        a = int(a)
        ik = data[2][a]
        ik = int(ik)
        gecensure += ik
        index += 1
        listea = []
        listeb = []
        kafekodlistesi.append(listea)
        mahallekodlistesi.append(listeb)

    kafekodlistesi [index].append(e)
    mahallekodlistesi [index].append(w)

    gecensure -= y
    gecensure -= 3
    if (g == len(analiste) - 1):
        break

    g += 1
    h += 1




rotasureleri = []


for i in range(0 , len(kafekodlistesi)):
    listerota = kafekodlistesi [i]
    rotasuresi = 0
    g = 0
    for t in range(1 , len(listerota)):
        if (len(listerota) == 1):
            break
        gidecek = listerota[g]
        gidecek = int(gidecek)
        gidilecek = listerota[t]
        gidilecek = int(gidilecek)
        sheet = workbook.sheet_by_index(3)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        aramesafe = data [gidecek] [gidilecek]
        aramesafe = int(aramesafe)
        rotasuresi += aramesafe
        g += 1
        if (t == len(listerota) - 1):
            break

    a = listerota[-1]
    a = int(a)
    listerotamahalle = mahallekodlistesi [i]
    b = listerotamahalle[0]
    b = int(b)
    sheet = workbook.sheet_by_index(4)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    kafemahalle = data [b] [a]
    kafemahalle = int(kafemahalle)
    rotasuresi += kafemahalle
    rotasuresi += 3

    g = 0
    for k in range(1 , len(listerotamahalle)):
        if (len(listerotamahalle) == 1):
            break
        gidecekmahalle = listerotamahalle[g]
        gidecekmahalle = int(gidecekmahalle)
        gidilecekmahalle = listerotamahalle[k]
        gidilecekmahalle = int(gidilecekmahalle)
        sheet = workbook.sheet_by_index(2)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        aramesafe = data [gidecekmahalle] [gidilecekmahalle]
        aramesafe = int(aramesafe)
        rotasuresi += aramesafe
        rotasuresi += 3
        g += 1
        if (k == len(listerotamahalle) - 1):
            break

    rotasureleri.append(rotasuresi)


file = "C:/Users/ASUS/Desktop/a.xls"
rb = xlrd.open_workbook(file)
sheet = rb.sheet_by_index(1)
data = [[sheet.cell_value(r ,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

indeks = 0
for i in range(0, len(kafekodlistesi)):
    ornekliste = kafekodlistesi[i]
    lista =[]
    kafeisim.append(lista)
    for r in ornekliste:
        kafeismi = data[r] [3]
        kafeisim[indeks].append(kafeismi)
    indeks += 1

indeks = 0
for i in range(0, len(mahallekodlistesi)):
    ornekliste = mahallekodlistesi[i]
    lista =[]
    analisteisim.append(lista)
    for r in ornekliste:
        mahalleismi = data[r] [1]
        analisteisim[indeks].append(mahalleismi)
    indeks += 1



j = 1
for i in range(0 , len(kafekodlistesi)):
    print(j,".Rota: {} Müşteriler: {} Rotasüresi: {}".format(kafeisim [i],analisteisim [i],rotasureleri [i]))
    j += 1

rotakm = 0
for i in range(0 , len(kafekodlistesi)):
    listerota = kafekodlistesi [i]
    g = 0
    for t in range(1 , len(listerota)):
        if (len(listerota) == 1):
            break
        gidecek = listerota[g]
        gidecek = int(gidecek)
        gidilecek = listerota[t]
        gidilecek = int(gidilecek)
        sheet = workbook.sheet_by_index(7)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        aramesafe = data [gidecek] [gidilecek]
        aramesafe = int(aramesafe)
        rotakm += aramesafe
        g += 1
        if (t == len(listerota) - 1):
            break

    a = listerota[-1]
    a = int(a)
    listerotamahalle = mahallekodlistesi [i]
    b = listerotamahalle[0]
    b = int(b)
    sheet = workbook.sheet_by_index(9)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    kafemahalle = data [b] [a]
    kafemahalle = int(kafemahalle)
    rotakm += kafemahalle

    g = 0
    for k in range(1 , len(listerotamahalle)):
        if (len(listerotamahalle) == 1):
            break
        gidecekmahalle = listerotamahalle[g]
        gidecekmahalle = int(gidecekmahalle)
        gidilecekmahalle = listerotamahalle[k]
        gidilecekmahalle = int(gidilecekmahalle)
        sheet = workbook.sheet_by_index(8)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        aramesafe = data [gidecekmahalle] [gidilecekmahalle]
        aramesafe = int(aramesafe)
        rotakm += aramesafe
        g += 1
        if (k == len(listerotamahalle) - 1):
            break

yakit_masrafi = rotakm * 0.1668
print("Toplam rota km:",rotakm)
print("Toplam yakıt masrafı:",yakit_masrafi)
