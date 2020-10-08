mahallerotasi = []
listenumarasi = []
analiste = []
analisteisim = []
kafeisim = []
sureler = []
kaferotasi = []

import xlrd
file = "C:/Users/ASUS/Desktop/Siparişler.xlsm"
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_index(0)
data = [[sheet.cell_value(r ,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

for row in range(2 , sheet.nrows):
    analiste.append(sheet.cell_value(row , 3))

g = 2
enkisayol = 10000
index = 0
listesuresi = 0

while True:
    baslangicmahalle = data[g][3]
    baslangicmahalle = int(baslangicmahalle)
    listesirasi1 = data[g][2]
    listesirasi1 = int(listesirasi1)
    listem = []
    listel = []
    mahallerotasi.append(listem)
    mahallerotasi[index].append(baslangicmahalle)
    listenumarasi.append(listel)
    listenumarasi[index].append(listesirasi1)
    a = 2
    while True:
        h = 2
        while True:
            digermahalle = data[h] [3]
            digermahalle = int(digermahalle)
            listesirasi2 = data[h] [2]
            listesirasi2 = int(listesirasi2)
            kayitlimahalleler = mahallerotasi[index]
            kayitlilistesirasi = listenumarasi[index]
            file = "C:/Users/ASUS/Desktop/a.xls"
            workbook = xlrd.open_workbook(file)
            sheet = workbook.sheet_by_index(2)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
            i = 0
            while True:
                if(listesirasi2 == kayitlilistesirasi[i]):
                    aralarindakiuzaklik = 1000
                    break
                else:
                    aralarindakiuzaklik = data[baslangicmahalle][digermahalle]
                    aralarindakiuzaklik = int(aralarindakiuzaklik)

                if(i == (len(kayitlilistesirasi) - 1)):
                    break
                i += 1


            if(aralarindakiuzaklik < enkisayol):
                enkisayol = aralarindakiuzaklik
                digermahallekodu = digermahalle
                digerlistesirasi = listesirasi2

            file = "C:/Users/ASUS/Desktop/Siparişler.xlsm"
            workbook = xlrd.open_workbook(file)
            sheet = workbook.sheet_by_index(0)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

            if (h == len(analiste) + 1):
                break
            h += 1

        listesuresi += enkisayol

        mahallerotasi[index].append(digermahallekodu)
        listenumarasi[index].append(digerlistesirasi)
        baslangicmahalle = digermahallekodu
        listesirasi1 = digerlistesirasi
        enkisayol = 10000


        if (a == len(analiste)):
            break
        a += 1

    sureler.append(listesuresi)
    listesuresi = 0

    if (g == len(analiste) + 1):
        break
    g += 1
    index += 1

en_kisa = 1000
be = 0
for i in sureler:
    sy = i
    if (sy<en_kisa ):
        en_kisa = sy
        en_kisa_mahalle_rotasi = mahallerotasi [be]
        rota_liste_numarasi = listenumarasi [be]
    be += 1

for i in rota_liste_numarasi:
    kafe = data[i + 1] [6]
    kafe = int(kafe)
    kaferotasi.append(kafe)


print("En kısa rotanın liste numaraları:",rota_liste_numarasi)
print("En kısa mahalle rotası:",en_kisa_mahalle_rotasi)
print("Liste numaralarına göre kafe rotası:",kaferotasi)
print("En kısa toplam mesafe:",en_kisa)

liste1 = []
liste1mahalle = []

kafekodlistesi = []
mahallekodlistesi = []

import xlrd
file = "C:/Users/ASUS/Desktop/a.xls"
workbook = xlrd.open_workbook(file)
a = kaferotasi [0]
a = int(a)
m = en_kisa_mahalle_rotasi[0]
m = int(m)

sheet = workbook.sheet_by_index(5)
data = [[sheet.cell_value(r ,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
gecensure = data [2] [a]
gecensure = int(gecensure)
liste1.append(a)
liste1mahalle.append(m)

kafekodlistesi.append(liste1)
mahallekodlistesi.append(liste1mahalle)

kisit = 45
index = 0

g = 1
h = 0
k = 0
while True:
    d = kaferotasi[h]
    d = int(d)
    e = kaferotasi[g]
    e = int(e)

    sheet = workbook.sheet_by_index(3)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    b = data[d][e]
    b = int(b)
    gecensure += b

    sheet = workbook.sheet_by_index(4)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    t = en_kisa_mahalle_rotasi[k]
    t = int(t)
    y = data[t][e]
    y = int(y)
    gecensure += y
    gecensure += 3



    r = en_kisa_mahalle_rotasi[h]
    r = int(r)
    w = en_kisa_mahalle_rotasi[g]
    w = int(w)
    sheet = workbook.sheet_by_index(2)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    p = data[r][w]
    p = int(p)
    gecensure += p
    gecensure += 3

    if (gecensure > kisit):
        gecensure = 0
        gecensure += y
        gecensure += 3
        k = g
        sheet = workbook.sheet_by_index(5)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        a = kaferotasi[g]
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
    if (g == len(kaferotasi) - 1):
        break

    g += 1
    h += 1


rotasureleri = []


for i in range(0 , len(kafekodlistesi)):
    listerota = kafekodlistesi [i]
    bk = kafekodlistesi[i] [0]
    bk = int(bk)
    sheet = workbook.sheet_by_index(5)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    rotasuresi = data[2][bk]
    rotasuresi = int(rotasuresi)
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


print(kafekodlistesi)
rotakm = 0
for i in range(0 , len(kafekodlistesi)):
    listerota = kafekodlistesi [i]
    bk = kafekodlistesi[i] [0]
    bk = int(bk)
    sheet = workbook.sheet_by_index(6)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    l = data[2][bk]
    l = int(l)
    rotakm += l
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
