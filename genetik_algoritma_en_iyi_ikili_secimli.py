kafelist = []
mahallelist = []
kromozom = []
fkromozom = []

import xlwt
from xlutils.copy import copy
import random
import xlrd
file = "C:/Users/ASUS/Desktop/Siparişler.xlsm"
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_index(0)
data = [[sheet.cell_value(r ,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]


listeuzunlugu = data[4][1]
listeuzunlugu = int(listeuzunlugu)

populasyonbuyuklugu = 1

for i in range(2, listeuzunlugu + 1):
    populasyonbuyuklugu *= i

populasyonbuyuklugu *= populasyonbuyuklugu


for row in range(2 , sheet.nrows):
    kafelist.append(sheet.cell_value(row , 6))

for row in range(2 , sheet.nrows):
    mahallelist.append(sheet.cell_value(row , 3))


for i in kafelist:
    kromozom.append(i)

for i in mahallelist:
    kromozom.append(i)

fkromozom.append(kromozom)

file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
rb = xlrd.open_workbook(file)
wb = copy(rb)
w_sheet = wb.get_sheet(0)

h = 0
while True:
    index1 = random.randint(0, listeuzunlugu - 1)
    index2 = random.randint(0, listeuzunlugu - 1)
    a = kromozom[index1]
    b = kromozom[index2]
    kromozom[index1] = b
    kromozom[index2] = a

    index3 = random.randint(listeuzunlugu, listeuzunlugu + listeuzunlugu -1)
    index4 = random.randint(listeuzunlugu, listeuzunlugu + listeuzunlugu -1)
    c = kromozom[index3]
    d = kromozom[index4]
    kromozom[index3] = d
    kromozom[index4] = c
    g = 0
    for i in kromozom:
        w_sheet.write(h, g, i)
        g += 1

    h += 1
    if(h == populasyonbuyuklugu):
        break
    if (h == 100):
        break

wb.save(file)



g = 0
c = 0
n = 10000
while True:
    analiste = []
    file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

    for row in range(0, sheet.ncols):
        analiste.append(sheet.cell_value(g, row))

    kafe1 = analiste[0]
    kafe1 = int(kafe1)

    file = "C:/Users/ASUS/Desktop/a.xls"
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(5)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

    ist_kafe = data[2][kafe1]
    ist_kafe = int(ist_kafe)
    gecensure = ist_kafe

    sheet = workbook.sheet_by_index(3)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    h = 1
    for i in analiste:
        i = int(i)
        gkafe = analiste[h]
        gkafe = int(gkafe)
        kafe_kafe = data[i][gkafe]
        kafe_kafe = int(kafe_kafe)
        gecensure += kafe_kafe
        if (h == listeuzunlugu - 1):
            break
        h += 1

    sheet = workbook.sheet_by_index(4)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

    kaf_mah = data[listeuzunlugu - 1][listeuzunlugu]
    kaf_mah = int(kaf_mah)
    gecensure += kaf_mah

    sheet = workbook.sheet_by_index(2)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    h = listeuzunlugu
    a = listeuzunlugu + 1
    while True:
        glmah = analiste[h]
        glmah = int(glmah)
        gmah = analiste[a]
        gmah = int(gmah)
        mah_mah = data[glmah][gmah]
        mah_mah = int(mah_mah)
        gecensure += mah_mah
        if (a == (listeuzunlugu * 2) - 1):
            break
        h += 1
        a += 1

    file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
    rb = xlrd.open_workbook(file)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    w_sheet.write(g, (listeuzunlugu * 2) + 2, gecensure)

    wb.save(file)

    a = gecensure
    if (a < n):
        n = a
        c = g

    if(g == populasyonbuyuklugu - 1 or g == 99):
        break
    g += 1

    print(gecensure)


file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_index(0)
data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
rb = xlrd.open_workbook(file)
wb = copy(rb)
w_sheet = wb.get_sheet(0)

enkucuklist = []
for row in range(0, listeuzunlugu * 2):
    enkucuklist.append(sheet.cell_value(c, row))

sira = (listeuzunlugu * 2) + 5
for i in enkucuklist:
    w_sheet.write(0, sira, i)
    sira += 1
w_sheet.write(0 , sira + 2, n)

wb.save(file)



son = 0
ei = 1
while True:
    dongu = 0
    h = 0
    listindex = []
    while True:
        g = 0
        b = 0
        m = 10000
        while True:
            file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
            workbook = xlrd.open_workbook(file)
            sheet = workbook.sheet_by_index(0)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

            a = data[g][12]
            a = int(a)

            for i in listindex:
                if(g == i):
                    a = 100000


            if (a < m):
                m = a
                b = g



            if(g == 99 or g == populasyonbuyuklugu - 1):
                break

            g += 1

        print("değer: {}  indis: {}".format(m, b))

        listindex.append(b)

        g = 0
        c = 0
        n = 10000
        while True:
            a = data[g] [12]
            a = int(a)

            for i in listindex:
                if(g == i):
                    a = 100000

            if (a < n):
                n = a
                c = g


            if(g == 99 or g == populasyonbuyuklugu - 1):
                break

            g += 1

        print("değer: {}  indis: {}".format(n, c))

        listindex.append(c)


        liste1 = []
        liste2 = []
        for row in range(0 , listeuzunlugu*2):
            liste1.append(sheet.cell_value(b , row))
        for row in range(0 , listeuzunlugu*2):
            liste2.append(sheet.cell_value(c , row))

        file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
        rb = xlrd.open_workbook(file)
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)


        a = (listeuzunlugu / 2) + 1
        a = int(a)
        sira = 0
        for i in range(0, a):
            deger = liste1[i]
            deger2 = liste2[i]
            w_sheet.write(b, sira, deger)
            w_sheet.write(c, sira, deger2)
            sira += 1

        a = (listeuzunlugu / 2) + 1
        a = int(a)
        sira = a
        for i in range(a, listeuzunlugu ):
            deger = liste2[i]
            deger2 = liste1[i]
            w_sheet.write(b, sira, deger)
            w_sheet.write(c, sira, deger2)
            sira += 1

        a = listeuzunlugu +  (listeuzunlugu / 2) + 1
        a = int(a)
        sira = listeuzunlugu
        for i in range(listeuzunlugu, a):
            deger = liste1[i]
            deger2 = liste2[i]
            w_sheet.write(b, sira, deger)
            w_sheet.write(c, sira, deger2)
            sira += 1

        a= listeuzunlugu +  (listeuzunlugu / 2) + 1
        a = int(a)
        sira = a
        for i in range(a , listeuzunlugu*2):
            deger = liste2[i]
            deger2 = liste1[i]
            w_sheet.write(b, sira, deger)
            w_sheet.write(c, sira, deger2)
            sira += 1

        wb.save(file)

        file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheet_by_index(0)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        rb = xlrd.open_workbook(file)
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)


        r = b
        dur = h + 2
        while True:
            sira = 0
            degiseceklistkafe = []
            degiseceklistmah = []
            olmayankafe = []
            olmayanmah = []
            p = 0
            for row in range(0, listeuzunlugu):
                degiseceklistkafe.append(sheet.cell_value(r, row))

            for row in range(listeuzunlugu, listeuzunlugu*2):
                degiseceklistmah.append(sheet.cell_value(r, row))

            for i in range(0, listeuzunlugu):
                p = 0
                deger = liste1[i]
                for j in degiseceklistkafe:
                    if (deger == j):
                        p += 1
                if(p == 0):
                    olmayankafe.append(deger)

            for i in range(listeuzunlugu, listeuzunlugu*2):
                p = 0
                deger = liste1[i]
                for j in degiseceklistmah:
                    if (deger == j):
                        p += 1
                if(p == 0):
                    olmayanmah.append(deger)



            t = 0
            sira = 0
            while True:
                k = 0
                degisken = data[r] [sira]
                degisken = int(degisken)
                for i in range(sira + 1, listeuzunlugu):
                    degisken2 = data[r] [i]
                    degisken2 = int(degisken2)
                    if(degisken == degisken2):
                        k += 1
                if(k >= 1):
                    yazilacak = olmayankafe[t]
                    yazilacak = int(yazilacak)
                    w_sheet.write(r, sira, yazilacak)
                    t += 1

                if(sira == listeuzunlugu - 2):
                    break

                sira += 1

            print(olmayankafe, olmayanmah)

            t = 0
            sira = listeuzunlugu
            while True:
                k = 0
                degisken = data[r] [sira]
                degisken = int(degisken)
                for i in range(sira + 1, listeuzunlugu*2):
                    degisken2 = data[r] [i]
                    degisken2 = int(degisken2)
                    if(degisken == degisken2):
                        k += 1
                if(k >= 1):
                    yazilacak = olmayanmah[t]
                    yazilacak = int(yazilacak)
                    w_sheet.write(r, sira, yazilacak)
                    t += 1

                if(sira == listeuzunlugu*2 - 2):
                    break

                sira += 1

            h += 1
            if(h == dur):
                break

            r = c


        wb.save(file)
        if(dongu == 98):
            break

        dongu += 2


    file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    rb = xlrd.open_workbook(file)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)


    h = 0
    while True:
        index1 = random.randint(0, listeuzunlugu - 1)
        index2 = random.randint(0, listeuzunlugu - 1)
        a = data[h][index1]
        b = data[h][index2]
        w_sheet.write(h, index1, b)
        w_sheet.write(h, index2, a)

        index3 = random.randint(listeuzunlugu, listeuzunlugu*2 - 1)
        index4 = random.randint(listeuzunlugu, listeuzunlugu*2 - 1)
        c = data[h][index3]
        d = data[h][index4]
        w_sheet.write(h, index3, d)
        w_sheet.write(h, index4, c)

        h += 1
        if(h == populasyonbuyuklugu):
            break
        if (h == 100):
            break

    wb.save(file)


    g = 0
    while True:
        analiste = []
        file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheet_by_index(0)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

        for row in range(0, sheet.ncols):
            analiste.append(sheet.cell_value(g, row))

        kafe1 = analiste[0]
        kafe1 = int(kafe1)

        file = "C:/Users/ASUS/Desktop/a.xls"
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheet_by_index(5)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

        ist_kafe = data[2][kafe1]
        ist_kafe = int(ist_kafe)
        gecensure = ist_kafe

        sheet = workbook.sheet_by_index(3)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        h = 1
        for i in analiste:
            i = int(i)
            gkafe = analiste[h]
            gkafe = int(gkafe)
            kafe_kafe = data[i][gkafe]
            kafe_kafe = int(kafe_kafe)
            gecensure += kafe_kafe
            if (h == listeuzunlugu - 1):
                break
            h += 1

        sheet = workbook.sheet_by_index(4)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

        kaf_mah = data[listeuzunlugu - 1][listeuzunlugu]
        kaf_mah = int(kaf_mah)
        gecensure += kaf_mah

        sheet = workbook.sheet_by_index(2)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        h = listeuzunlugu
        a = listeuzunlugu + 1
        while True:
            glmah = analiste[h]
            glmah = int(glmah)
            gmah = analiste[a]
            gmah = int(gmah)
            mah_mah = data[glmah][gmah]
            mah_mah = int(mah_mah)
            gecensure += mah_mah
            if (a == (listeuzunlugu * 2) - 1):
                break
            h += 1
            a += 1

        file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
        rb = xlrd.open_workbook(file)
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)
        w_sheet.write(g, (listeuzunlugu * 2) + 2, gecensure)

        wb.save(file)

        a = gecensure
        if (a < n):
            n = a
            c = g

        if(g == populasyonbuyuklugu - 1 or g == 99):
            break
        g += 1

    file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    rb = xlrd.open_workbook(file)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)

    enkucuklist = []
    for row in range(0, listeuzunlugu * 2):
        enkucuklist.append(sheet.cell_value(c, row))

    sira = (listeuzunlugu * 2) + 5
    for i in enkucuklist:
        w_sheet.write(ei, sira, i)
        sira += 1
    w_sheet.write(ei, sira + 2, n)

    wb.save(file)


    if(son == 10):
        break

    ei += 1
    son += 1