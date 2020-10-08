import xlwt
import xlrd
import random
from xlutils.copy import copy

listeuzunlugu = 4

file = "C:/Users/ASUS/Desktop/Kurye.xls"
rb = xlrd.open_workbook(file)
wb = copy(rb)
w_sheet = wb.get_sheet(0)

g = 0
while True:
    w_sheet.write(g, 0, )
    w_sheet.write(g, 1, )
    w_sheet.write(g, 2, )
    w_sheet.write(g, 3, )
    g += 1
    if(g == 200):
        break
wb.save(file)


bt = 0
L = 1
sira = 1
issuresi = 1
kuryenumarasi = 2
toplamyakitmasrafi = 0
toplam_yakit_masrafi = 0

while True:
    file = "C:/Users/ASUS/Desktop/a.xls"
    rb = xlrd.open_workbook(file)
    sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

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

    g = 0
    i = 1
    while True:
        w_sheet.write(g, 0, i)
        g += 1
        i += 1
        if (g == listeuzunlugu):
            break
        wb.save(file)

    liste = [3, 11, 23, 24, 28]
    g = 0
    while True:
        if (listeuzunlugu < 6):
            if (g < 3):
                rmahallekod = random.choice(liste)
                w_sheet.write(g, 1, rmahallekod)
            else:
                rmahallekod = random.randint(3, 28)
                w_sheet.write(g, 1, rmahallekod)
        elif (listeuzunlugu > 5 and listeuzunlugu < 10):
            if (g < 5):
                rmahallekod = random.choice(liste)
                w_sheet.write(g, 1, rmahallekod)
            else:
                rmahallekod = random.randint(3, 28)
                w_sheet.write(g, 1, rmahallekod)
        else:
            if (g < 8):
                rmahallekod = random.choice(liste)
                w_sheet.write(g, 1, rmahallekod)
            else:
                rmahallekod = random.randint(3, 28)
                w_sheet.write(g, 1, rmahallekod)
        g += 1
        if (g == listeuzunlugu):
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
        mahallekod = data[g][1]
        mahallekod = int(mahallekod)
        sheet = rb.sheet_by_index(1)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        adres = data[mahallekod][1]
        w_sheet.write(g, 2, adres)
        g += 1
        sheet = rb.sheet_by_index(0)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        if (g == listeuzunlugu):
            break

    g = 0
    while True:
        rsiparissayisi = random.randint(1, 3)
        w_sheet.write(g, 3, rsiparissayisi)
        g += 1
        if (g == listeuzunlugu):
            break

    g = 0
    while True:
        rkafekod = random.randint(3, 22)
        w_sheet.write(g, 4, rkafekod)
        g += 1
        if (g == listeuzunlugu):
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
        kafekod = data[g][4]
        kafekod = int(kafekod)
        sheet = rb.sheet_by_index(1)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        adres = data[kafekod][3]
        w_sheet.write(g, 5, adres)
        g += 1
        sheet = rb.sheet_by_index(0)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        if (g == listeuzunlugu):
            break

    g = 0
    while True:
        kafekod = data[g][4]
        kafekod = int(kafekod)
        sheet = rb.sheet_by_index(1)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        shazirlamasuresi = data[kafekod][4]
        w_sheet.write(g, 6, shazirlamasuresi)
        g += 1
        sheet = rb.sheet_by_index(0)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        if (g == listeuzunlugu):
            break

    wb.save(file)



    kafelist = []
    mahallelist = []
    kromozom = []
    fkromozom = []

    kuzunluk = 30
    r_uzunluk = 1

    import xlwt
    from xlutils.copy import copy
    import random
    import xlrd

    file = "C:/Users/ASUS/Desktop/a.xls"
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]



    for row in range(0, sheet.nrows):
        kafelist.append(sheet.cell_value(row, 4))

    for row in range(0, sheet.nrows):
        mahallelist.append(sheet.cell_value(row, 1))

    for i in kafelist:
        kromozom.append(i)

    for i in mahallelist:
        kromozom.append(i)

    file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
    workbook = xlrd.open_workbook(file)
    sheet = workbook.sheet_by_index(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
    rb = xlrd.open_workbook(file)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)

    row = 0
    while True:
        col = 0
        while True:
            w_sheet.write(row, col, )
            if (col == 54):
                break
            col += 1
        if (row == kuzunluk + 10):
            break
        row += 1
    wb.save(file)

    sayilar = list(range(listeuzunlugu))
    rotalistesi = []
    surelistesi = []
    list_ind = 0
    while True:
        file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheet_by_index(0)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        rb = xlrd.open_workbook(file)
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)

        row = 0
        while True:
            col = (listeuzunlugu * 6) - 2
            while True:
                w_sheet.write(row, col, )
                if (col == (listeuzunlugu * 6) + listeuzunlugu * 2 + 4):
                    break
                col += 1
            if (row == listeuzunlugu):
                break
            row += 1
        wb.save(file)
        while True:
            file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
            workbook = xlrd.open_workbook(file)
            sheet = workbook.sheet_by_index(0)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
            rb = xlrd.open_workbook(file)
            wb = copy(rb)
            w_sheet = wb.get_sheet(0)

            row = 0
            while True:
                col = 0
                while True:
                    w_sheet.write(row, col, )
                    if (col == 27):
                        break
                    col += 1
                if (row == kuzunluk):
                    break
                row += 1
            wb.save(file)

            file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
            workbook = xlrd.open_workbook(file)
            sheet = workbook.sheet_by_index(0)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
            rb = xlrd.open_workbook(file)
            wb = copy(rb)
            w_sheet = wb.get_sheet(0)

            h = 0
            while True:
                fkromozom = []
                index = random.sample(sayilar, r_uzunluk)

                g = 0
                for i in index:
                    w_sheet.write(h, g, i)
                    g += 1

                h += 1

                if (h == kuzunluk):
                    break

            wb.save(file)



            if(r_uzunluk == 1):
                g = 0
                c = 0
                n = 10000
                while True:
                    analiste = []
                    kk = []
                    file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
                    workbook = xlrd.open_workbook(file)
                    sheet = workbook.sheet_by_index(0)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

                    for row in range(0, r_uzunluk):
                        kk.append(sheet.cell_value(g, row))

                    file = "C:/Users/ASUS/Desktop/a.xls"
                    workbook = xlrd.open_workbook(file)
                    sheet = workbook.sheet_by_index(0)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

                    for i in kk:
                        i = int(i)
                        a = data[i][4]
                        analiste.append(a)
                    for i in kk:
                        i = int(i)
                        b = data[i][1]
                        analiste.append(b)

                    kafe = analiste[0]
                    kafe = int(kafe)
                    mahalle = analiste[1]
                    mahalle = int(mahalle)

                    file = "C:/Users/ASUS/Desktop/a.xls"
                    workbook = xlrd.open_workbook(file)
                    sheet = workbook.sheet_by_index(5)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

                    ist_kafe = data[2][kafe]
                    ist_kafe = int(ist_kafe)
                    gecensure = ist_kafe

                    sheet = workbook.sheet_by_index(4)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

                    kafe_mah = data[mahalle][kafe]
                    kafe_mah = int(kafe_mah)
                    gecensure += kafe_mah

                    gecensure += r_uzunluk * 4

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

                    if (g == kuzunluk - 1):
                        break
                    g += 1

            else:
                g = 0
                c = 0
                n = 10000
                while True:
                    analiste = []
                    kk = []
                    file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
                    workbook = xlrd.open_workbook(file)
                    sheet = workbook.sheet_by_index(0)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

                    for row in range(0, r_uzunluk):
                        kk.append(sheet.cell_value(g, row))

                    file = "C:/Users/ASUS/Desktop/a.xls"
                    workbook = xlrd.open_workbook(file)
                    sheet = workbook.sheet_by_index(0)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

                    for i in kk:
                        i = int(i)
                        a = data[i][4]
                        analiste.append(a)
                    for i in kk:
                        i = int(i)
                        b = data[i][1]
                        analiste.append(b)

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
                        if (h == r_uzunluk - 1):
                            break
                        h += 1

                    sheet = workbook.sheet_by_index(4)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

                    gkafe = analiste[r_uzunluk - 1]
                    gkafe = int(gkafe)
                    gmah = analiste[r_uzunluk]
                    gmah = int(gmah)
                    kaf_mah = data[gmah][gkafe]
                    kaf_mah = int(kaf_mah)
                    gecensure += kaf_mah

                    sheet = workbook.sheet_by_index(2)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
                    h = r_uzunluk
                    a = r_uzunluk + 1
                    while True:
                        glmah = analiste[h]
                        glmah = int(glmah)
                        gmah = analiste[a]
                        gmah = int(gmah)
                        mah_mah = data[glmah][gmah]
                        mah_mah = int(mah_mah)
                        gecensure += mah_mah
                        if (a == (r_uzunluk * 2) - 1):
                            break
                        h += 1
                        a += 1

                    gecensure += r_uzunluk * 4

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

                    if (g == kuzunluk - 1):
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
            for row in range(0, r_uzunluk):
                enkucuklist.append(sheet.cell_value(c, row))
            for row in range(0, r_uzunluk):
                enkucuklist.append(sheet.cell_value(c, row))

            sira = (listeuzunlugu * 2) + 5
            for i in enkucuklist:
                w_sheet.write(0, sira, i)
                sira += 1
            w_sheet.write(0, (listeuzunlugu * 2) + 5 + listeuzunlugu * 2 + 1, n)

            wb.save(file)



            son = 0
            ei = 1
            while True:
                if(r_uzunluk == 1):
                    break
                dongu = 0
                h = 0
                listindex = []
                while True:
                    file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
                    workbook = xlrd.open_workbook(file)
                    sheet = workbook.sheet_by_index(0)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

                    b = dongu
                    m = data[b][(listeuzunlugu * 2) + 2]
                    m = int(m)
                    listindex.append(b)

                    c = dongu + 1
                    n = data[c][(listeuzunlugu * 2) + 2]
                    n = int(n)
                    listindex.append(c)

                    liste1 = []
                    liste2 = []
                    for row in range(0, r_uzunluk):
                        liste1.append(sheet.cell_value(b, row))
                    for row in range(0, r_uzunluk):
                        liste2.append(sheet.cell_value(c, row))

                    file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
                    rb = xlrd.open_workbook(file)
                    wb = copy(rb)
                    w_sheet = wb.get_sheet(0)

                    a = (r_uzunluk / 2) + 1
                    a = int(a)
                    if (r_uzunluk == 2):
                        a = 1
                    sira = 0
                    for i in range(0, a):
                        deger = liste1[i]
                        deger2 = liste2[i]
                        w_sheet.write(b, sira, deger)
                        w_sheet.write(c, sira, deger2)
                        sira += 1

                    a = (r_uzunluk / 2) + 1
                    a = int(a)
                    if (r_uzunluk == 2):
                        a = 1
                    sira = a
                    for i in range(a, r_uzunluk):
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
                        olmayankafe = []
                        p = 0
                        for row in range(0, r_uzunluk):
                            degiseceklistkafe.append(sheet.cell_value(r, row))

                        for i in range(0, r_uzunluk):
                            p = 0
                            deger = liste1[i]
                            for j in degiseceklistkafe:
                                if (deger == j):
                                    p += 1
                            if (p == 0):
                                olmayankafe.append(deger)

                        t = 0
                        sira = 0
                        while True:
                            k = 0
                            degisken = data[r][sira]
                            degisken = int(degisken)
                            for i in range(sira + 1, r_uzunluk):
                                degisken2 = data[r][i]
                                degisken2 = int(degisken2)
                                if (degisken == degisken2):
                                    k += 1
                            if (k >= 1):
                                yazilacak = olmayankafe[t]
                                yazilacak = int(yazilacak)
                                w_sheet.write(r, sira, yazilacak)
                                t += 1

                            if (sira == r_uzunluk - 2):
                                break

                            sira += 1

                        h += 1
                        if (h == dur):
                            break

                        r = c



                    wb.save(file)
                    if (dongu == kuzunluk - 2):
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
                    listea = []
                    for row in range(0, r_uzunluk):
                        listea.append(sheet.cell_value(h, row))
                    sira = r_uzunluk
                    for i in listea:
                        w_sheet.write(h, sira, i)
                        sira += 1
                    wb.save(file)
                    file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
                    workbook = xlrd.open_workbook(file)
                    sheet = workbook.sheet_by_index(0)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
                    rb = xlrd.open_workbook(file)
                    wb = copy(rb)
                    w_sheet = wb.get_sheet(0)

                    index1 = random.randint(0, r_uzunluk - 1)
                    index2 = random.randint(0, r_uzunluk - 1)
                    a = data[h][index1]
                    b = data[h][index2]
                    w_sheet.write(h, index1, b)
                    w_sheet.write(h, index2, a)

                    index3 = random.randint(r_uzunluk, r_uzunluk * 2 - 1)
                    index4 = random.randint(r_uzunluk, r_uzunluk * 2 - 1)
                    c = data[h][index3]
                    d = data[h][index4]
                    w_sheet.write(h, index3, d)
                    w_sheet.write(h, index4, c)

                    h += 1
                    if (h == kuzunluk):
                        break

                wb.save(file)



                c = 0
                n = 10000
                g = 0
                while True:
                    analiste = []
                    kk = []
                    file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
                    workbook = xlrd.open_workbook(file)
                    sheet = workbook.sheet_by_index(0)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

                    for row in range(0, r_uzunluk * 2):
                        kk.append(sheet.cell_value(g, row))

                    file = "C:/Users/ASUS/Desktop/a.xls"
                    workbook = xlrd.open_workbook(file)
                    sheet = workbook.sheet_by_index(0)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

                    for i in range(0, r_uzunluk):
                        s = kk[i]
                        s = int(s)
                        a = data[s][4]
                        a = int(a)
                        analiste.append(a)
                    for i in range(r_uzunluk, r_uzunluk * 2):
                        s = kk[i]
                        s = int(s)
                        b = data[s][1]
                        b = int(b)
                        analiste.append(b)

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
                        if (h == r_uzunluk - 1):
                            break
                        h += 1

                    sheet = workbook.sheet_by_index(4)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

                    gkafe = analiste[r_uzunluk - 1]
                    gkafe = int(gkafe)
                    gmah = analiste[r_uzunluk]
                    gmah = int(gmah)
                    kaf_mah = data[gmah][gkafe]
                    kaf_mah = int(kaf_mah)
                    gecensure += kaf_mah

                    sheet = workbook.sheet_by_index(2)
                    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
                    h = r_uzunluk
                    a = r_uzunluk + 1
                    while True:
                        glmah = analiste[h]
                        glmah = int(glmah)
                        gmah = analiste[a]
                        gmah = int(gmah)
                        mah_mah = data[glmah][gmah]
                        mah_mah = int(mah_mah)
                        gecensure += mah_mah
                        if (a == (r_uzunluk * 2) - 1):
                            break
                        h += 1
                        a += 1

                    gecensure += r_uzunluk * 4

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

                    if (g == kuzunluk - 1):
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
                for row in range(0, r_uzunluk * 2):
                    enkucuklist.append(sheet.cell_value(c, row))

                sira = (listeuzunlugu * 2) + 5
                for i in enkucuklist:
                    w_sheet.write(ei, sira, i)
                    sira += 1
                w_sheet.write(ei, (listeuzunlugu * 2) + 5 + listeuzunlugu * 2 + 1, n)

                wb.save(file)



                if (son == 10):
                    break

                ei += 1
                son += 1

            file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
            workbook = xlrd.open_workbook(file)
            sheet = workbook.sheet_by_index(0)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
            rb = xlrd.open_workbook(file)
            wb = copy(rb)
            w_sheet = wb.get_sheet(0)

            g = 0
            b = 0
            m = 10000
            while True:
                a = data[g][(listeuzunlugu * 2) + 5 + listeuzunlugu * 2 + 1]
                a = int(a)

                print("ÇAPRAZLAMALAR SONUCU EN KÜÇÜK DEĞERLER:",a)
                if (a < m):
                    m = a
                    b = g

                if(r_uzunluk == 1):
                    break
                if (g == 10):
                    break

                g += 1

            klist = []
            for i in range(((listeuzunlugu * 2) + 5), (listeuzunlugu * 2) + 5 + r_uzunluk * 2):
                v = data[b][i]
                v = int(v)
                klist.append(v)

            sira = (listeuzunlugu * 6)
            for i in klist:
                w_sheet.write(r_uzunluk - 1, sira, i)
                sira += 1
            w_sheet.write(r_uzunluk - 1, (listeuzunlugu * 6) - 2, m)

            wb.save(file)
            ass = 0
            if (m > 40):
                ass += 1
                break
            if (r_uzunluk == len(sayilar)):
                break

            if (listeuzunlugu == 4 and r_uzunluk == 3):
                break

            r_uzunluk = r_uzunluk + 1

        file = "C:/Users/ASUS/Desktop/Kromozom_Populasyonu.xls"
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheet_by_index(0)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        rb = xlrd.open_workbook(file)
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)

        listem = []
        rotalistesi.append(listem)
        if (ass == 1):
            for i in range(listeuzunlugu * 6, listeuzunlugu * 6 + (r_uzunluk - 1) * 2):
                rotalistesi[list_ind].append(sheet.cell_value(r_uzunluk - 2, i))
            surelistesi.append(sheet.cell_value(r_uzunluk - 2, listeuzunlugu * 6 - 2))

        else:
            for i in range(listeuzunlugu * 6, listeuzunlugu * 6 + r_uzunluk * 2):
                rotalistesi[list_ind].append(sheet.cell_value(r_uzunluk - 1, i))
            surelistesi.append(sheet.cell_value(r_uzunluk - 1, listeuzunlugu * 6 - 2))

        sayilar = []
        for i in range(0, listeuzunlugu):
            k = 0
            for w in rotalistesi:
                for c in w:
                    if (i == c):
                        k += 1
            if (k == 0):
                sayilar.append(i)

        r_uzunluk = 1

        if (len(sayilar) == 0):
            break

        if (len(sayilar) == 1):
            analiste = []
            file = "C:/Users/ASUS/Desktop/a.xls"
            workbook = xlrd.open_workbook(file)
            sheet = workbook.sheet_by_index(0)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

            s = sayilar[0]
            s = int(s)
            a = data[s][4]
            a = int(a)
            analiste.append(a)

            b = data[s][1]
            b = int(b)
            analiste.append(b)

            kafe = analiste[0]
            kafe = int(kafe)
            mahalle = analiste[1]
            mahalle = int(mahalle)

            sheet = workbook.sheet_by_index(5)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

            ist_kafe = data[2][kafe]
            ist_kafe = int(ist_kafe)
            gecensure = ist_kafe

            sheet = workbook.sheet_by_index(4)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

            kafe_mah = data[mahalle][kafe]
            kafe_mah = int(kafe_mah)
            gecensure += kafe_mah
            gecensure += r_uzunluk * 4

            listem = []
            rotalistesi.append(listem)
            list_ind += 1
            i = 0
            rotalistesi[list_ind].append(sayilar[i])
            rotalistesi[list_ind].append(sayilar[i])
            surelistesi.append(gecensure)
            break

        list_ind += 1

    print(rotalistesi)
    print(surelistesi)





    g = 0
    toplam_rota_km = 0
    gecenkm = 0
    while True:
        analiste = []
        kk = []

        liste = rotalistesi[g]

        for k in liste:
            kk.append(k)

        print("Rota kodlaması:", kk)

        if (len(kk) == 2):
            file = "C:/Users/ASUS/Desktop/a.xls"
            workbook = xlrd.open_workbook(file)
            sheet = workbook.sheet_by_index(0)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

            s = kk[0]
            s = int(s)
            a = data[s][4]
            a = int(a)
            analiste.append(a)

            b = data[s][1]
            b = int(b)
            analiste.append(b)

            print("Açık rota:", analiste)

            kafe = analiste[0]
            kafe = int(kafe)
            mahalle = analiste[1]
            mahalle = int(mahalle)

            file = "C:/Users/ASUS/Desktop/a.xls"
            workbook = xlrd.open_workbook(file)
            sheet = workbook.sheet_by_index(6)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

            ist_kafe = data[2][kafe]
            gecenkm += ist_kafe

            sheet = workbook.sheet_by_index(9)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

            kafe_mah = data[mahalle][kafe]
            gecenkm += kafe_mah

            print("Rotada geçen Km:", gecenkm)


        else:
            file = "C:/Users/ASUS/Desktop/a.xls"
            workbook = xlrd.open_workbook(file)
            sheet = workbook.sheet_by_index(0)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

            sira = len(kk) / 2
            sira = int(sira)
            for i in range(0, sira):
                s = kk[i]
                s = int(s)
                a = data[s][4]
                analiste.append(a)
            for i in range(sira, sira * 2):
                s = kk[i]
                s = int(s)
                b = data[s][1]
                analiste.append(b)

            print("Açık rota:", analiste)

            kafe1 = analiste[0]
            kafe1 = int(kafe1)

            file = "C:/Users/ASUS/Desktop/a.xls"
            workbook = xlrd.open_workbook(file)
            sheet = workbook.sheet_by_index(6)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

            ist_kafe = data[2][kafe1]
            gecenkm += ist_kafe

            sheet = workbook.sheet_by_index(7)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
            h = 1
            for i in analiste:
                i = int(i)
                gkafe = analiste[h]
                gkafe = int(gkafe)
                kafe_kafe = data[i][gkafe]
                gecenkm += kafe_kafe
                if (h == sira - 1):
                    break
                h += 1

            sheet = workbook.sheet_by_index(9)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

            gkafe = analiste[sira - 1]
            gkafe = int(gkafe)
            gmah = analiste[sira]
            gmah = int(gmah)
            kaf_mah = data[gmah][gkafe]
            gecenkm += kaf_mah

            sheet = workbook.sheet_by_index(8)
            data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
            h = sira
            a = sira + 1
            while True:
                glmah = analiste[h]
                glmah = int(glmah)
                gmah = analiste[a]
                gmah = int(gmah)
                mah_mah = data[glmah][gmah]
                gecenkm += mah_mah
                if (a == (sira * 2) - 1):
                    break
                h += 1
                a += 1

            print("Rotada geçen Km:", gecenkm)

        if (g == len(rotalistesi) - 1):
            break

        g += 1

    toplam_rota_km += gecenkm

    print("Toplam Km:", toplam_rota_km)

    print("Rotalarda harcanan toplam KM:", toplam_rota_km)

    yakit_masrafi = toplam_rota_km * 0.1668

    print("Toplam yakıt masrafı:", yakit_masrafi)


    toplam_yakit_masrafi += yakit_masrafi


    file = "C:/Users/ASUS/Desktop/Kurye.xls"
    rb = xlrd.open_workbook(file)
    sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)
    data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

    for i in surelistesi:
        w_sheet.write(issuresi, 2, i)
        w_sheet.write(issuresi, 0, L)

        wb.save(file)
        file = "C:/Users/ASUS/Desktop/Kurye.xls"
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

        grup = data[issuresi][0]
        grup = int(grup)
        vakit = data[issuresi][2]
        vakit = int(vakit)
        formul = (vakit / 5) + grup
        w_sheet.write(issuresi, 3, formul)

        wb.save(file)
        file = "C:/Users/ASUS/Desktop/Kurye.xls"
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

        w_sheet.write(1, 1, 1)
        listindeks = data[issuresi][0]
        listindeks = int(listindeks)
        for i in range(1, issuresi):
            islemsirasi = data[i][3]
            islemsirasi = int(islemsirasi)

            if (islemsirasi <= listindeks):
                kurye = data[i][1]
                kurye = int(kurye)
                w_sheet.write(issuresi, 1, kurye)
                w_sheet.write(i, 3, 100000)
                break

            else:
                w_sheet.write(issuresi, 1, kuryenumarasi)

        wb.save(file)
        file = "C:/Users/ASUS/Desktop/Kurye.xls"
        rb = xlrd.open_workbook(file)
        sheet = rb.sheet_by_index(0)
        wb = copy(rb)
        w_sheet = wb.get_sheet(0)
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

        v = data[issuresi][1]
        v = int(v)
        if (kuryenumarasi == v):
            kuryenumarasi += 1

        issuresi += 1
        wb.save(file)

    wb.save(file)
    print(bt*5,"dakikadayız.")
    L += 1
    if (bt == 24):
        break

    bt += 1

print(""""











    """)
print("ÇALIŞTIRILMASI GEREKEN KURYE ELEMANI SAYISI:", kuryenumarasi - 1)
print("Saat dilimindeki toplam yakıt masrafı:",toplam_yakit_masrafi)
