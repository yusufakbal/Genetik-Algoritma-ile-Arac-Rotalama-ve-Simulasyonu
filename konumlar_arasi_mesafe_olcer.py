import googlemaps
import xlrd

file = "C:/Users/ASUS/Desktop/MESAFEÖLÇER.xlsx"
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_index(0)
data = [[sheet.cell_value(r ,c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]

mahalleler = []

for i in range(0, 20):
    mahalleler.append(sheet.cell_value(i, 9))

for i in mahalleler:

    gmaps = googlemaps.Client(key="AIzaSyDCS0xBMic6lkMdvA0eJsskemtFCxH7IJ8")

    my_dist = gmaps.distance_matrix('Eskibağlar, 26170 Tepebaşı/Eskişehir', i)['rows'][0]['elements'][0]

    print(my_dist)
