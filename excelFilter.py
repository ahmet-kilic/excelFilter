import xlrd
import xlsxwriter
import sys

x = input("Çevrilecek excel dosya ismi: ")

path = x + ".xlsx"

inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)

if inputWorksheet.ncols != 10:
    print("Sütun sayısı farklı.. Programdan çıkılıyor..")
    sys.exit()

date = inputWorksheet.cell_value(0,6)

lplates = []

for y in range(0, inputWorksheet.nrows):
    if inputWorksheet.cell_value(y,5) not in lplates:
        lplates.append(inputWorksheet.cell_value(y,5))

plateVal = {}

for x in lplates:
    plateVal[x] = []

for y in range(0, inputWorksheet.nrows):
    plateVal[inputWorksheet.cell_value(y,5)].append(inputWorksheet.cell_value(y,9))

max = 0
total = 0
sums = 0
for x in lplates:
    if max < len(plateVal[x]):
        max = len(plateVal[x])
    total += sum(plateVal[x])
    sums += len(plateVal[x])

outWorkbook = xlsxwriter.Workbook("out.xlsx")
outSheet = outWorkbook.add_worksheet()

cell_format1 = outWorkbook.add_format()
cell_format1.set_num_format('##,###')

cell_format3 = outWorkbook.add_format()
cell_format3.set_num_format('##,###')
cell_format3.set_bold()

cell_format4 = outWorkbook.add_format()
cell_format4.set_bold()

for i in range(len(lplates)):
    outSheet.write(2,i, len(plateVal[lplates[i]]), cell_format4)
    outSheet.write(3,i, lplates[i], cell_format4)
    for j in range(len(plateVal[lplates[i]])):
        outSheet.write(j + 4,i, plateVal[lplates[i]][j], cell_format1)
    outSheet.write(max + 4, i, sum(plateVal[lplates[i]]), cell_format3)

finalString = "%d KG %d SEFER" % (total, sums)
outSheet.write(max+5,0, finalString, cell_format4)

cell_format2 = outWorkbook.add_format()
cell_format2.set_num_format('dd/mm/yy')
outSheet.write(0,0, inputWorksheet.cell_value(0,6), cell_format2)


outWorkbook.close()

print("Başarılı")