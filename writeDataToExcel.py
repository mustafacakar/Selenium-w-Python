import openpyxl

path = r"C:\Users\MUSTAFA\Desktop\writeDataToExcel.xlsx"
workbook = openpyxl.load_workbook(path)
sheet = workbook.active


for r in range(1,6):
    for c in range(1,4):
        sheet.cell(row=r,column=c).value = "Deneme Yazı"


workbook.save(path)
