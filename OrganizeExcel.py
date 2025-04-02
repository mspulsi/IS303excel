import openpyxl
from openpyxl import Workbook, load_workbook

oldWB = load_workbook('/Users/mspul/Documents/GitHub/IS303excel/Poorly_Organized_Data_2.xlsx')
newWB = Workbook()
newWB.remove(newWB["Sheet"])

currSheet = oldWB.active
currClassName = ""
prevClassName = ""
firstRow = True

max_row = currSheet.max_row
for row in currSheet.iter_rows(min_row=2, max_row=max_row, max_col=3):
    for cell in row:
        if cell.col_idx == 1:
            currClassName = cell.value
            if currClassName != prevClassName:
                newWB.create_sheet(currClassName)
            prevClassName = currClassName
        elif cell.col_idx == 2:
            studentInfo = cell.value.split('_')
        else:
            studentInfo.append(cell.value)
    newWB[currClassName].append(studentInfo)

for sheetName in newWB.sheetnames:
    newWB[sheetName].insert_rows(1)
    newWB[sheetName]['A1'] = "Last Name"
    newWB[sheetName]['B1'] = "First Name"
    newWB[sheetName]['C1'] = "Student ID"
    newWB[sheetName]['D1'] = "Grade"

newWB.save('Organized_Data.xlsx')
newWB.close()
oldWB.close()