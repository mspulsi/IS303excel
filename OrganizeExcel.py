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

    newWB[sheetName]['F2'] = "Highest Grade"
    newWB[sheetName]['F3'] = "Lowest Grade"
    newWB[sheetName]['F4'] = "Mean Grade"
    newWB[sheetName]['F5'] = "Median Grade"
    newWB[sheetName]['F6'] = "Number of Students"

    iLastRow = newWB[sheetName].max_row
    newWB[sheetName]["G2"] = f"=MAX(D2:D{iLastRow})"
    newWB[sheetName]["G3"] = f"=MIN(D2:D{iLastRow})"
    newWB[sheetName]["G4"] = f"=AVERAGE(D2:D{iLastRow})"
    newWB[sheetName]["G5"] = f"=MEDIAN(D2:D{iLastRow})"
    newWB[sheetName]["G6"] = f"=COUNT(D2:D{iLastRow})"

newWB.save('formatted_grades.xlsx')
newWB.close()
oldWB.close()