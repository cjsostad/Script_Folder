import openpyxl
from openpyxl import Workbook, load_workbook

#load an excel sheet
book = load_workbook(filename = r'\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\mcunning\Script_Folder_Marina\FSW_Objectives_Copy.xlsx')

#printsheetnames to verify
print(book.sheetnames)

#View active data (don't quite understand this function, is it needed?)
#all_fsw = book.active

#create new sheets that extract specific data
sheet0 = book.create_sheet("ALL_FISH_SENSITIVE_WS_POLY_ID", 0)

#Get the sheetnames in the file (0-12, 14 total)

#Label columns
sheet0["A1"] = "FSW_TAG"
sheet0["B1"] = "FISH_SENSITIVE_WS_POLY_ID"
sheet0["C1"] = "OBJECTIVE_1"
sheet0["D1"] = "OBJECTIVE_2"
sheet0["E1"] = "OBJECTIVE_3"
sheet0["F1"] = "OBJECTIVE_4"
sheet0["G1"] = "OBJECTIVE_5"
sheet0["H1"] = "OBJECTIVE_6"
sheet0["I1"] = "OBJECTIVE_7"
sheet0["J1"] = "OBJECTIVE_8"
sheet0["K1"] = "OBJECTIVE_9"

#define all the sheets
sheet1 = book["FSW_ID_1_001_to_F_1_011_edit"]
sheet2 = book["FSW_ID_F_4_001_edit"]
sheet3 = book["FSW_ID_F_6_001_to_F_6_005_edit"]
sheet4 = book["FSW_ID_F_3_001_to_F_3_006_edit"]
sheet5 = book["FSW_ID_F_8_001_to_F_8_008_edit"]
sheet6 = book["FSW_ID_F_7_001_and_F_7_005_edit"]
sheet7 = book["FSW_ID_F_7_002_edit"]
sheet8 = book["FSW_ID_F_7_003_F_7_004_edit"]
sheet9 = book["FSW_ID_F_7_006_to_F_7_008_edit"]
sheet10 = book["FSW_ID_F_7_019_edit"]
sheet11 = book["FSW_ID_F_7_020_to_F_7_023_edit"]
sheet12 = book["FSW_F_3_007_and_F_3_008_edit"]
sheet13 = book["FSW_ID_F_3_009_to_F_3_014_edit"]
sheet14 = book["FSW_ID_F_5_001_edit"]


#delete first row in all sheets so there is not duplication(keep for later)
#sheet1.delete_rows(idx=1, amount=1)
#sheet2.delete_rows(idx=1, amount=1)
#sheet3.delete_rows(idx=1, amount=1)
#sheet4.delete_rows(idx=1, amount=1)
#sheet5.delete_rows(idx=1, amount=1)
#sheet6.delete_rows(idx=1, amount=1)
#sheet7.delete_rows(idx=1, amount=1)
#sheet8.delete_rows(idx=1, amount=1)
#sheet9.delete_rows(idx=1, amount=1)
#sheet10.delete_rows(idx=1, amount=1)
#sheet11.delete_rows(idx=1, amount=1)
#sheet12.delete_rows(idx=1, amount=1)
#sheet13.delete_rows(idx=1, amount=1)
#sheet14.delete_rows(idx=1, amount=1)

#Append the data (didnt work)
#for columns in sheet1:
    #sheet0.append(columns)

#book.copy_worksheet(sheet1)
#print a column
#print(sheet1.columns=1)

#Copy over the data from (did not work) sheet1, sheet2, sheet3, sheet4, sheet5, sheet6, sheet7, sheet8, sheet9, sheet10, sheet11, sheet12, sheet13)
#target = book.copy_worksheet(sheet1)
#Append the data from the sheets into all_fsw(this append function didnt work) sheet1, sheet2, sheet3, sheet4, sheet5, sheet6, sheet7, sheet8, sheet9, sheet10, sheet11, sheet12, sheet13)
#all_fsw.append(sheet1) 

#possible code solutions
#Dictionary or classes 
#iterate through all operational sheets to extract row and column data

book.save (filename = r'\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\mcunning\Script_Folder_Marina\FSW_Objectives_Copy_Python.xlsx')



