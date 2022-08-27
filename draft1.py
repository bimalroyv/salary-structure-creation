import openpyxl                                                      #library to interact with python
from win32com import client                                          #for converting the final result to a pdf
path = input("Enter the path of the Excel file")                     #get path for the salary structure excel 
cell = input("Enter the cell address for ctc")                       #get the cell address for the cell to eter ctc
wb=openpyxl.load_workbook(path)                                      #opening the excel file
sh=wb['Sheet1']                                                      #select the sheet

ctc=input("enter the ctc")                                           #getting value for ctc 
wb[cell]=ctc                                                         #inserting the ctc value into the formulated excel
wb.save()                                                            #save the ecxel with the new ctc value

app = client.DispatchEx("Excel.Application")                         #lines 12 to 16 for converting the saved excel sheet to a pdf format
app.Interactive = False
app.Visible = False
wb.ActveSheet.ExportAsFixedFormat(0,path)
wb.Close()

