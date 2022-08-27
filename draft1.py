#for a single file at a time

import openpyxl                                                      #library to interact with python
from win32com import client                                          #for converting the final result to a pdf
path = input("Enter the path of the Excel file")                     #get path for the salary structure excel 
pathout = input("Enter the destination path for the pdf file") 
cell = input("Enter the cell address for ctc")                       #get the cell address for the cell to eter ctc
wb=openpyxl.load_workbook(path)                                      #opening the excel file
sh=wb['Sheet1']                                                      #select the sheet

ctc=input("enter the ctc")                                           #getting value for ctc 
wb[cell]=ctc                                                         #inserting the ctc value into the formulated excel
wb.save()                                                            #save the ecxel with the new ctc value

excel = client.Dispatch("Excel.Application")                         #open excel
  
excel.Interactive = False
excel.Visible = False
sheets = excel.Workbooks.Open(path)                                    
work_sheets = sheets.Worksheets[0]                                    #read excel file

work_sheets.ExportAsFixedFormat(0, pathout)                           #convert to pdf and save at the path mentioned for output
work_sheet.close()                                                    #close excel file
