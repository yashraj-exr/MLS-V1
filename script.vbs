Set objExcel = CreateObject("Excel.Application")



objExcel.Application.Run  "'C:\Users\hp\Desktop\yash4\programming\MLS - Copy\loanData.xlsm'!Module1.import_data"

objExcel.DisplayAlerts = False

objExcel.Application.Quit

Set objExcel = Nothing

