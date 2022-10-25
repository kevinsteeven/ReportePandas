from openpyxl import load_workbook
from openpyxl.styles import Font
wb=load_workbook('report.xlsx')
sheet = wb['Report']
#Escribir titulo en celda A1
sheet['A1']='Reporte de ventas'
#Escribir subtitulo en celda A2
sheet['A2']='Octubre'
#Editar fuentes
sheet['A1'].font=Font('Arial', bold= True, size=20)
sheet['A2'].font=Font('Arial', bold= True, size=10)
wb.save('reporte_octubre.xlsx')