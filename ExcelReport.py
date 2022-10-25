import pandas as pd
df=pd.read_excel(r'C:\Users\natal\OneDrive\Documentos\cosasKevin\python\ExcelReport\supermarket_sales.xlsx')
df= df[['Gender','Product line','Total']]
pivot_table=df.pivot_table(index='Gender', columns='Product line', values='Total',aggfunc='sum').round(0)
pivot_table.to_excel('pivot_table.xlsx','Report',startrow=4)#Primer parametro es nombre del archivo, segundo parametro nombre de la hoja