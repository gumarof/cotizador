import pandas as pd
import requests as r
from cotizador import app

url=('https://1drv.ms/x/c/9f79dc42cb78dc5e/EVQwX_3G4ZFHjb_htK6vv8QBMdvrze3LRweEYoZmHBZwog?download=1')

try:
    response = r.get(url)

    if response.status_code == 200:
        # Data retrieved successfully
        # For binary files like Excel, you would save content to a file
        with open("Tabulador Fumigacion con Dron.xlsx", "wb") as f:
            f.write(response.content)
    else:
        print(f"Error: Could not retrieve data. Status code: {response.status_code}")

    excel_file = 'Tabulador Fumigacion con Dron.xlsx'
    sheet_name = 'Hoja1'

    h=int(app.cotizador.area)
    f=int(app.cotizador.amount)

    excel_file_path = 'Tabulador Fumigacion con Dron.xlsx'
except:
    print("Tabulador no actualizado. Revise su conexión a internet")

try:
    df = pd.read_excel(excel_file_path, sheet_name, index_col=0, skiprows=11,  nrows=57, usecols='B:F')
    
    # Busca índice de las hectáreas a fumigar
    def findColIndex():
        for column in df.columns:
            x = column.split('-')
            
            if x == ['100+']:
                if h>100:
                    col_index = column
                    return col_index
                elif h<=0:
                    raise ValueError("El valor debe ser mayor a cero")
            elif h>=int(x[0]) and h<=int(x[1]):
                col_index = column
                return col_index
     
    # busca indice de los litrosnpor hectárea
    def findRowIndex():
        for row in df.index:
            x = row.split('-')
            
            if f>=int(x[0]) and f<=int(x[1]):
                row_index = row
                return row_index
            elif f<=0:
                raise ValueError("El valor debe ser mayor a cero")

    row_index = findRowIndex()
    col_index = findColIndex()
    price_per_hectarea = df.loc[row_index, col_index]
    print("Precio por hectárea = $%s" %price_per_hectarea)
    
    total = h*price_per_hectarea
    print("Total = $%s" %total)
    
except FileNotFoundError:
    print(f"Error: Excel file not found at {excel_file_path}")
except Exception as e:
    print(f"An error occurred: {e}")	