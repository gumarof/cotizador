"""
Cotizador para fumigacion con dron
"""

import toga
from toga.style import Pack
from toga.style.pack import COLUMN, ROW
import pandas as pd
import requests as r
#from cotizador import calc
sheet_name = 'Hoja1'
excel_file_path = ""
#Maximo de hectareas
MAX_H_VALUE = 5000

class cotizador(toga.App):
    def startup(self):
        main_box = toga.Box(direction=COLUMN)

        self.area_label = toga.Label(
            "Hectáreas: ",
            margin=(0, 5),
        )
        self.amount_label = toga.Label(
            "Litros/ha: ",
            margin=(0, 5),
        )
        self.area = toga.NumberInput(flex=1,)
        self.amount = toga.NumberInput(flex=1)
        self.result = toga.Label("") 
        self.conn_status = toga.Label("", style=Pack(color="red")) 

        self.input_box = toga.Box(direction=COLUMN, margin=5)
        self.input_box.add(self.area_label)
        self.input_box.add(self.area)
        self.input_box.add(self.amount_label)
        self.input_box.add(self.amount)

        self.output_box = toga.Box(direction=COLUMN, margin=5)
        self.output_box.add(self.result)
        self.output_box.add(self.conn_status)

        button = toga.Button(
            "Calcula costo",
            on_press=self.calculate,
            margin=5,
        )

        main_box.add(self.input_box)
        main_box.add(button)
        main_box.add(self.output_box)

        self.main_window = toga.MainWindow(title=self.formal_name)
        self.main_window.content = main_box
        self.main_window.show()

        try:
            url=('https://1drv.ms/x/c/9f79dc42cb78dc5e/EVQwX_3G4ZFHjb_htK6vv8QBMdvrze3LRweEYoZmHBZwog?download=1')
            global excel_file_path
            excel_file_path = self.app.paths.data / 'Tabulador Fumigacion con Dron.xlsx'
            response = r.get(url)

            if response.status_code == 200:
                self.conn_status.text =""
                # Data retrieved successfully
                # For binary files like Excel, you would save content to a file
                with open(excel_file_path, "wb") as f:
                    f.write(response.content)
            else:
                print(f"Error: Could not retrieve data. Status code: {response.status_code}")
            
        except Exception as e:
            self.conn_status.text = "Tabulador no actualizado."
            print("Tabulador no actualizado. Revise su conexión a internet %s" %e)

    def redraw(self):
        self.input_box.add(self.area_label)
        self.input_box.add(self.area)
        self.input_box.add(self.amount_label)
        self.input_box.add(self.amount)
        self.output_box.add(self.result)
        self.output_box.add(self.conn_status)

    def calculate(self, widget):
        
        try:
            h=(self.area.value)
            f=(self.amount.value)
            df = pd.read_excel(excel_file_path, sheet_name, index_col=0, skiprows=11,  nrows=57, usecols='B:F')
            
            # Busca índice de las hectáreas a fumigar
            def findColIndex():
                for column in df.columns:
                    x = column.split('-')
                    if h is None:
                        self.result.text = "Introduzca valores mayores a 0"
                        self.redraw()
                        raise ValueError("No hay valor")
                        """ 
                    elif x == ['100+']:
                        if h>100:
                            col_index = column
                            return col_index
                        """
                    elif h<=0:
                        self.result.text = "Los valores de entrada deben ser mayores a 0"
                        self.redraw()
                        raise ValueError("El valor debe ser mayor a cero")
                    elif h> MAX_H_VALUE:
                        self.result.text = "Las hectáreas deben ser menores a " + str(MAX_H_VALUE)
                        self.redraw()

                    elif h>=int(x[0]) and h<=int(x[1]):
                        col_index = column
                        return col_index
            
            # busca indice de los litrosnpor hectárea
            def findRowIndex():
                for row in df.index:
                    x = row.split('-')
                    if f is None:
                        self.result.text = "Introduzca valores mayores a 0"
                        self.redraw()
                    elif f > 300:
                        self.result.text = "Litros/ha no pueden ser mayores a 300"
                        self.redraw()
                        raise ValueError("Fuera de rango")
                    elif f>=int(x[0]) and f<=int(x[1]):
                        row_index = row
                        return row_index
                    elif f<=0:
                        self.result.text = "Los valores de entrada deben ser mayores a 0"
                        self.redraw()
                        raise ValueError("El valor debe ser mayor a cero")

            row_index = findRowIndex()
            col_index = findColIndex()
            price_per_hectarea = df.loc[row_index, col_index]
            print("Precio por hectárea = $%s" %price_per_hectarea)
            
            total = h*price_per_hectarea
            print(f"Hello, {self.area.value}")
            
            print("Total = $%s" %total)
            self.result.text = "\nPrecio por ha = $" + str(price_per_hectarea) + "\n\nTotal = $" + str(total)
            self.redraw()
            
            
            
            
        except FileNotFoundError:
            print(f"Error: Excel file not found at {excel_file_path}")
        except Exception as e:
            print(f"An error occurred: {e}")	
            
                


def main():
    return cotizador()
