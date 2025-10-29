"""
Cotizador para fumigacion con dron
"""

import toga
from toga.style import Pack
from toga.style.pack import COLUMN, ROW
import pandas as pd
import requests as r
import asyncio
#from cotizador import calc
sheet_name = 'Hoja1'
#columns = 'B:G'
excel_file_path = ""
NO_TABULADOR = "Tabulador no encontrado. Conéctese a internet"
#Maximo de hectareas
MAX_H_VALUE = 5000

class cotizador(toga.App):

    def getData(self):
            try:
                print(11111)
                url=('https://1drv.ms/x/c/9f79dc42cb78dc5e/EVQwX_3G4ZFHjb_htK6vv8QBMdvrze3LRweEYoZmHBZwog?download=1')
                global excel_file_path
                excel_file_path = self.app.paths.data / 'Tabulador Fumigacion con Dron.xlsx'
                response = r.get(url)
                #print(self.app.paths.app / "resources/tarjeta.png")

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
                #print(self.app.paths.data)

    def startup(self):
        print("startup*******")
        self.main_box = toga.Box(direction=COLUMN)

        self.area_label = toga.Label(
            "Hectáreas: ",
            margin=(0, 5),
        )
        self.amount_label = toga.Label(
            "Litros/ha: ",
            margin=(0, 5),
        )
        self.image = toga.Image(self.app.paths.app / "resources/tarjeta.png")
        self.image_view = toga.ImageView(self.image, style=Pack(direction=COLUMN, height=200, flex=True))
    
        self.image_box = toga.Box(children=[self.image_view,], style=Pack(direction=COLUMN, flex=True))
        
        

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

        self.button = toga.Button(
            "Calcula costo",
            on_press=self.calculate,
            margin=5,
        )

        self.button_2 = toga.Button(
            "DOSIFICADOR",
            on_press=self.dosificador,
            margin=5,
        )

        self.button_3 = toga.Button("COTIZADOR",
                               on_press=self.cotizador,
                               margin=5,)
        
        self.button_4 = toga.Button("Calcula la dosis por mezclador",
                               on_press=self.calculaDosis,
                               margin=5,)

        self.main_box.add(self.input_box)
        self.main_box.add(self.button)
        self.main_box.add(self.output_box)
        self.main_box.add(self.image_box)
        self.main_box.add(self.button_2)

        self.main_window = toga.MainWindow(title=self.formal_name)
        self.main_window.content = self.main_box
        self.main_window.show()
        print("showwwwwwww")
           
    def on_running(self):
        self.getData() 
        print ("after GetData")    

    def cotizador(self, widget):
        self.main_box.remove(self.input_box, self.output_box, self.image_box, self.button_3, self.button_4)

        self.image_box = toga.Box(children=[self.image_view,], style=Pack(direction=COLUMN, flex=True))
        
        

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

        self.button = toga.Button(
            "Calcula costo",
            on_press=self.calculate,
            margin=5,
        )

        self.button_2 = toga.Button(
            "DOSIFICADOR",
            on_press=self.dosificador,
            margin=5,
        )

        self.button_3 = toga.Button("COTIZADOR",
                               on_press=self.cotizador,
                               margin=5,)
        
        self.button_4 = toga.Button("Calcula la dosis por mezclador",
                               on_press=self.calculaDosis,
                               margin=5,)

        self.main_box.add(self.input_box)
        self.main_box.add(self.button)
        self.main_box.add(self.output_box)
        self.main_box.add(self.image_box)
        self.main_box.add(self.button_2)

        #self.main_window = toga.MainWindow(title=self.formal_name)
        #self.main_window.content = self.main_box
        #self.main_window.show()
        #print("showwwwwwww")

    def redraw(self):
      
        self.input_box.add(self.area)
        self.input_box.add(self.amount_label)
        self.input_box.add(self.amount)
        self.output_box.add(self.result)
        self.output_box.add(self.conn_status)

    def calculate(self, widget):
        
        try:
            h=(self.area.value)
            f=(self.amount.value)
            if self.conn_status.text == NO_TABULADOR:
                self.getData()
            df = pd.read_excel(excel_file_path, sheet_name, header=None)
            # Reference to value that defines the cells to select of tabulator
            columns = df.iloc[8, 0] 
            print(columns)   
            df = pd.read_excel(excel_file_path, sheet_name, index_col=0, skiprows=11,  nrows=57, usecols= columns)
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
            self.conn_status.text = NO_TABULADOR

        except Exception as e:
            print(f"An error occurred: {e}")

    def dosificador(self, widget):
        self.main_box.remove(self.input_box, self.button, self.output_box, self.image_box, self.button_2, self.button_4)
        self.amount_label = toga.Label(
            "Litros/ha: ",
            margin=(0, 5),
        )
        self.litrosMezclador = toga.NumberInput(flex=1,)
        self.litrosporhectarea = toga.NumberInput(flex=1)
        self.litrosmezclaporhectarea = toga.NumberInput(flex=1)
        self.dosispormezclador = toga.Label("") 

        self.dosis_label = toga.Label(
            "Producto por hectárea (Litros): ",
            margin=(0, 5),
        )
        self.litrosdelmezclador_label = toga.Label(
            "Capacidad del mezcador (Litros): ",
            margin=(0, 5),
        )
        self.dosispormezclador_label = toga.Label(
            "",
            margin=(0, 5),
        )
        self.litrosmezclaporhectara_label = toga.Label(
            "Mezcla por hectárea (Litros): ",
            margin=(0, 5),
        )

        self.input_box = toga.Box(direction=COLUMN, margin=5)
        self.input_box.add(self.dosis_label)
        self.input_box.add(self.litrosporhectarea)
        self.input_box.add(self.litrosdelmezclador_label)
        self.input_box.add(self.litrosMezclador)
        self.input_box.add(self.litrosmezclaporhectara_label)
        self.input_box.add(self.litrosmezclaporhectarea)

        self.output_box = toga.Box(direction=COLUMN, margin=5)
        self.output_box.add(self.dosispormezclador_label)
        

        self.main_box.add(self.input_box)
        self.main_box.add(self.button_4)
        self.main_box.add(self.output_box)
        self.main_box.add(self.image_box)
        self.main_box.add(self.button_3)

        #self.main_window = toga.MainWindow(title=self.formal_name)
        #self.main_window.content = self.main_box
        #self.main_window.show()

    def calculaDosis(self, widget):
        self.hectareaspormezclador = self.litrosMezclador.value/self.litrosmezclaporhectarea.value
        print(f"Hectareas por mezclador = {self.hectareaspormezclador}")
        self.dosispormezclador = (self.litrosporhectarea.value)* (self.hectareaspormezclador)
        self.dosispormezclador_label.text = f"Dosis por mezclador: {round(self.dosispormezclador, 2)} Litros"        


def main():
    return cotizador()
