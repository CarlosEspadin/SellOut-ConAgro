#!/C:/Users/carlo/AppData/Local/Programs/Python/Python311/python.exe
# coding: latin-1
######## Programa para subir información de distribuidores manualmente a ConAgro ########
import os
import pandas as pd
import numpy as np
import json
import requests
import random
import tkinter as tk
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import sys
import time

# Definimos la clase Distribuidor, esta clase nos permite instanciar los nombres de las columnas de cada distribuidor y que se ajuste al LayOut correspondiente, siento la llave principal v_Num_Distri.
# Retomar la construcción de la clase Distribuidor.
## Crear los atributos:
##  Columna: Lista con el nombre de las columnas del archivo de excel para cada distribuidor.
##  Ruta: Ruta del archivo de excel que se va lipiar y subir al webservice.
##  Numero de distribuidor: Sold to del distribuidor.
##  Nombre de distribuidor: Nombre del distribuidor.

class Distribuidor:
    def __init__(self, v_Num_Distri, v_ruta, Columnas):
        self.ruta = v_ruta
        self.Num_Distri = v_Num_Distri
        self.Columnas = Columnas
    
    # Función para generar números aleatorios usados en Folio.
    def lista_aleatorios(self, n):
        return [str(random.randint(0, 99)) + str(random.randint(0, 99)) for _ in range(n)]
    
    # Método de limpieza para Ecuaquimica.
    def Ecuaqui_Clean(self, distribuidor, Columnas):
        Mes = {'Mes': list(range(1, 13)), 'Num': [str(i).zfill(2) for i in range(1, 13)]}
        df_Mes = pd.DataFrame(Mes)
        distribuidor['folio'] = ""
        while True:
            n = len(distribuidor[Columnas[7]])
            aleatorios = self.lista_aleatorios(n)
            np.random.seed(22)
            n = len(distribuidor[Columnas[0]])
            aleatorios1 = self.lista_aleatorios(n)
            Marcas1 = pd.DataFrame(distribuidor[Columnas[7]].unique().tolist(), columns=[Columnas[7]])
            Marcas1['ID'] = Marcas1.index
            distribuidor = distribuidor.merge(Marcas1, on=Columnas[7], how='left')
            distribuidor['Index'] = aleatorios1
            distribuidor['folio'] = distribuidor.apply(lambda row: ''.join(map(str, [row[Columnas[10]], row[Columnas[0]], row[Columnas[6]], row['ID'], row['Index']])), axis=1)
            if not distribuidor['folio'].duplicated().any():
                break
        # Sustitución de valores null por "" o por 0 para valores str y num respectivamente.
        distribuidor[Columnas[6]] = distribuidor[Columnas[6]].fillna(0)
        distribuidor[Columnas[11]] = distribuidor[Columnas[11]].fillna(0)
        distribuidor[Columnas[12]] = distribuidor[Columnas[12]].fillna(0)
        df_Mes = df_Mes.rename(columns={'Mes': 'MES'})
        distribuidor = distribuidor.merge(df_Mes, on=Columnas[0], how='left')
        distribuidor['fechaFactura'] = distribuidor[Columnas[10]].astype(str) + '-' + distribuidor['Num'].astype(str) + '-01'
        distribuidor['fechaFactura'] = distribuidor['fechaFactura'].astype(str)
        distribuidor['pais'] = 'Ecuador'
        distribuidor['marca'] = 'Syngenta'
        distribuidor = distribuidor.rename(columns={'fechaFactura': 'fecha_Facturacion', Columnas[12]: 'volumen_Facturado', Columnas[11]: 'valorT_Facturado', Columnas[5]: 'nombre_Cliente', Columnas[3]: 'nombre_Vendedor_Distribuidor', Columnas[7]: 'codeProduct_Distribuidor', Columnas[6]: 'rfc', Columnas[1]: 'localidad'})
        distribuidor['unidad_Medida'] = ""
        distribuidor['sucursal'] = ""
        distribuidor['linea_Producto'] = ""
        distribuidor['fecha_Facturacion'] = distribuidor['fecha_Facturacion'].astype(str)
        distribuidor['rfc'] = distribuidor['rfc'].astype(str)
        print(distribuidor.columns)
        columnas_deseadas = ['folio', 'fecha_Facturacion', 'volumen_Facturado', 'unidad_Medida', 'valorT_Facturado', 'nombre_Cliente', 'nombre_Vendedor_Distribuidor', 'codeProduct_Distribuidor', 'localidad', 'sucursal', 'linea_Producto', 'marca', 'pais', 'rfc']
        df = distribuidor[columnas_deseadas]
        # Verificación del dataframe final.
        print(df)
        return df
        
    def Agripa_clean(self, distribuidor, Columnas):
        Mes = {'Mes': list(range(1, 13)), 'Num': [str(i).zfill(2) for i in range(1, 13)]}
        df_Mes = pd.DataFrame(Mes)
        ## Verificacion para que solo se den folios distintos con un bucle infinito que para hasta que folio tenga solo valores distintos
        distribuidor['folio'] = ""
        while True:
            # print("Ingresa cuantos numeros aleatorios deseas obtener:")
            n = len(distribuidor[Columnas[10]])
            aleatorios = self.lista_aleatorios(n)
            # print(aleatorios)
            np.random.seed(22)
            n = len(distribuidor[Columnas[1]])
            aleatorios1 = self.lista_aleatorios(n)
            Marcas = pd.DataFrame(distribuidor[Columnas[9]].unique().tolist(), columns=[Columnas[9]])
            Marcas['ID'] = Marcas.index
            # Marcas
            distribuidor = distribuidor.merge(Marcas, on=Columnas[9], how='left')
            ### Creación de folio unico para subir a ConAgro
            distribuidor['Index'] = aleatorios
            distribuidor['folio'] = distribuidor.apply(lambda row: ''.join(map(str, [row[Columnas[0]], row[Columnas[1]], row[Columnas[5]], row['ID'], row['Index']])), axis=1)
            distribuidor[distribuidor['folio'].duplicated()]
            if distribuidor['folio'].duplicated().any()==False:
                break
        # Sustitución de valores null por "" o por 0 para valores str y num respectivamente.
        distribuidor[Columnas[7]] = distribuidor[Columnas[7]].fillna(0)
        distribuidor[Columnas[11]] = distribuidor[Columnas[11]].fillna(0)
        distribuidor[Columnas[12]] = distribuidor[Columnas[12]].fillna(0)
        #Verificacion por consola.
        print("Presencia de repetidos:", distribuidor['folio'].duplicated().any())
        ### Trasnformacion de datos de tipo fecha y el mes:
        distribuidor = distribuidor.merge(df_Mes, on=Columnas[1], how='left')
        ### Trasnformacion de datos de tipo fecha:
        distribuidor['fechaFactura'] = distribuidor[Columnas[0]].astype(str) + '-' + distribuidor['Num'].astype(str)+'-01'
        distribuidor['fechaFactura'] = distribuidor['fechaFactura'].astype(str)
        ## Datos distribuidor
        distribuidor['pais']='Ecuador'
        distribuidor['marca'] = 'Syngenta'
        #CAMBIO A LAYOUT DE CARGA PARA WS
        distribuidor = distribuidor.rename(columns={'fechaFactura': 'fecha_Facturacion',
                                                Columnas[11]: 'volumen_Facturado',
                                                Columnas[12]: 'valorT_Facturado',
                                                Columnas[6]: 'nombre_Cliente',
                                                Columnas[3]: 'nombre_Vendedor_Distribuidor',
                                                Columnas[9]: 'codeProduct_Distribuidor',
                                                Columnas[7]: 'rfc'})
        distribuidor['rfc']=distribuidor['rfc'].astype(str)
        distribuidor['unidad_Medida'] = ""
        distribuidor['sucursal'] = ""
        distribuidor['localidad']= ""
        distribuidor['sucursal'] = distribuidor['sucursal'].fillna("", inplace=False)
        distribuidor['unidad_Medida'] = distribuidor['unidad_Medida'].fillna("", inplace=False)
        # distribuidor = distribuidor.rename(columns={'Sucursal': 'sucursal'})
        distribuidor['linea_Producto'] = ""
        distribuidor['linea_Producto'] = distribuidor['linea_Producto'].fillna("", inplace=False)
        distribuidor['fecha_Facturacion'] = distribuidor['fecha_Facturacion'].astype(str)
        distribuidor['rfc']= distribuidor['rfc'].astype(str)
        # Comprobación del dataframe que se esta exportando.
        print(distribuidor.columns)

        columnas_deseadas = ['folio', 'fecha_Facturacion', 'volumen_Facturado', 'unidad_Medida', 'valorT_Facturado', 'nombre_Cliente', 'nombre_Vendedor_Distribuidor', 'codeProduct_Distribuidor', 'localidad', 'sucursal', 'linea_Producto', 'marca', 'pais', 'rfc']
        df = distribuidor[columnas_deseadas]
        # Verificación del dataframe final.
        print(df)
        return df
        # Aquí puedes agregar más atributos y métodos según tus necesidades

# Unimos ambas clases por composión, Esto es útil cuando quieres utilizar los métodos y atributos de Distribuidor dentro de App, pero App no es un tipo de Distribuidor.

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        # Aquí creas la instancia de la clase Distribuidor   
        self.geometry("700x200")
        self.resizable(0, 0)
        self.title("Sell Out to Web Service")
        
        # UI options
        paddings = {'padx': 5, 'pady': 5}
        entry_font = {'font': ('Helvetica', 11)}
        
        # configure the grid
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)
        
        # Creación de las variables globales para almacenar ruta y nombre de distribuidor.
        global v1, v2
        v1 = tk.StringVar()
        v2 = tk.StringVar()
        
        # Instanciamos el objeto Ecuaquimica de tipo distribuidor.
        self.Ecuaquimica = Distribuidor(v_Num_Distri='61610097', v_ruta=v1, 
                                        Columnas=[
                                                'MES', # Columnas[0]
                                                'UNIDAD NEGOCIO', # Columnas[1]
                                                'ZONA COMERCIAL', # Columnas[2]
                                                'VENDEDOR', # Columnas[3]
                                                'ESTABLECIMIENTO', # Columnas[4]
                                                'NEGOCIO ESTABLECIMIENTO', # Columnas[5]
                                                'IDENTIFICACION', # Columnas[6]
                                                'DESCRIPCION PRODUCTO', # Columnas[7]
                                                'NOMBRE PROVEEDOR', # Columnas[8]
                                                'TIPO CLIENTE', # Columnas[9]
                                                'año', # Columnas[10]
                                                'VENTA NETA', # Columnas[11]
                                                'KILOS LITROS TOTAL']) # Columnas[12]
        # Instanciamos el objeto Agripac de tipo distribuidor:
        self.Agripac = Distribuidor(v_Num_Distri='61610107', v_ruta=v1, 
                                        Columnas=['Año', # Columnas[0]
                                                'Mes', # Columnas[1]
                                                'Cod Rep Venta', # Columnas[2]
                                                'Rep Venta', # Columnas[3]
                                                'Division', # Columnas[4]
                                                'Cod Solicitante', # Columnas[5]
                                                'Solicitante', # Columnas[6]
                                                'RUC', # Columnas[7]
                                                'Tipo', # Columnas[8]
                                                'Cod Marca', # Columnas[9]
                                                'Marca', # Columnas[10]
                                                'Neto KG/LT', # Columnas[11]
                                                'Neto Venta']) # Columnas[12]
        # Ruta del archivo.
        ruta_label = ttk.Label(self, text="Ingresa la ruta del archivo:")
        ruta_label.grid(column=0, row=0, sticky=tk.W, **paddings)
        
        self.ruta_entry = ttk.Entry(self, validate="key", validatecommand=(self.validar_ruta_archivo, "%P"), textvariable=v1, **entry_font)
        self.ruta_entry.grid(row=0, column=2)
        
        # Número de distribuidor.
        distri_label = ttk.Label(self, text="Ingresa el Número de cliente:")
        distri_label.grid(column=0, row=1, sticky=tk.W, **paddings)
        
        self.distri_entry = ttk.Entry(self, validate="key", validatecommand=(self.validar_ruta_archivo, "%P"), textvariable=v2, **entry_font)
        self.distri_entry.grid(row=1, column=2)
        
        # Botones adicionales
        self.B1 = ttk.Button(self, text="Insertar", command=self.obtener_ruta)
        self.B1.grid(row=0, column=5)
        
        self.B3 = ttk.Button(self, text="Insertar", command=self.obtener_num)
        self.B3.grid(row=1, column=5)
        
        self.B2 = ttk.Button(self, text="Aceptar", command=lambda: self.principal())
        self.B2.place(relx=0.5, rely=0.6, anchor='center')
        
        # configure style
        self.style = ttk.Style(self)
        self.style.configure('TLabel', font=('Helvetica', 11))
        self.style.configure('TButton', font=('Helvetica', 11)) 
    # Método para cerrar la venta App.
    def cerrar_ventana(self):
        self.destroy()
    # Método que nos ayuda a validar la longitud de la ruta dada.
    def validar_ruta_archivo(self, nuevo_valor):
        return len(nuevo_valor) <= 200

    # Método que valida la longitud del número de distribuidor proporcionado.
    def validar_numero_cliente(self, new_text):
        return len(new_text) <=25
    # Ventana de error si el número de cliente ingresado no existe.
    def Cliente_Not_foud(self):
        messagebox.showerror("Error", "El cliente que estás ingresando no está registrado en ConAgro")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Ventana que muestra error si la ruta del archivo es incorrecta.
    def mostrar_error(self):
        messagebox.showerror("Error", "La ruta del archivo o el número de distribuidor no son válidos. Por favor vuelve a ejecutar el programa e ingresa los valores adecuador.")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Venta que se muestra si el archivo ingresado es de un tipo incorrecto o simplemente no existe.
    def error_Archivo(self):
        messagebox.showerror("Error", "Archivo no permitido, por favor vuelve a ejecutar el programa ingresando un archivo de tipo .xlsx")
        self.after(30000, self.cerrar_ventana)
        sys.exit()    
    # Ventana que se muestra si el layout de excel no es el correcto.
    def error_columnas(self):
        messagebox.showerror("Error", "El nombre de los encabezados no coincide, reportar al administrador.")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Funcion para mostrar error si no se creo el json correctamente:
    def error_json(self):
        messagebox.showerror("Error", "El Json no se creo de manera adecuada.")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Esta vwenta se muestra si el archivo Json no se cargo correctamente en el web service.
    def error_WebService(self, msj_web):
        messagebox.showerror("Error", msj_web)
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Funcion para mostrar que se creo el json correctamente:
    def msj_json(self,ruta_json):
        messagebox.showinfo("Info.", "Archivo Json %s creado con exito."% (ruta_json))  
    # Ventana que se muestra cuando el archivo Json se cargo correctamente en el Web Service.
    def msj_webservice(self, msj_web):
        messagebox.showinfo("Info", msj_web)
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Método que sirve para obtener la ruta de la entrada ruta_entry
    def obtener_ruta(self):
        global v1
        v1 = self.ruta_entry.get()
        self.B1.config(state=tk.DISABLED) 
    # Función para obtener número de distribuidor desde las entradas de texto
    def obtener_num(self):
        global v2
        v2 = self.distri_entry.get()  # Cambia 'distri_label' a 'distri_entry'
        self.B3.config(state=tk.DISABLED)
    # Función para consumir el webservice de ConAgro.
    def to_web_service(self, json_data):
        url = 'https://conagrosyngentapp.syngentadigitalapps.com/syngenta-service-0.0.1/api/v1/Sellout'
        headers = {'Content-type': 'application/json'}
        response = requests.post(url, data=json_data, headers=headers)
        if response.status_code == True:
            self.error_WebService(response.text)
        else:
            self.msj_webservice(response.text)
        return response.text
    # Función para convertir los datos de Sell Out en Json con el formato ConAgro.
    def json_conver(self,ruta, df, name_distri, num_distri, short):
        # Convertir DataFrame a JSON
        global json_data
        json_data = df.to_json(ruta + short+".json", orient="records")
        with open(ruta + short+".json", 'r') as archivo_json:
            data = json.load(archivo_json)
        json_str = json.dumps(data)
        json_headers = '{"clave_Distribuidor":"%s","nombre_distribuidor":"%s","productLines":' % (num_distri, name_distri)
        concatenado_data_json = json_headers + json_str + '}'
        with open(ruta + short.upper()+".json", 'w') as archivo_conagro:
            archivo_conagro.write(concatenado_data_json)
        nueva_ruta = ruta + short.upper()+".json"
        if os.path.exists(nueva_ruta):
            # Mensaje si existe el archivo
            self.msj_json(nueva_ruta)
            self.to_web_service(concatenado_data_json) # Comentar para pruebas.
        else:
            # Mensaje si no esite el archivo
            self.error_json()
        return concatenado_data_json
    ## Modulo principal, aquí empezamos 
    def principal(self):
        # Deshabilitamos el boton aceptar para que no se pueda usar más de una vez.
        self.B2.config(state=DISABLED)
        ## Almacenamos las variables globales para su posterior uso.
        ruta = v1
        Num_Distri = v2
        ## Verificacion:
        print("Ruta del archivo:", ruta)
        print("Numero de clientes:", Num_Distri)
        if self.validar_numero_cliente(Num_Distri) == False or self.validar_ruta_archivo(ruta) == False:
            self.mostrar_error()
        else:
            ## Inicia el manejo de excepciones:
            try:
                print("Inicio de procesamiento...")
                # Empezamos leyendo el archivo de Excel que se proporciona el RPA.
                distribuidor = pd.read_excel(ruta,sheet_name=0)
                print(distribuidor.columns)
                # Modulo Ecuaquimica.
                if self.Ecuaquimica.Num_Distri == Num_Distri:
                    print("Inicialización de transformación...")
                    Name_Distri = 'ECUATORIANA DE PROD. QUIM. S.A'
                    Name_short = 'Ecuaquimica'
                    df = self.Ecuaquimica.Ecuaqui_Clean(distribuidor, self.Ecuaquimica.Columnas)
                    
                    indice = ruta.rfind("\\")
                    nueva_ruta = ruta[:indice + 1]
                    
                    self.json_conver(ruta=nueva_ruta, df=df, name_distri=Name_Distri, num_distri=Num_Distri,short=Name_short)
                    
                    print("Procedimiento finalizado con exito...")
                    sys.exit()
                # Modulo para Agripac
                elif self.Agripac.Num_Distri == Num_Distri:
                    print("Inicialización de transformación...")
                    Name_Distri = 'AGRIPAC, S.A.'
                    Name_short = 'Agripac'
                    df = self.Ecuaquimica.Agripa_clean(distribuidor, self.Agripac.Columnas)
                    
                    indice = ruta.rfind("\\")
                    nueva_ruta = ruta[:indice + 1]
                    
                    self.json_conver(ruta=nueva_ruta, df=df, name_distri=Name_Distri, num_distri=Num_Distri, short=Name_short)
                    print("Procedimiento finalizado con exito...")
                    sys.exit()
                else:
                    self.Cliente_Not_foud()
            except FileNotFoundError:
                self.error_Archivo()
            except IndexError:
                self.error_columnas()
            except KeyError:
                self.error_columnas()


# Inicialización de la clase App        
if __name__ == "__main__":
    app = App()
    app.mainloop()