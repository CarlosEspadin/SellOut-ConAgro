#!/C:/Users/carlo/AppData/Local/Programs/Python/Python311/python.exe
# coding: latin-1
######## Programa para subir información de distribuidores manualmente a ConAgro SellOut y Stock ########
## Requerimientos
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
import traceback
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
    Mes = {'Mes': list(range(1, 13)), 'Num': [str(i).zfill(2) for i in range(1, 13)]}
    df_Mes = pd.DataFrame(Mes)
    
    def __init__(self, v_Num_Distri, v_ruta, Columnas):
        self.ruta = v_ruta
        self.Num_Distri = v_Num_Distri
        self.Columnas = Columnas
    
    # Función para generar números aleatorios usados en Folio.
    def lista_aleatorios(self, n):
        return [str(random.randint(0, 99)) + str(random.randint(0, 99)) for _ in range(n)]
    
    # Metodo para subir información de los demás distribuidores de México
    def Sell_Out(self, distribuidor, Columnas):
        # Condición para Ecuaquimica
        if 'MES' in distribuidor.columns and 'año' in distribuidor.columns:
            # Generación de folio sintentico unicos.
            distribuidor['folio'] = ""
            # Casteamos y redondeamos los datos númericos.
            distribuidor['MES'] = distribuidor['MES'].astype(int) # año
            distribuidor['año'] = distribuidor['año'].astype(int)  # mes
            # Sustitucion de valores null por "" o por 0 para valores str y num.
            distribuidor["rfc"] = distribuidor["rfc"].fillna(0)
            distribuidor["valorT_Facturado"] = distribuidor["valorT_Facturado"].fillna(0)
            distribuidor["volumen_Facturado"] = distribuidor["volumen_Facturado"].fillna(0)
            while True:
                n = len(distribuidor["codeProduct_Distribuidor"])
                aleatorios = self.lista_aleatorios(n)
                # np.random.seed(22)
                n = len(distribuidor["MES"])
                aleatorios1 = self.lista_aleatorios(n)
                Marcas1 = pd.DataFrame(distribuidor["codeProduct_Distribuidor"].unique().tolist(), columns=["codeProduct_Distribuidor"])
                Marcas1['ID'] = Marcas1.index
                distribuidor = distribuidor.merge(Marcas1, on="codeProduct_Distribuidor", how='left')
                print(distribuidor)
                distribuidor['Index'] = aleatorios1
                distribuidor['folio'] = distribuidor.apply(lambda row: ''.join(map(str, [row["año"], row["MES"], row["rfc"], row['ID'], row['Index']])), axis=1)
                if not distribuidor['folio'].duplicated().any():
                    break
                else:
                    distribuidor.drop(columns=["ID"], inplace=True)
            
            self.df_Mes = self.df_Mes.rename(columns={'Mes': 'MES'})
            distribuidor = distribuidor.merge(self.df_Mes, on='MES', how='left')
            distribuidor['fechaFactura'] = distribuidor["año"].astype(str) + '-' + distribuidor['Num'].astype(str) + '-01'
            distribuidor['fechaFactura'] = distribuidor['fechaFactura'].astype(str)
            distribuidor['pais'] = 'Ecuador'
            distribuidor['marca'] = 'Syngenta'
            distribuidor = distribuidor.rename(columns={'fechaFactura': 'fecha_Facturacion'})
            distribuidor['unidad_Medida'] = ""
            distribuidor['sucursal'] = ""
            distribuidor['linea_Producto'] = ""
            distribuidor['numero_Convenio'] = ""
            distribuidor['localidad']= ""
            distribuidor['fecha_Facturacion'] = distribuidor['fecha_Facturacion'].astype(str)
            distribuidor['rfc'] = distribuidor['rfc'].astype(str)
            columnas_deseadas = Columnas
        # Casp para Agripac
        elif 'Mes' in distribuidor.columns and 'Año' in distribuidor.columns:
            ## Verificacion para que solo se den folios distintos con un bucle infinito que para hasta que folio tenga solo valores distintos
            distribuidor['folio'] = ""
            # Sustitución de valores null por "" o por 0 para valores str y num respectivamente.
            distribuidor["rfc"] = distribuidor["rfc"].fillna(0)
            distribuidor["volumen_Facturado"] = distribuidor["volumen_Facturado"].fillna(0)
            distribuidor["valorT_Facturado"] = distribuidor["valorT_Facturado"].fillna(0)
            while True:
                # print("Ingresa cuantos numeros aleatorios deseas obtener:")
                n = len(distribuidor["Marca"])
                aleatorios = self.lista_aleatorios(n)
                # print(aleatorios)
                np.random.seed(22)
                n = len(distribuidor["Mes"])
                aleatorios1 = self.lista_aleatorios(n)
                Marcas = pd.DataFrame(distribuidor["codeProduct_Distribuidor"].unique().tolist(), columns=["codeProduct_Distribuidor"])
                Marcas['ID'] = Marcas.index
                # Marcas
                distribuidor = distribuidor.merge(Marcas, on="codeProduct_Distribuidor", how='left')
                ### Creación de folio unico para subir a ConAgro
                distribuidor['Index'] = aleatorios
                distribuidor['folio'] = distribuidor.apply(lambda row: ''.join(map(str, [row["Año"], row['Mes'], row["Cod Solicitante"], row['ID'], row['Index']])), axis=1)
                distribuidor[distribuidor['folio'].duplicated()]
                if distribuidor['folio'].duplicated().any()==False:
                    break
            #Verificacion por consola.
            print("Presencia de repetidos:", distribuidor['folio'].duplicated().any())
            ### Trasnformacion de datos de tipo fecha y el mes:
            distribuidor = distribuidor.merge(self.df_Mes, on='Mes', how='left')
            ### Trasnformacion de datos de tipo fecha:
            distribuidor['fechaFactura'] = distribuidor["Año"].astype(str) + '-' + distribuidor['Num'].astype(str)+'-01'
            distribuidor['fechaFactura'] = distribuidor['fechaFactura'].astype(str)
            ## Datos distribuidor
            distribuidor['pais']='Ecuador'
            distribuidor['marca'] = 'Syngenta'
            distribuidor = distribuidor.rename(columns={'fechaFactura': 'fecha_Facturacion'})
            distribuidor['unidad_Medida'] = ""
            distribuidor['sucursal'] = ""
            distribuidor['linea_Producto'] = ""
            distribuidor['numero_Convenio'] = ""
            distribuidor['localidad']= ""
            distribuidor['fecha_Facturacion'] = distribuidor['fecha_Facturacion'].astype(str)
            distribuidor['rfc'] = distribuidor['rfc'].astype(str)
            # Comprobación del dataframe que se esta exportando.
            columnas_deseadas = Columnas
            
        # Caso para todos los demás distribuidores provenientes de méxico
        else:
            columnas_deseadas = Columnas
            distribuidor['fecha_Facturacion'] = distribuidor['fecha_Facturacion'].sort_values().apply(lambda x: x.strftime("%Y-%m-%d"))
        
        df = distribuidor[columnas_deseadas]
        
        print(distribuidor.columns)
        print(df)
        return df

## Ventana introducción

class VentanaIntroduccion(tk.Toplevel):
    def habilitar(self):
        self.boton_continuar.config(state=tk.ACTIVE)
    
    def resource_path(self,relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            self.base_path = sys._MEIPASS
        except Exception:
            self.base_path = os.path.abspath(".")
        return os.path.join(self.base_path, relative_path)
    
    def __init__(self, parent, app):
        self.ruta_ico_s = self.resource_path("ConAgro_icon_small.png")
        self.ruta_ico_b = self.resource_path("ConAgro_icon_big.png")
        super().__init__(parent)
        self.title("Introducción a la Aplicación")
        self.app = app  # Referencia a la aplicación principal
        self.configure(bg="white", relief="sunken", highlightcolor="green")
        # configure icon
        self.icon_big = tk.PhotoImage(file=self.ruta_ico_b)
        self.icon_small = tk.PhotoImage(file=self.ruta_ico_s)
        self.iconphoto(False, self.icon_big, self.icon_small)
        
        # Agregar texto explicativo
        label = ttk.Label(self, text="Bienvenido a la Aplicación para tratamiento y carga de información al webservice ConAgro ")
        label.pack(padx=10, pady=5)
        
        # Selección de tipo de carga.
        Label_radio = ttk.Label(self, text="Seleciona el tipo de carga:")
        Label_radio.pack(padx=10, pady=2)
        
        # Botones de radio
        self.selected = tk.StringVar()
        
        self.r1 = ttk.Radiobutton(self, text='SellOut', value='SellOut', variable=self.selected, command=self.habilitar)
        self.r1.pack(padx=10, pady=5)
        
        self.r2 = ttk.Radiobutton(self, text='Stock', value='Stock', variable=self.selected, command=self.habilitar)
        self.r2.pack(padx=10, pady=5)

        # Botón para continuar
        self.boton_continuar = ttk.Button(self, text="Continuar", command=self.continuar)
        self.boton_continuar.pack(pady=5)
        self.boton_continuar.config(state=tk.DISABLED)
        
        
    def continuar(self):
        self.destroy()  # Cerrar esta ventana de introducción
        self.app.iniciar_app()  # Llamar al método iniciar_app de la aplicación principal


# Unimos ambas clases por composión, Esto es útil cuando quieres utilizar los métodos y atributos de Distribuidor dentro de App, pero App no es un tipo de Distribuidor.

class App(tk.Tk):
    def resource_path(self,relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            self.base_path = sys._MEIPASS
        except Exception:
            self.base_path = os.path.abspath(".")

        return os.path.join(self.base_path, relative_path)
    
    def __init__(self):
        super().__init__()
        self.title("Sell Out to Web Service")
        self.geometry("700x200")
        self.resizable(0, 0)
        
        # Crear las variables de instancia
        self.ruta = tk.StringVar(self)
        self.num_distri = tk.StringVar(self)
        self.tipo_proceso = ""
        
        # Ocultar la ventana principal hasta que se muestre la ventana de introducción
        self.withdraw()
        
        # Crear la ventana de introducción y pasar la referencia a la aplicación principal
        self.introduccion = VentanaIntroduccion(self, self)
        
    def iniciar_app(self):
        # Mostrar la ventana principal de la aplicación
        self.deiconify()
        
        # Extraemos el valor para el tipo de proceso.
        Tipo = self.introduccion.selected.get()
        # Titulo de la venta principal:
        self.title(Tipo +" to Web Service")
        
        # UI options
        paddings = {'padx': 6, 'pady': 6}
        entry_font = {'font': ('Helvetica', 11)}
        
        # configure the grid
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)
        
        # Ruta del archivo.
        ruta_label = ttk.Label(self, text="Ingresa la ruta del archivo:")
        ruta_label.grid(column=0, row=1, sticky=tk.W, **paddings)
        
        self.ruta_entry = ttk.Entry(self, validate="key", validatecommand=(self.validar_ruta_archivo, "%P"), textvariable=self.ruta, **entry_font)
        self.ruta_entry.grid(row=1, column=2)
        self.ruta_entry.config(width=40)
        
        # Número de distribuidor.
        distri_label = ttk.Label(self, text="Ingresa el Número de cliente:")
        distri_label.grid(column=0, row=2, sticky=tk.W, **paddings)
        
        self.distri_entry = ttk.Entry(self, validate="key", validatecommand=(self.validar_numero_cliente, "%P"), textvariable=self.num_distri, **entry_font)
        self.distri_entry.grid(row=2, column=2)
        self.distri_entry.config(width=40)
        
        # Botones adicionales
        self.B1 = ttk.Button(self, text="Insertar", command=self.obtener_ruta)
        self.B1.grid(row=1, column=5)
        
        self.B3 = ttk.Button(self, text="Insertar", command=self.obtener_num)
        self.B3.grid(row=2, column=5)
        
        self.B2 = ttk.Button(self, text="Aceptar", command=lambda: self.principal())
        self.B2.place(relx=0.5, rely=0.7, anchor='center')
        
        # configure style
        self.style = ttk.Style(self)
        self.style.configure('TLabel', font=('Helvetica', 11))
        self.style.configure('TButton', font=('Helvetica', 11)) 
        
        # Establecemos la ruta absoluta:
        self.ruta_ico_s = self.resource_path("ConAgro_icon_small.png")
        self.ruta_ico_b = self.resource_path("ConAgro_icon_big.png")
        
        # configure icon
        self.icon_big = tk.PhotoImage(file=self.ruta_ico_b)
        self.icon_small = tk.PhotoImage(file=self.ruta_ico_s)
        self.iconphoto(False, self.icon_big, self.icon_small)
    
    # Método para cerrar la venta App.
    def cerrar_ventana(self):
        self.destroy()
    # Método que nos ayuda a validar la longitud de la ruta dada.
    def validar_ruta_archivo(self, nuevo_valor):
        try:
            return len(nuevo_valor) <= 200
        except TypeError as e4:
                self.mostrar_error()
                

    # Método que valida la longitud del número de distribuidor proporcionado.
    def validar_numero_cliente(self, new_text):
        try:
            return len(new_text) <=25
        except TypeError as e4:
            self.mostrar_error()

    # Ventana de error si el número de cliente ingresado no existe.
    def Cliente_Not_foud(self):  # sourcery skip: class-extract-method
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
        # sys.exit()   
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
        self.B1.config(state=tk.DISABLED) 
    # Función para obtener número de distribuidor desde las entradas de texto
    def obtener_num(self):
        self.B3.config(state=tk.DISABLED)
    # Función para consumir el webservice de ConAgro.
    def to_web_service(self, json_data):
        sess = requests.Session()
        adapter = requests.adapters.HTTPAdapter(max_retries = 20)
        sess.mount('https://', adapter)
        url = 'https://conagrosyngentapp.syngentadigitalapps.com/syngenta-service-0.0.1/api/v1/Sellout'
        headers = {'Content-type': 'application/json'}
        json_data
        print("Este es el encavezado: ", headers)
        # sess = sess.post(url, data=json_data, headers=headers, verify=False)
        sess = sess.post(url, data=json_data, headers=headers)
        if sess.status_code == True:
            self.error_WebService(sess.text)
            print("Este es el encavezado: ", sess.text)
        else:
            self.msj_webservice(sess.text)
            print("Este es el encavezado: ", sess.text)
        return sess.text
    # Función para convertir los datos de Sell Out en Json con el formato ConAgro. 
    def json_conver(self,ruta, df, name_distri, num_distri):
        # Convertir DataFrame a JSON
        global json_data
        print("Validación de ruta: ", ruta + ".json")
        json_data = df.to_json(ruta + ".json", orient="records")
        with open(ruta +".json", 'r') as archivo_json:
            data = json.load(archivo_json)
        json_str = json.dumps(data)
        json_headers = '{"clave_Distribuidor":"%s","nombre_distribuidor":"%s","productLines":' % (num_distri, name_distri)
        
        concatenado_data_json = json_headers + json_str + '}'
        with open(ruta +".json", 'w') as archivo_conagro:
            archivo_conagro.write(concatenado_data_json)
        nueva_ruta = ruta +".json"
        if os.path.exists(nueva_ruta):
            # Mensaje si existe el archivo
            self.msj_json(nueva_ruta)
            self.to_web_service(concatenado_data_json) # Comentar para pruebas.
        else:
            # Mensaje si no esite el archivo
            self.error_json()
        return concatenado_data_json
    ## Modulo principal, que manda a llamar los objetos instanciados de la clase Distribuidor y hace validaciones para errores.
    
                
    def principal(self):  # sourcery skip: extract-method
        self.iniciar_app()# sourcery skip: extract-duplicate-method, extract-methodate-method, extract-method
        # Deshabilitamos el boton aceptar para que no se pueda usar más de una vez.
        self.B2.config(state=tk.DISABLED)
        ## Almacenamos las variables globales para su posterior uso.
        # Extraemos el valor para el tipo de proceso.
        v_Tipo = self.introduccion.selected.get()
        print("El proceso sera de tipo: ", v_Tipo)

        v1 = self.ruta.get()
        v2 = self.num_distri.get()  # Cambia 'distri_label' a 'distri_entry'
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
                if v_Tipo == 'SellOut':
                # Modulo General.
                    print("Validación 1")
                    self.Distribuidores = Distribuidor(v_Num_Distri=Num_Distri, v_ruta=self.ruta,  Columnas=[
                                                'folio',
                                                'fecha_Facturacion',
                                                'volumen_Facturado',
                                                'unidad_Medida',
                                                'valorT_Facturado',
                                                'rfc',
                                                'nombre_Cliente',
                                                'codeProduct_Distribuidor',
                                                'localidad',
                                                'sucursal',
                                                'nombre_Vendedor_Distribuidor',
                                                'linea_Producto',
                                                'marca',
                                                'pais',
                                                'numero_Convenio'])
                    
                    if self.Distribuidores.Num_Distri == Num_Distri:
                        print("Inicialización de transformación...")
                        # Obtener el Nombre de distribuidor.
                        self.ruta_base = self.resource_path("Base clientes.xlsx")
                        self.soldTo = pd.read_excel(self.ruta_base, sheet_name=0)
                        print("Tipo de dato de Número de distribuidores: ", Num_Distri)
                        Name_Distri_aux=self.soldTo[self.soldTo["Num_Distri"]==int(Num_Distri)]
                        Name_Distri = list(Name_Distri_aux["Name_Distri"])
                        Name_Distri=Name_Distri[0]
            
                        print("Nombre del distribuidor: ", Name_Distri)
                            
                        if 'folio' not in distribuidor.columns:
                            df = self.Distribuidores.Sell_Out(distribuidor, self.Distribuidores.Columnas)
                            
                            indice = ruta.rfind(".xlsx")
                            nueva_ruta = ruta[:indice]
                            print(nueva_ruta)
                            
                            self.json_conver(ruta=nueva_ruta, df=df, name_distri=Name_Distri, num_distri=Num_Distri)
                            
                            print("Procedimiento finalizado con exito...")
                            sys.exit()
                        elif distribuidor[self.Distribuidores.Columnas[0]].isnull().any() or distribuidor[self.Distribuidores.Columnas[10]].isnull().any():
                            self.mostrar_error()
                        else:
                            df = self.Distribuidores.Sell_Out(distribuidor, self.Distribuidores.Columnas)
                            
                            indice = ruta.rfind(".xlsx")
                            nueva_ruta = ruta[:indice]
                            print(nueva_ruta)
                            
                            self.json_conver(ruta=nueva_ruta, df=df, name_distri=Name_Distri, num_distri=Num_Distri)
                            
                            print("Procedimiento finalizado con exito...")
                            sys.exit()
                    else:
                        self.Cliente_Not_foud()
                else: # Proceso para Inventarios
                    print("Test...")
                    sys.exit()
            except FileNotFoundError as e1:
                traceback.print_exc()
                self.error_Archivo()
            except IndexError as e2:
                traceback.print_exc()
                self.error_columnas()
            except KeyError as e3:
                traceback.print_exc()
                self.error_columnas()
            except TypeError as e4:
                self.mostrar_error()


# Inicialización de la clase App        
if __name__ == "__main__":
    
    app = App()
    app.mainloop()