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
from tkinter import font
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
    
    def __init__(self, v_Num_Distri, v_ruta, Columnas, Columnas_Stk):
        self.ruta = v_ruta
        self.Num_Distri = v_Num_Distri
        self.Columnas = Columnas
        self.Columnas_Stk = Columnas_Stk
    
    # Función para generar números aleatorios usados en Folio.
    def lista_aleatorios(self, n):
        return [str(random.randint(0, 99)) + str(random.randint(0, 99)) for _ in range(n)]
    
    
    # Metodo para subir información de los demás distribuidores de México
    def Sell_Out(self, distribuidor, Columnas):
        # Condición para Ecuaquimica
        if int(distribuidor['clave_Distribuidor'][0])==61610097:
            # Generación de folio sintentico unicos.
            distribuidor['folio'] = ""
            # Casteamos y redondeamos los datos númericos.
            distribuidor[['Año', 'Mes', 'Dia']] = distribuidor['fecha_Facturacion'].str.split('-', expand=True)
            distribuidor['Mes'] = distribuidor['Mes'].astype(int) # año
            distribuidor['Año'] = distribuidor['Año'].astype(int)  # mes
            # Sustitucion de valores null por "" o por 0 para valores str y num.
            distribuidor["rfc"] = distribuidor["rfc"].fillna(0)
            distribuidor["valorT_Facturado"] = distribuidor["valorT_Facturado"].fillna(0)
            distribuidor["volumen_Facturado"] = distribuidor["volumen_Facturado"].fillna(0)
            while True:
                n = len(distribuidor["codeProduct_Distribuidor"])
                aleatorios = self.lista_aleatorios(n)
                # np.random.seed(22)
                n = len(distribuidor["Mes"])
                aleatorios1 = self.lista_aleatorios(n)
                Marcas1 = pd.DataFrame(distribuidor["codeProduct_Distribuidor"].unique().tolist(), columns=["codeProduct_Distribuidor"])
                Marcas1['ID'] = Marcas1.index
                distribuidor = distribuidor.merge(Marcas1, on="codeProduct_Distribuidor", how='left')
                print(distribuidor)
                distribuidor['Index'] = aleatorios1
                distribuidor['folio'] = distribuidor.apply(lambda row: ''.join(map(str, [row["Año"], row["Mes"], row["rfc"], row['ID'], row['Index']])), axis=1)
                if not distribuidor['folio'].duplicated().any():
                    break
                else:
                    distribuidor.drop(columns=["ID"], inplace=True)
            
            # # self.df_Mes = self.df_Mes.rename(columns={'Mes': 'MES'})
            # distribuidor = distribuidor.merge(self.df_Mes, on='MES', how='left')
            # distribuidor['fechaFactura'] = distribuidor["año"].astype(str) + '-' + distribuidor['Num'].astype(str) + '-01'
            # distribuidor['fechaFactura'] = distribuidor['fechaFactura'].astype(str)
            distribuidor['pais'] = 'Ecuador'
            distribuidor['marca'] = 'Syngenta_xlsx'
            # distribuidor = distribuidor.rename(columns={'fechaFactura': 'fecha_Facturacion'})
            distribuidor['unidad_Medida'] = ""
            distribuidor['sucursal'] = ""
            distribuidor['linea_Producto'] = ""
            distribuidor['numero_Convenio'] = ""
            distribuidor['localidad']= ""
            distribuidor['fecha_Facturacion'] = distribuidor['fecha_Facturacion'].astype(str)
            distribuidor['rfc'] = distribuidor['rfc'].astype(str)
            columnas_deseadas = Columnas
        # Casp para Agripac
        elif int(distribuidor['clave_Distribuidor'][0])==61610107:
            ## Verificacion para que solo se den folios distintos con un bucle infinito que para hasta que folio tenga solo valores distintos
            distribuidor['folio'] = ""
            distribuidor[['Año', 'Mes', 'Dia']] = distribuidor['fecha_Facturacion'].str.split('-', expand=True)
            distribuidor['Mes'] = distribuidor['Mes'].astype(int) # año
            distribuidor['Año'] = distribuidor['Año'].astype(int)  # mes
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
            # distribuidor = distribuidor.merge(self.df_Mes, on='Mes', how='left')
            ### Trasnformacion de datos de tipo fecha:
            # distribuidor['fechaFactura'] = distribuidor["Año"].astype(str) + '-' + distribuidor['Num'].astype(str)+'-01'
            # distribuidor['fecha_Facturacion'] = distribuidor['fecha_Facturacion'].astype(str)
            ## Datos distribuidor
            distribuidor['pais']='Ecuador'
            distribuidor['marca'] = 'Syngenta_xlsx'
            # distribuidor = distribuidor.rename(columns={'fechaFactura': 'fecha_Facturacion'})
            distribuidor['unidad_Medida'] = ""
            distribuidor['sucursal'] = ""
            distribuidor['linea_Producto'] = ""
            distribuidor['numero_Convenio'] = ""
            distribuidor['localidad']= ""
            distribuidor['fecha_Facturacion'] = pd.to_datetime(distribuidor['fecha_Facturacion'], format="%Y-%m-%d")
            distribuidor['fecha_Facturacion'] = distribuidor['fecha_Facturacion'].sort_values().apply(lambda x: x.strftime("%Y-%m-%d"))
            distribuidor['rfc'] = distribuidor['rfc'].astype(str)
            # Comprobación del dataframe que se esta exportando.
            columnas_deseadas = Columnas
            
        # Caso para todos los demás distribuidores provenientes de méxico
        else:
            columnas_deseadas = Columnas
            distribuidor['fecha_Facturacion'] = pd.to_datetime(distribuidor['fecha_Facturacion'], format="%Y-%m-%d")
            distribuidor['fecha_Facturacion'] = distribuidor['fecha_Facturacion'].sort_values().apply(lambda x: x.strftime("%Y-%m-%d"))
        
        distribuidor['marca'] = 'Syngenta_xlsx'
        df = distribuidor[columnas_deseadas]
        
        print(distribuidor.columns)
        print(df)
        return df

    def Stocks(self, distribuidor, Columnas):
            columnas_deseadas = Columnas
            distribuidor['fecha_Inventario'] = distribuidor['fecha_Inventario'].sort_values().apply(lambda x: x.strftime("%Y-%m-%d"))
        
            df = distribuidor[columnas_deseadas]
            
            print(distribuidor.columns)
            print(df)
            return df

# Clase App Iterface visual
class App(tk.Tk):
    def resource_path(self,relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            self.base_path = sys._MEIPASS
        except Exception:
            self.base_path = os.path.abspath(".")

        return os.path.join(self.base_path, relative_path)

    def create_intro_label(self):
        intro_text = (
            "Bienvenido a la Aplicación para tratamiento y carga de información al webservice ConAgro\n\n"
            "Esta herramienta te permitirá cargar información de \n"
            "SellOut e Inventarios. Por favor, sigue las instrucciones \n"
            "para cargar y procesar los datos.\n\n"
            "Para comenzar, selecciona el tipo de proceso en el menú lateral\n"
            "y selecciona el archivo a cargar en el webservice, luego proporciona la \n"
            "información requerida en los campos correspondientes."
        )
        self.intro_label = ttk.Label(self, text=intro_text, style='TLabel', justify=tk.LEFT)
    
    def create_all_widgets(self):
        # UI options
        paddings = {'padx': 6, 'pady': 6}
        entry_font = {'font': ('Helvetica', 11)}
        # configure the grid
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)
        
        self.ruta_label = ttk.Label(self, text="Ingresa la ruta del archivo:")
        
        # Ruta del archivo.
        self.ruta_label = ttk.Label(self, text="Ingresa la ruta del archivo:")
        
        self.ruta_entry = ttk.Entry(self, validate="key", validatecommand=(self.validar_ruta_archivo, "%P"), textvariable=self.ruta, **entry_font)
        self.ruta_entry.config(width=45)
        
        # Número de distribuidor.
        self.distri_label = ttk.Label(self, text="Ingresa el Número de cliente:")
        
        self.distri_entry = ttk.Entry(self, validate="key", validatecommand=(self.validar_numero_cliente, "%P"), textvariable=self.num_distri, **entry_font)
        self.distri_entry.config(width=45)
        
        # Botones adicionales
        self.B1 = ttk.Button(self, text="Insertar", command=self.obtener_ruta)
        
        self.B3 = ttk.Button(self, text="Insertar", command=self.obtener_num)
        
        self.B2 = ttk.Button(self, text="Aceptar", command=lambda: self.principal())
        
        self.tipo_label = ttk.Label(self, text="Proceso de carga al webservice.")
        # self.B2.place(relx=0.5, rely=0.7, anchor='center')
        
    def show_intro_label(self):
        self.intro_label.place(relx=0.2, rely=0.1, relwidth=0.8, relheight=0.3)
    
    def hide_intro_labe(self):
        self.intro_label.place_forget()
        
    def show_all_widgets(self):
        self.hide_intro_labe()
        self.nx = 0.1
        # Tipo
        self.tipo_label.place(relx=0.3-self.nx, rely=0.15)
        # Ruta.
        self.ruta_label.place(relx=0.3-self.nx, rely=0.25)
        self.ruta_entry.place(relx=0.3-self.nx, rely=0.3)
        self.B1.place(relx=0.7-self.nx, rely=0.3)
        # Numero de distribuidor.
        self.distri_label.place(relx=0.3-self.nx, rely=0.35)
        self.distri_entry.place(relx=0.3-self.nx, rely=0.4)
        self.B3.place(relx=0.7-self.nx, rely=0.4)
        
        self.B2.place(relx=0.5, rely=0.7, anchor='center')
        
    def hide_all_widgets(self):
        self.tipo_label.place_forget()
        self.ruta_label.place_forget()
        self.ruta_entry.place_forget()
        self.B2.place_forget()
        self.B1.place_forget()
        self.B3.place_forget()
        self.distri_label.place_forget()
        self.distri_entry.place_forget()
        
    def paneles(self):        
        # Configuración de barra superior
        self.barra_superior = tk.Frame(
            self, bg='#5f7800', height=40
        )
        
        self.barra_superior.pack(side=tk.TOP, fill='both')   
        
        # Etiqueta de título
        self.labelTitulo = tk.Label(self.barra_superior, text="Aplicación para consumo del webservice ConAgro")
        self.labelTitulo.config(fg="#fff", font=(
            "Roboto", 15), bg="#5f7800", pady=10, width=46)
        self.labelTitulo.pack(side=tk.TOP)  

        self.menu_lateral = tk.Frame(self, bg="#abb400", width=150)
        self.menu_lateral.pack(side=tk.LEFT, fill='both', expand=False)
        
    def controles_barra_superior(self):
        # Configuración de la barra superior
        font_awesome = font.Font(family='FontAwesome', size=12)

        # Establecemos la ruta absoluta del icono del boton
        self.ruta_menu = self.resource_path("menu.png")
        
        # configure icon
        self.icon_menu = tk.PhotoImage(file=self.ruta_menu)
        
        # self.nx = float(self.selected.get())-0.8 if self.menu_lateral.winfo_ismapped() else float(self.selected.get())-0.8
        # Botón del menú lateral
        self.buttonMenuLateral = tk.Label(self.barra_superior, text="Menú", font=("Roboto", 15),
                                bd=0, bg="#5f7800", fg="#fff", image=self.icon_menu, compound="right", padx=10)
        self.buttonMenuLateral.place(relx=0.02, rely=0.25)

    
    def controles_menu_lateral(self):
        # Configuración del menú lateral
        ancho_menu = 20
        alto_menu = 2
        font_awesome = font.Font(family='FontAwesome', size=15)
        
        # Etiqueta de perfil
        self.labelPerfil = tk.Label(self.menu_lateral, image=self.ico_Home, bg="#abb400")
        self.labelPerfil.pack(side=tk.TOP, pady=10)

        
        # Crear los botones con icono y texto
        self.buttonDashBoard = self.create_menu_button(self.menu_lateral, "Inicio", self.ico_Step1, self.create_greeting_message, "2")
        self.bottonSync = self.create_menu_button(self.menu_lateral, "SellOut", self.ico_Step2, self.create_greeting_message, "0")
        self.bottonExt = self.create_menu_button(self.menu_lateral, "Inventarios", self.ico_Step3, self.create_greeting_message, "1")
        # self.bottonCL = self.create_menu_button(self.menu_lateral, "Chile", self.ico_Step4, self.create_greeting_message, "3")
        self.buttonSettings = self.create_menu_button(self.menu_lateral, "Settings", self.icoSettings, self.create_greeting_message)
        
    def create_menu_button(self, parent, text, icon, command, value=None):
        frame = tk.Frame(parent, bg="#abb400")
        frame.pack(side=tk.TOP, fill=tk.X)

        label_icon = tk.Label(frame, image=icon, font=("FontAwesome", 15), bg="#abb400", fg="white")
        label_icon.pack(side=tk.LEFT, padx=5)

        label_text = tk.Label(frame, text=text, font=("Arial", 12), bg="#abb400", fg="white")
        label_text.pack(side=tk.LEFT, padx=5)

        if value is not None:
            label_icon.bind("<Button-1>", lambda event: self.on_click(value, command))
            label_text.bind("<Button-1>", lambda event: self.on_click(value, command))
        else:
            label_icon.bind("<Button-1>", lambda event: command())
            label_text.bind("<Button-1>", lambda event: command())

        self.bind_hover_events(frame, label_icon, label_text)

        return frame
    
    def bind_hover_events(self, frame, label_icon, label_text):
        # Asociar eventos Enter y Leave con la función dinámica
        frame.bind("<Enter>", lambda event: self.on_enter(event, frame, label_icon, label_text))
        frame.bind("<Leave>", lambda event: self.on_leave(event, frame, label_icon, label_text))
        label_icon.bind("<Enter>", lambda event: self.on_enter(event, frame, label_icon, label_text))
        label_icon.bind("<Leave>", lambda event: self.on_leave(event, frame, label_icon, label_text))
        label_text.bind("<Enter>", lambda event: self.on_enter(event, frame, label_icon, label_text))
        label_text.bind("<Leave>", lambda event: self.on_leave(event, frame, label_icon, label_text))

    def on_enter(self, event, frame, label_icon, label_text):
        # Cambiar estilo al pasar el ratón por encima
        frame.config(bg="#009933")
        label_icon.config(bg="#009933")
        label_text.config(bg="#009933")

    def on_leave(self, event, frame, label_icon, label_text):
        # Restaurar estilo al salir el ratón
        frame.config(bg="#abb400")
        label_icon.config(bg="#abb400")
        label_text.config(bg="#abb400")

    def on_click(self, value, command):
        # Obtener el valor de la variable de tipo de ventana
        self.selected.set(value)
        command()
    
    def __init__(self):
        super().__init__()
        self.title("Sell Out to Web Service")
        self.geometry("950x600")
        self.resizable(0, 0)
        
        # Crear las variables de instancia
        self.ruta = tk.StringVar(self)
        self.num_distri = tk.StringVar(self)
        self.tipo_proceso = ""
        self.selected = tk.StringVar()
        
        self.configure(bg='#f0f0f0')       
        # Ocultar la ventana principal hasta que se muestre la ventana de introducción
        # self.withdraw()
        
        # Crear la ventana de introducción y pasar la referencia a la aplicación principal
        # self.introduccion = VentanaIntroduccion(self, self)
        self.iniciar_app()
    
    def create_greeting_message(self):
        # Obtener el valor de Tipo
        # Tipo de
        
        
        Tipo = self.selected.get()
        # Tipo de 
        if Tipo == "0":
            v_Tipo = "SellOut"
        elif Tipo == "1":
            v_Tipo = "Stock   "
        else:
            v_Tipo = ""
                
        if Tipo == "2":
            # self.tipo_label.place_forget()
            self.hide_all_widgets()
            self.show_intro_label()
        else:
            # Tipo
            # self.tipo_label.place(relx=0.3, rely=0.15)
            self.show_all_widgets()
        
    def iniciar_app(self):
        
        
                # Establecemos la ruta absoluta:
        self.ruta_ico_s = self.resource_path("ConAgro_icon_small.png")
        self.ruta_ico_b = self.resource_path("ConAgro_icon_big.png")
        self.ruta_icoHome = self.resource_path("agriculture_6739552.png")
        self.ruta_icoStep1 = self.resource_path("home_2115185.png")
        self.ruta_icoStep2 = self.resource_path("workflow_14254834.png")
        self.ruta_icoStep3 = self.resource_path("onboarding_14753055.png")
        self.ruta_icoStep4 = self.resource_path("review-document_14752610.png")
        self.ruta_icoSettings = self.resource_path("setting_1146744.png")
        
        
        # configure icon
        self.icon_big = tk.PhotoImage(file=self.ruta_ico_b)
        self.icon_small = tk.PhotoImage(file=self.ruta_ico_s)
        self.ico_Home = tk.PhotoImage(file=self.ruta_icoHome)
        self.ico_Step1 = tk.PhotoImage(file=self.ruta_icoStep1)
        self.ico_Step2 = tk.PhotoImage(file=self.ruta_icoStep2)
        self.ico_Step3 = tk.PhotoImage(file=self.ruta_icoStep3)
        self.ico_Step4 = tk.PhotoImage(file=self.ruta_icoStep4)
        self.icoSettings = tk.PhotoImage(file=self.ruta_icoSettings)
        
        self.iconphoto(False, self.icon_big, self.icon_small, self.ico_Home, self.ico_Step1, self.ico_Step2, self.ico_Step3, self.icoSettings)
        
        # Mostrar la ventana principal de la aplicación
        self.deiconify()
        # Invocacion de paneles
        self.paneles()
        self.controles_barra_superior()
        self.controles_menu_lateral() 
        
        self.create_intro_label()
        self.show_intro_label()
        self.create_all_widgets()
        
        # self.create_widgets()
        # Extraemos el valor para el tipo de proceso.
        # Tipo = self.introduccion.selected.get()
        # Titulo de la venta principal:
        # self.title(Tipo +" to Web Service")
        
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
        self.withdraw()
        messagebox.showerror("Error", "El cliente que estás ingresando no está registrado en ConAgro")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Ventana que muestra error si la ruta del archivo es incorrecta.
    def mostrar_error(self):
        self.withdraw()
        messagebox.showerror("Error", "La ruta del archivo o el número de distribuidor no son válidos. Por favor vuelve a ejecutar el programa e ingresa los valores adecuador.")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Venta que se muestra si el archivo ingresado es de un tipo incorrecto o simplemente no existe.
    def error_Archivo(self):
        self.withdraw()
        messagebox.showerror("Error", "Archivo no permitido, por favor vuelve a ejecutar el programa ingresando un archivo de tipo .xlsx")
        self.after(30000, self.cerrar_ventana)
        sys.exit()   
    # Ventana que se muestra si el layout de excel no es el correcto.
    def error_columnas(self):
        self.withdraw()
        messagebox.showerror("Error", "El nombre de los encabezados no coincide, reportar al administrador.")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Ventana que se muestra si el layout de excel no es el correcto.
    def error_fechas(self):
        self.withdraw()
        messagebox.showerror("Error", "Registros vacíos en la columna de fecha, informar a los administradores.")
        self.after(30000, self.cerrar_ventana)
        sys.exit()  
    
    # Funcion para mostrar error si no se creo el json correctamente:
    def error_json(self):
        self.withdraw()
        messagebox.showerror("Error", "El Json no se creo de manera adecuada.")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Esta vwenta se muestra si el archivo Json no se cargo correctamente en el web service.
    def error_WebService(self, msj_web):
        self.withdraw()
        messagebox.showerror("Error", msj_web)
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Funcion para mostrar que se creo el json correctamente:
    def msj_json(self,ruta_json):
        # sourcery skip: hoist-similar-statement-from-if
        # sourcery skip: hoist-similar-statement-from-if
        messagebox.showinfo("Info.", "Archivo Json %s creado con exito."% (ruta_json))
        self.hide_all_widgets()
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
        # Tipo de 
        if self.selected.get() == "0":
            url = 'https://conagrosyngentapp.syngentadigitalapps.com/syngenta-service-0.0.1/api/v1/Sellout'
        elif self.selected.get() == "1":
            url = 'https://conagrosyngentapp.syngentadigitalapps.com/syngenta-service-0.0.1/api/v1/inventario'
        else:
            url = ""
        
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
        if self.selected.get() == "0":
            json_headers = '{"clave_Distribuidor":"%s","nombre_distribuidor":"%s","productLines":' % (num_distri, name_distri)
        elif self.selected.get() == "1":
            json_headers = '{"clave_Distribuidor":"%s","nombre_Distribuidor":"%s","productLines":' % (num_distri, name_distri)
        else:
            json_headers = ""
        
        
        concatenado_data_json = json_headers + json_str + '}'
        with open(ruta +".json", 'w') as archivo_conagro:
            archivo_conagro.write(concatenado_data_json)
        nueva_ruta = ruta +".json"
        if os.path.exists(nueva_ruta):
            # Mensaje si existe el archivo
            self.withdraw()
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
        if self.selected.get() == "0":
            v_Tipo = "SellOut"
        elif self.selected.get() == "1":
            v_Tipo = "Stock"
        else:
            v_Tipo = ""
        # v_Tipo = self.selected.get()
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
                if not os.path.isfile(ruta):
                    raise FileNotFoundError(f"El archivo no existe: {ruta}")
                else:
                    distribuidor = pd.read_excel(ruta,sheet_name=0)
                print(distribuidor.columns)
                if v_Tipo == 'SellOut':
                    if distribuidor['fecha_Facturacion'].isnull().any() or distribuidor['fecha_Facturacion'].any() == "":
                        self.error_fechas()
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
                                                'numero_Convenio'],
                                                Columnas_Stk=[])
                    
                    if self.Distribuidores.Num_Distri == Num_Distri:
                        print("Inicialización de transformación...")
                        # Obtener el Nombre de distribuidor.
                        self.ruta_base = "Base clientes.xlsx"
                        # self.ruta_base = self.resource_path("Base clientes.xlsx")
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
                        elif distribuidor[self.Distribuidores.Columnas[0]].isnull().any() or distribuidor[self.Distribuidores.Columnas[1]].isnull().any():
                            print("Columnas vacias ", self.Distribuidores.Columnas[0], " y ", self.Distribuidores.Columnas[1])
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
                    print("Validación 2")
                    self.Distribuidores = Distribuidor(v_Num_Distri=Num_Distri, v_ruta=self.ruta,  Columnas=[],
                                                        Columnas_Stk=['fecha_Inventario',
                                                                    'linea_Negocio',
                                                                    'codeProduct_Distribuidor',
                                                                    'presentacion',
                                                                    'unidad_Medida',
                                                                    'volumen_Inventario',
                                                                    # 'clave_Distribuidor',
                                                                    # 'nombre_Distribuidor',
                                                                    'almacen',
                                                                    'municipio',
                                                                    'no_ShipTo',
                                                                    'pais',
                                                                    'inventario_EnTransito'])
                    
                    if self.Distribuidores.Num_Distri == Num_Distri:
                        print("Inicialización de transformación...")
                        # Obtener el Nombre de distribuidor.
                        self.ruta_base = "Base clientes.xlsx"
                        # self.ruta_base = self.resource_path("Base clientes.xlsx")
                        self.soldTo = pd.read_excel(self.ruta_base, sheet_name=0)
                        print("Tipo de dato de Número de distribuidores: ", Num_Distri)
                        Name_Distri_aux=self.soldTo[self.soldTo["Num_Distri"]==int(Num_Distri)]
                        Name_Distri = list(Name_Distri_aux["Name_Distri"])
                        Name_Distri=Name_Distri[0]
            
                        print("Nombre del distribuidor: ", Name_Distri)
                            
                        if 'folio' not in distribuidor.columns:
                            df = self.Distribuidores.Stocks(distribuidor, self.Distribuidores.Columnas_Stk)
                            
                            indice = ruta.rfind(".xlsx")
                            nueva_ruta = ruta[:indice]
                            print(nueva_ruta)
                            
                            self.json_conver(ruta=nueva_ruta, df=df, name_distri=Name_Distri, num_distri=Num_Distri)
                            
                            print("Procedimiento finalizado con exito...")
                            sys.exit()
                        elif distribuidor[self.Distribuidores.Columnas[0]].isnull().any() or distribuidor[self.Distribuidores.Columnas[10]].isnull().any():
                            self.mostrar_error()
                        else:
                            df = self.Distribuidores.fecha_Inventario(distribuidor, self.Distribuidores.Columnas)
                            
                            indice = ruta.rfind(".xlsx")
                            nueva_ruta = ruta[:indice]
                            print(nueva_ruta)
                            
                            self.json_conver(ruta=nueva_ruta, df=df[self.Distribuidores.Columnas_Stk], name_distri=Name_Distri, num_distri=Num_Distri)
                            
                            print("Procedimiento finalizado con exito...")
                            sys.exit()
                    else:
                        self.Cliente_Not_foud()
                    sys.exit()
            except FileNotFoundError or PermissionError as e1:
                traceback.print_exc()
                self.error_Archivo()
            except IndexError as e2:
                traceback.print_exc()
                self.error_columnas()
            except KeyError as e3:
                traceback.print_exc()
                self.error_columnas()
            except TypeError as e4:
                traceback.print_exc()
                self.mostrar_error()
            except ValueError as e5:
                traceback.print_exc()
                self.Cliente_Not_foud()



# Inicialización de la clase App        
if __name__ == "__main__":
    
    app = App()
    app.mainloop()