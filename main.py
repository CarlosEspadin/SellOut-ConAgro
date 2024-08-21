#!/C:/Users/carlo/AppData/Local/Programs/Python/Python311/python.exe
# coding: latin-1
import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
import traceback
import sys
import time
from tkinter import font
# importamos la dependencia General
import SellOut

class App(tk.Tk):
    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            self.base_path = sys._MEIPASS
        except Exception:
            self.base_path = os.path.abspath(".")

        return os.path.join(self.base_path, relative_path)
    
    def __init__(self):
        super().__init__()
        self.title("Homologaciones ConAgro")
        self.geometry("950x600")
        self.resizable(0, 0)
        
        # Inicializar variables
        self.ruta = tk.StringVar()
        self.num_distri = tk.StringVar()
        self.name_distri = tk.StringVar()
        self.selected = tk.StringVar()
        # Variables de texto para cada entrada
        self.cod_syngenta_var = tk.StringVar()
        self.cod_ext_var = tk.StringVar()
        self.name_prod_distr_var = tk.StringVar()
        self.cod_ext_syc_var = tk.StringVar()
        
        self.configure(bg='#f0f0f0')  # Cambiar el color de fondo de la ventana
        
        self.iniciar_app() 
    
    def create_intro_label(self):
        intro_text = (
            "Bienvenido a la Aplicaci�n de Homologaci�n de Cat�logos ConAgro.\n\n"
            "Esta herramienta te permitir� homologar los cat�logos de productos \n"
            "de manera eficiente y precisa. Por favor, sigue las instrucciones \n"
            "para cargar y procesar los datos.\n\n"
            "Para comenzar, selecciona el tipo de proceso en el men� lateral\n"
            "y selecciona el archivo a homologar, luego proporciona la \n"
            "informaci�n requerida en los campos correspondientes."
        )
        self.intro_label = ttk.Label(self, text=intro_text, style='TLabel', justify=tk.LEFT)
    
    def create_widgets(self):
        entry_font = {'font': ('Helvetica', 12)}
        
        # C�digos de producto con clave Syngenta
        self.CodSygenta = ttk.Label(self, text="Ingresa el nombre de la columna los c�digos Syngenta:", style='TLabel')
        self.CodSygenta_entry = ttk.Entry(self, textvariable=self.cod_syngenta_var, **entry_font)
        self.CodSygenta_entry.config(width=20)
        self.B5 = ttk.Button(self, text="Insertar", command=self.obtener_CodSycDise)
        
        # C�digos de productos con clave externa
        self.CodExt = ttk.Label(self, text="Ingresa el nombre de la columna los c�digos Externos:", style='TLabel')
        self.CodExt_entry = ttk.Entry(self, textvariable=self.cod_ext_var, **entry_font)
        self.CodExt_entry.config(width=20)
        self.B7 = ttk.Button(self, text="Insertar", command=self.obtener_CodExt)
        
        # Descripciones de los productos
        self.NameProdDistr = ttk.Label(self, text="Ingresa el nombre de la columna las descripciones:", style='TLabel')
        self.NameProdDistr_entry = ttk.Entry(self, textvariable=self.name_prod_distr_var, **entry_font)
        self.NameProdDistr_entry.config(width=20)
        self.B6 = ttk.Button(self, text="Insertar", command=self.obtener_NameProdDistr)
        
        # Codigos de productos con clave externa cuando de origen viene con claves Syngenta
        self.CodExt_Syc = ttk.Label(self, text="Ingresa el nombre de la columna los c�digos Externos:", style='TLabel')
        self.CodExt_entry_Syc = ttk.Entry(self, textvariable=self.cod_ext_syc_var, **entry_font)
        self.CodExt_entry_Syc.config(width=20)
        self.B8 = ttk.Button(self, text="Insertar", command=self.obtener_CodExt_Syc)

        # Inicialmente ocultar todos los elementos
        self.hide_all()
        
    def search_num(self, event):
        value = self.num_distri.get()
        if value == '':
            data = self.aux_Num
        else:
            data = []
            for item in self.aux_Num:
                if value.lower() in item.lower():
                    data.append(item)
        self.distri_entry['values'] = data

    def search_name(self, event):
        value = self.name_distri.get()
        if value == '':
            data = self.aux_Name
        else:
            data = []
            for item in self.aux_Name:
                if value.lower() in item.lower():
                    data.append(item)
        self.name_entry['values'] = data

                    

    def create_all_widgets(self):
        self.ruta_label = ttk.Label(self, text="Archivo seleccionado con exito", style='TLabel')
        
        self.browse_button = ttk.Button(self, text="Seleccionar archivo", command=self.seleccionar_archivo)
        
        self.ruta_label = ttk.Label(self, text="Selecciona el cat�logo a Homologar:", style='TLabel')
        
        # Establecemos ruta absoluta de archivo con datos de distribuidores:
        self.ruta_CatalogoDistribuidores = self.resource_path("Catalogo Distribuidores.xlsx")
        self.Catalogo = pd.read_excel(self.ruta_CatalogoDistribuidores, sheet_name=0)
        self.aux_Num = list(self.Catalogo['Sold to'].sort_values().unique())
        self.aux_Name = list(self.Catalogo['Descripci�n'].unique())
        # self.aux_Num = ""
        # self.aux_Name = ""
        
        # Convertir n�meros a cadenas
        self.aux_Num = [str(num) for num in self.aux_Num]
        
        # Convertir n�meros a cadenas
        self.aux_Name = [str(num) for num in self.aux_Name]
        
        # Entradas para el numero de distribuidor
        self.distri_label = ttk.Label(self, text="Ingresa el N�mero de cliente:", style='TLabel')
        
        self.distri_entry = ttk.Combobox(self, textvariable=self.num_distri, font=('Helvetica', 11), state='normal')
        self.distri_entry['values'] = self.aux_Num
        self.distri_entry.bind('<KeyRelease>', self.search_num)
        self.distri_entry.config(width=40)
        
        
        self.B3 = ttk.Button(self, text="Insertar", command=self.obtener_num)
        
        self.name_label = ttk.Label(self, text="Ingresa el nombre del cliente:", style='TLabel')
        
        # Entrada para el nombre del distribuidor
        self.name_entry = ttk.Combobox(self, textvariable=self.name_distri, font=('Helvetica', 11), state='normal')
        self.name_entry['values'] = self.aux_Name
        self.name_entry.bind('<KeyRelease>', self.search_name)
        self.name_entry.config(width=40)

        
        self.B4 = ttk.Button(self, text="Insertar", command=self.obtener_nombre)
        
        self.B2 = ttk.Button(self, text="Homologar cat�logo", command=self.mostrar_mensaje)
        
    def show_intro_label(self):
        self.intro_label.place(relx=0.33, rely=0.1, relwidth=0.8, relheight=0.3)
        
    def hide_intro_labe(self):
        self.intro_label.place_forget()
    
    def show_all_widgets(self):
        self.hide_intro_labe()
        self.nx = 0 if self.menu_lateral.winfo_ismapped() else 0.2
        self.ruta_label.place(relx=0.3-self.nx, rely=0.15)
        self.browse_button.place(relx=0.3-self.nx, rely=0.1)
        self.ruta_label.place(relx=0.3-self.nx, rely=0.2)
        self.distri_label.place(relx=0.3-self.nx, rely=0.3)
        self.distri_entry.place(relx=0.35-self.nx, rely=0.35)
        self.B3.place(relx=0.78-self.nx, rely=0.35)
        self.name_label.place(relx=0.3-self.nx, rely=0.4)
        self.name_entry.place(relx=0.35-self.nx, rely=0.45)
        self.B4.place(relx=0.78-self.nx, rely=0.45)
        self.B2.place(relx=0.85-self.nx, rely=0.9, anchor='center')
        
    def show_hide_widgets(self):
        self.hide_intro_labe()
        self.ruta_label.place_forget()
        self.browse_button.place_forget()
        self.ruta_label.place_forget()
        self.distri_label.place_forget()
        self.distri_entry.place_forget()
        self.B3.place_forget()
        self.name_label.place_forget()
        self.name_entry.place_forget()
        self.B4.place_forget()
        self.B2.place_forget()
    
    def show_chile_widgets(self):
        self.browse_button.place(relx=0.3, rely=0.1)
        
    def hide_chile_widgets(self):
        self.browse_button.place_forget()
        
    def create_greeting_message(self):
        # Obtener el valor de Tipo
        Tipo = self.selected.get()
        
        # Mostrar/Ocultar elementos seg�n el valor de Tipo
        if Tipo == "1":
            self.show_all_widgets()
            # Mostrar elementos de tipo Syngenta
            # self.show_syngenta_elements()
            # Ocultar elementos de tipo Externo
            # self.hide_external_elements()
            # self.hide_chile_widgets()
        elif Tipo == "0":
            self.show_hide_widgets()
            self.hide_syngenta_elements()
            self.hide_external_elements()
            self.hide_NameProdDistr()
            self.hide_chile_widgets()
            self.show_intro_label()
        elif Tipo == "2":
            self.show_all_widgets()
            # Mostrar elementos de tipo Externo
            # self.show_external_elements()
            # Ocultar elementos de tipo Syngenta
            # self.hide_syngenta_elements()
            # self.hide_chile_widgets()
        else:
            self.show_hide_widgets()
            self.hide_syngenta_elements()
            self.hide_external_elements()
            self.hide_NameProdDistr()
            self.hide_intro_labe()
            # self.show_hide_widgets()
            self.show_chile_widgets()
    
    def show_syngenta_elements(self):
        self.CodSygenta.place(relx=0.31-self.nx, rely=0.55)
        self.CodSygenta_entry.place(relx=0.65-self.nx, rely=0.6)
        self.B5.place(relx=0.88-self.nx, rely=0.6)
        self.NameProdDistr.place(relx=0.31-self.nx, rely=0.65)
        self.NameProdDistr_entry.place(relx=0.65-self.nx, rely=0.7)
        self.B6.place(relx=0.88-self.nx, rely=0.7)
        self.CodExt_Syc.place(relx=0.31, rely=0.75)
        self.CodExt_entry_Syc.place(relx=0.65-self.nx, rely=0.8)
        self.B8.place(relx=0.88-self.nx, rely=0.8)

    def hide_syngenta_elements(self):
        self.CodSygenta.place_forget()
        self.CodSygenta_entry.place_forget()
        self.B5.place_forget()
        # self.NameProdDistr.place_forget()
        # self.NameProdDistr_entry.place_forget()
        # self.B6.place_forget()
        self.CodExt_Syc.place_forget()
        self.CodExt_entry_Syc.place_forget()
        self.B8.place_forget()

    def show_external_elements(self):
        self.CodExt.place(relx=0.31-self.nx, rely=0.55)
        self.CodExt_entry.place(relx=0.65-self.nx, rely=0.6)
        self.B7.place(relx=0.88-self.nx, rely=0.6)
        self.NameProdDistr.place(relx=0.31-self.nx, rely=0.65)
        self.NameProdDistr_entry.place(relx=0.65-self.nx, rely=0.7)
        self.B6.place(relx=0.88-self.nx, rely=0.7)

    def hide_external_elements(self):
        self.CodExt.place_forget()
        self.CodExt_entry.place_forget()
        self.B7.place_forget()
        
    def hide_NameProdDistr(self):
        self.NameProdDistr.place_forget()
        self.NameProdDistr_entry.place_forget()
        self.B6.place_forget()
    
    def hide_all(self):
        self.hide_syngenta_elements()
        self.hide_external_elements()
    
    def paneles(self):        
        # Configuraci�n de barra superior
        self.barra_superior = tk.Frame(
            self, bg='#5f7800', height=40
        )
        
        self.barra_superior.pack(side=tk.TOP, fill='both')   
        
        # Etiqueta de t�tulo
        self.labelTitulo = tk.Label(self.barra_superior, text="Aplicaci�n de consumo SellOut/Stocks")
        self.labelTitulo.config(fg="#fff", font=(
            "Roboto", 15), bg="#5f7800", pady=10, width=46)
        self.labelTitulo.pack(side=tk.TOP)  

        self.menu_lateral = tk.Frame(self, bg="#abb400", width=150)
        self.menu_lateral.pack(side=tk.LEFT, fill='both', expand=False)
        
    def controles_barra_superior(self):
        # Configuraci�n de la barra superior
        font_awesome = font.Font(family='FontAwesome', size=12)

        # Establecemos la ruta absoluta del icono del boton
        self.ruta_menu = self.resource_path("menu.png")
        
        # configure icon
        self.icon_menu = tk.PhotoImage(file=self.ruta_menu)
        
        # self.nx = float(self.selected.get())-0.8 if self.menu_lateral.winfo_ismapped() else float(self.selected.get())-0.8
        # Bot�n del men� lateral
        self.buttonMenuLateral = tk.Label(self.barra_superior, text="Men�", font=("Roboto", 15),
                                bd=0, bg="#5f7800", fg="#fff", image=self.icon_menu, compound="right", padx=10)
        self.buttonMenuLateral.place(relx=0.02, rely=0.25)

    
    def controles_menu_lateral(self):
        # Configuraci�n del men� lateral
        ancho_menu = 20
        alto_menu = 2
        font_awesome = font.Font(family='FontAwesome', size=15)
        
        # Etiqueta de perfil
        self.labelPerfil = tk.Label(self.menu_lateral, image=self.icon_big, bg="#abb400")
        self.labelPerfil.pack(side=tk.TOP, pady=10)

        # Crear los botones con icono y texto
        self.buttonDashBoard = self.create_menu_button(self.menu_lateral, "Inicio", "\uf109", self.create_greeting_message, "0")
        self.bottonSync = self.create_menu_button(self.menu_lateral, "Tipo SellOut", "\uf007", self.create_greeting_message, "1")
        self.bottonExt = self.create_menu_button(self.menu_lateral, "Tipo Stocks", "\uf03e", self.create_greeting_message, "2")
        # self.bottonCL = self.create_menu_button(self.menu_lateral, "Chile", "\uf129", self.create_greeting_message, "3")
        self.buttonSettings = self.create_menu_button(self.menu_lateral, "Settings", "\uf013", self.create_greeting_message, "3")

        
    def create_menu_button(self, parent, text, icon, command, value=None):
        frame = tk.Frame(parent, bg="#abb400")
        frame.pack(side=tk.TOP, fill=tk.X)

        label_icon = tk.Label(frame, text=icon, font=("FontAwesome", 15), bg="#abb400", fg="white")
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
        # Asociar eventos Enter y Leave con la funci�n din�mica
        frame.bind("<Enter>", lambda event: self.on_enter(event, frame, label_icon, label_text))
        frame.bind("<Leave>", lambda event: self.on_leave(event, frame, label_icon, label_text))
        label_icon.bind("<Enter>", lambda event: self.on_enter(event, frame, label_icon, label_text))
        label_icon.bind("<Leave>", lambda event: self.on_leave(event, frame, label_icon, label_text))
        label_text.bind("<Enter>", lambda event: self.on_enter(event, frame, label_icon, label_text))
        label_text.bind("<Leave>", lambda event: self.on_leave(event, frame, label_icon, label_text))

    def on_enter(self, event, frame, label_icon, label_text):
        # Cambiar estilo al pasar el rat�n por encima
        frame.config(bg="#009933")
        label_icon.config(bg="#009933")
        label_text.config(bg="#009933")

    def on_leave(self, event, frame, label_icon, label_text):
        # Restaurar estilo al salir el rat�n
        frame.config(bg="#abb400")
        label_icon.config(bg="#abb400")
        label_text.config(bg="#abb400")

    def on_click(self, value, command):
        # Obtener el valor de la variable de tipo de ventana
        self.selected.set(value)
        command()
    
    def iniciar_app(self):
        self.deiconify()
        
        paddings = {'padx': 10, 'pady': 10}
        entry_font = {'font': ('Helvetica', 12)}
        
        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=3)
        
        # Estilos
        self.style = ttk.Style(self)
        self.style.configure('TLabel', font=('Helvetica', 12), foreground='#333333', background='#f0f0f0')
        self.style.configure('TButton', font=('Helvetica', 12), background='#28536b', foreground='black')
        self.style.map('TButton', background=[('active', '#45a049')], foreground=[('active', 'black')])  # Color al pasar el rat�n

        
        # Aplicar transparencia
        self.attributes('-alpha', 0.98)  # Valor entre 0 (completamente transparente) y 1 (opaco)
        
        # Establecemos la ruta absoluta:
        self.ruta_ico_s = self.resource_path("ConAgro_icon_small.png")
        self.ruta_ico_b = self.resource_path("ConAgro_icon_big.png")
        
        # configure icon
        self.icon_big = tk.PhotoImage(file=self.ruta_ico_b)
        self.icon_small = tk.PhotoImage(file=self.ruta_ico_s)
        self.iconphoto(False, self.icon_big, self.icon_small)
        # self.iconbitmap('ConAgro.ico')
        
        # Etiqueta de informacion
        self.labelTitulo = ttk.Label(self, text="Reportar cualquier incidente con el aplicativo con: \ncarlos.espadin@syngenta.com", font=("Roboto", 10), background='#f0f0f0', foreground='black') # Modificar
        self.labelTitulo.place(relx=0.35, rely=0.93, anchor='center')
        # self.labelTitulo.pack(side=tk.RIGHT)
        
        # Invocacion de paneles
        self.paneles()
        self.controles_barra_superior()
        self.controles_menu_lateral()
        
        self.create_intro_label()
        self.show_intro_label()
        self.create_all_widgets()
        
        self.create_widgets()
        

    # M�todo para cerrar la venta App.
    def cerrar_ventana(self):
        self.destroy()
    
    def seleccionar_archivo(self):
        filetypes = (('Archivo de Excel', '*.xlsx'), ('All files', '*.*'))
        filename = filedialog.askopenfilename(title='Abrir archivo', initialdir='/', filetypes=filetypes)
        ruta_archivo = os.path.basename(filename)
        self.ruta_label.config(text="Archivo:\n"+ruta_archivo+" seleccionado con exito")
    # Ventana de error si el n�mero de cliente ingresado no existe.
    def Cliente_Not_foud(self):  # sourcery skip: class-extract-method
        messagebox.showerror("Error", "El cliente que est�s ingresando no est� registrado en ConAgro")
        self.after(30000, self.cerrar_ventana)
        sys.exit()
    # Ventana que muestra error si la ruta del archivo es incorrecta.
    def mostrar_error(self):
        messagebox.showerror("Error", "La ruta del archivo o el n�mero de distribuidor no son v�lidos. Por favor vuelve a ejecutar el programa e ingresa los valores adecuador.")
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
    
    def obtener_num(self):
        numero = self.num_distri.get()
        messagebox.showinfo("N�mero de cliente", f"Se ha ingresado el n�mero de cliente: {numero}")
        self.B3.config(state=tk.DISABLED)
        
    
    def obtener_nombre(self):
        nombre = self.name_distri.get()
        messagebox.showinfo("Nombre de cliente", f"Se ha ingresado el nombre de cliente: {nombre}")
        self.B4.config(state=tk.DISABLED)
    
    def obtener_CodSycDise(self):
        cod_syngenta = self.cod_syngenta_var.get()
        messagebox.showinfo("C�digo Syngenta", f"Se ha ingresado el c�digo Syngenta: {cod_syngenta}")
        self.B5.config(state=tk.DISABLED)
    
    def obtener_CodExt(self):
        cod_ext = self.cod_ext_var.get()
        messagebox.showinfo("C�digo Externo", f"Se ha ingresado el c�digo Externo: {cod_ext}")
        self.B7.config(state=tk.DISABLED)
    
    def obtener_NameProdDistr(self):
        name_prod_distr = self.name_prod_distr_var.get()
        messagebox.showinfo("Nombre de Producto Distribuidor", f"Se ha ingresado el nombre del producto distribuidor: {name_prod_distr}")
        self.B6.config(state=tk.DISABLED)
    
    def obtener_CodExt_Syc(self):
        cod_ext_syc = self.cod_ext_syc_var.get()
        messagebox.showinfo("C�digo Externo de Syngenta", f"Se ha ingresado el c�digo externo de Syngenta: {cod_ext_syc}")
        self.B8.config(state=tk.DISABLED)

    def format_dataframe_as_table(self, df):
        # Convertir DataFrame a texto con formato de tabla
        # Ajustar el ancho de las columnas basado en el contenido
        col_widths = [max(df[col].astype(str).map(len).max(), len(col)) for col in df.columns]
        
        table = ''
        # Crear encabezado
        header = ' | '.join(f"{col:{col_widths[i]}}" for i, col in enumerate(df.columns))
        table += header + '\n'
        table += '-' * len(header) + '\n'
        
        # Crear filas
        for index, row in df.iterrows():
            row_text = ' | '.join(f"{str(row[col]):{col_widths[i]}}" for i, col in enumerate(df.columns))
            table += row_text + '\n'
        
        return table
    
    def close_toplevel(self):
        if hasattr(self, 'top'):
            self.top.destroy()
            self.destroy()
            sys.exit()


    
    def mostrar_mensaje(self):
        try:
            DF_Proces = self.Procesamiento(
                Tipo=int(self.selected.get()),
                Num_Distri=self.num_distri.get(),
                Name_Distri=self.name_distri.get(),
                ruta=self.ruta.get(),
                CodSyngenta=self.cod_syngenta_var.get(),
                NomDistriProd=self.name_prod_distr_var.get(),
                CodDistriProd_Syc=self.cod_ext_syc_var.get(), 
                CodDistriProd=self.cod_ext_var.get()
            )
            
            
            # Crear ventana Toplevel
            self.top = tk.Toplevel(self)
            self.top.title("Informaci�n de homologaci�n")
            self.top.geometry("800x400")
            
            # Crear Treeview
            tree = ttk.Treeview(self.top, columns=list(DF_Proces.columns), show='headings')
            tree.pack(fill='both', expand=True)
            
            # Configurar columnas del Treeview
            for col in DF_Proces.columns:
                tree.heading(col, text=col)
                tree.column(col, anchor='center', width=100)
            
            # Insertar datos en el Treeview
            for index, row in DF_Proces[DF_Proces['DescSyngenta'].duplicated()].iterrows():
                tree.insert("", "end", values=list(row))
            
            # Crear y mostrar informaci�n adicional
            info_label = tk.Label(self.top, text=(
                "Ruta: " + self.ruta.get() +
                "\nN�mero de distribuidor: " + self.num_distri.get() +
                "\nNombre del distribuidor: " + self.name_distri.get() +
                "\nTipo de producto: " + self.selected.get()
            ), padx=10, pady=10)
            info_label.pack(side='top', anchor='w')
            
            self.close_button = ttk.Button(self.top, text="Terminar Homologaci�n", command=self.close_toplevel)
            self.close_button.pack(pady=20)
        
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
        
    def Procesamiento(self, Tipo, Num_Distri, Name_Distri, ruta):
        self.iniciar_app()# sourcery skip: extract-duplicate-method, extract-methodate-method, extract-method
        # Deshabilitamos el boton aceptar para que no se pueda usar m�s de una vez.
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
                    print("Validaci�n 1")
                    self.Distribuidores = SellOut.Distribuidores(v_Num_Distri=Num_Distri, v_ruta=self.ruta,  Columnas=[
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
                        print("Inicializaci�n de transformaci�n...")
                        # Obtener el Nombre de distribuidor.
                        self.ruta_base = self.resource_path("Base clientes.xlsx")
                        self.soldTo = pd.read_excel(self.ruta_base, sheet_name=0)
                        print("Tipo de dato de N�mero de distribuidores: ", Num_Distri)
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
                    print("Validaci�n 2")
                    self.Distribuidores = Distribuidor(v_Num_Distri=Num_Distri, v_ruta=self.ruta,  Columnas=[],
                                                        Columnas_Stk=['fecha_Inventario',
                                                                    'linea_Negocio',
                                                                    'codeProduct_Distribuidor',
                                                                    'presentacion',
                                                                    'unidad_Medida',
                                                                    'volumen_Inventario',
                                                                    'clave_Distribuidor',
                                                                    'nombre_Distribuidor',
                                                                    'almacen',
                                                                    'municipio',
                                                                    'no_ShipTo',
                                                                    'pais',
                                                                    'inventario_EnTransito'])
                    
                    if self.Distribuidores.Num_Distri == Num_Distri:
                        print("Inicializaci�n de transformaci�n...")
                        # Obtener el Nombre de distribuidor.
                        self.ruta_base = self.resource_path("Base clientes.xlsx")
                        self.soldTo = pd.read_excel(self.ruta_base, sheet_name=0)
                        print("Tipo de dato de N�mero de distribuidores: ", Num_Distri)
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

if __name__ == "__main__":
    app = App()
    app.mainloop()