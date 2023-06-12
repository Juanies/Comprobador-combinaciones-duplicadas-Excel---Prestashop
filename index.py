# librerias necesarias para utilizar la aplicación
import pandas as pd 
import requests
import xmltodict
import json
import webbrowser
from datetime import datetime


from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import os 

from ttkwidgets import CheckboxTreeview

# Generamos la infertaz
win = Tk()
# Geometria de la ventana principal
# Titulo de la Interfaz principal
win.title("Comprobador EAN Prestashop")


# Tamaño maximo de la ventana principal
win.resizable(0, 0)
# Tamaño Minimo de la ventana principal

# Frame principales (Arriba y abajo)
left_frame  =  Frame(win,  width=200,  height=  400)
left_frame.grid(row=0,  column=1,  padx=0,  pady=5)
right_frame  =  Frame(win,  width=650,  height=400)
right_frame.grid(row=1,  column=1,  padx=0,  pady=0)

# Frame Directorio y buscar además de Analizar y Eliminar duplicados
tool_bar  =  Frame(left_frame,  width=180,  height=185)
tool_bar.grid(row=2,  column=0,  padx=5,  pady=5)

# Frame Analizar y Eliminar duplicados
frame2 = Frame(tool_bar,width=100,height=100)
frame2.grid(row=2,  column=1, padx=0,  pady=0)

# Lectura del JSON
with open('shops.json') as file:
    # Cargamos el fichero Json
    datashop = json.load(file)
# Nombre de todas las tiendas
nombres = datashop['tiendas']
nombres2 = []
# Diccionario 
d3 = []



# Creacion del menu para seleccionar las tiendas
combo = CheckboxTreeview(win)
combo.grid(row=0, column=0, sticky="WE", padx=20)    

# Función para obtener el valor seleccionado
txt_output = Listbox(bd = 2, fg = "black", cursor = "target")
txt_output.grid(row=1,  column=0,  padx=0,  pady=20)
# Funcion coger elementos seleccionados
txt_output.insert(0, "Tiendas seleccionadas")

global selected
# Funcion para coger las tiendas seleccionadas
def get_selected_items(event=None):
    txt_output.delete(0, END)
    global selected
    selected = []
    # Cogemos los elementos seleccionados
    selected_items = combo.get_checked()
    # Recorremos los elementos seleccionados
    for item in selected_items:
        # Si existe el elemento
        if combo.exists(item):
            # Lo pone en Selected
            selected.append(combo.item(item)['text'])
            # Recorremos los elementos seleccionados
    for item in selected:
        # Lo insertamos al txt_output
        txt_output.insert(END, item)
    return selected

# Registrar el evento de selección en el CheckboxTreeview
combo.bind("<<TreeviewSelect>>", get_selected_items)

s = []
tienda = {}
# Recorremos las tiendas seleccionadas y las añadimos al diccionaro 
for i in range(len(datashop['tiendas'])): 
    # Variable con el nombre de las tiendas
    nombrestiendas = datashop['tiendas'][i]['name']
    # Variable con la api de las tiendas
    apistiendas = datashop['tiendas'][i]['api']
    # Variable con el link de las tiendas
    linktiendas = datashop['tiendas'][i]['link']
    # Creacion del diccionario
    tienda = {
        "name": nombrestiendas,
        "data":{
            'link': linktiendas,
            'api': apistiendas,
        },
    }
    # Añadimos los datos al diccionario
    d3.append(tienda)
    # Guardamos los nombres en una variable
    nombres2.append(datashop['tiendas'][i]['name'])
    # Insertamos los nombres paa poder elegirlos 

# Insertamos los nombres de las tiendas para poder seleccionarlo
for i in range(len(datashop['tiendas'])):
    combo.insert("", "end", text=(datashop['tiendas'][i]['name']))
# Generamos la variable URL Global
api = []
global url

# Variable Mensaje EAN duplicado
doopemsg = StringVar()
# Variable filas global
global filas
# Duplicados
tableelement = []

# Generamos las variables duplicados como global
global duplicados
# Elemento sin duplicar

# Generamos la variable de las filsd
filas = " "

# Funcion para realizar la URL para solicitar datos a Prestashop 
def apiurl():
    # Colocamos la variable URL como Global
    global url, duplicados
    # Recorremos las tiendas seleccionadas
    for item in selected:
        # Guardamos la api de la tienda en otra variable
        api_value = apid3
        # Guardamos el link de la tienda en otra variable
        link_values = linkd3
        if api_value is not None:
            # Variable api parte urk
            api = "api/"
            # Formamos la url de la api
            api_url = f'{link_values}{api}'
            # Cogemos el valor de la API
            api_key = api_value
            # Colocamos el elemento que vamos a usar de la API
            resource = 'combinations'
            # Cogemos el display que vamos a usar de la tienda
            schema_param = 'display=[ean13]'
            # url final  
            url = f'{api_url}/{resource}?{schema_param}&ws_key={api_key}'
        else:
            print("API no encontrada")
    
schema_param = 'display=[ean13]'

# Analizamos el fichero y comparamos con PrestaShop
def Analizar():
    global duplicados, filas, df, fila, df_output
    # Leemos el excel
    df = pd.read_excel(filepath, header = 0, dtype='object')
    df_output = pd.read_excel(filepath, header = 0, dtype='object')
    dftable = pd.read_excel(filepath, header = 0, usecols = "A, B, C, H", dtype='object')
    # Guardando los datos del Excel 
    global named3, linkd3, apid3
    # Recorremos las tiendas seleccionadas

    for item in selected:
        # Cogemos los datos del diccionario
        print(selected)
        print(item)
        
        for datosD3 in d3:
            fila = df.iloc[7] 
            filas = df.shape[0]
            codigoarticulo = []
            ids = []
            descripcion = []
            talla = []
            codbarras2 = []
            # Guardamos los nombres del diccionario con los datos
            named3 = datosD3['name']
            if (named3 == item):            # Guardamos el link del diccionario con los datos
                linkd3 = datosD3['data']['link']
                # Guardamos la api del diccionario con los datos
                apid3 = datosD3['data']['api']
                # Ejecutamos API URL 
                apiurl()    

                # Solicitamos los datos de Prestashop para coger los datos de los Ean13 y posteriormente generar el fichero
                # Solicitamos la respuesta
                response = requests.get(url, params=schema_param)
                # Cogemos el contenido de la respuesta
                content = response.content
                # Pasamos los datos a JSON
                doc_json = xmltodict.parse(content)
                # Pasamos los datos a str
                json_data = json.dumps(doc_json)
                # Generamos diccionario con llos datos del JSON
                json_dict = json.loads(json_data)   
                # Cogemos los datos de los EAN de la API de Prestashop
                principal = json_dict['prestashop']['combinations']['combination']
                ean = principal
                # Generamos las variables de los duplicados y los no dupliucados
                duplicados = []
                noduplicados = []
                # Recorremos las filas
                for i in range(filas -1):
                    # Cogemos los datos y los guardamos en varias listas segun el tipo de dato     
                    codigoarticulo.append(df.iloc[i][0])
                    descripcion.append(df.iloc[i][1])
                    talla.append(df.iloc[i][2])
                    codbarras2.append(df.iloc[i][7])
                    # Recorremos los EAN
                    for eanAux in ean:
                        # Cogemos el valor de los Ean13
                        ean13 = eanAux.get('ean13')
                        # Añadimos los ean a la lista de los duplicados o no duplicados dependiendo de si lo estan o no
                        if ean13 is not None:
                            if ean13 != df.iloc[i][7]:
                                # Añadimos a la lista de noduplicados los Ean sin duplicar
                                noduplicados.append(ean13)
                            else:
                                # Añadimos a la Lista duplicados los Ean duplicados
                                duplicados.append(ean13)
                                # Añadimos a la Lista las Ids 
                                ids.append(df.iloc[i][0])
                        if i == 0 or i == filas -1:
                            continue
                print(duplicados)
                if duplicados:
                    # Cogemos los datos de duplicados y lo colocamos en el msg
                    doopemsg.set(duplicados)
                else:
                    # Si no hay duplicados lo colocamos en el msg
                    doopemsg.set("Sin duplicados")
                    #Generamos los elementos duplicados
                tableelement.append(duplicados)
                # Insertamos los datos a la tabla de Tkinter
                for codigoarticulo, descripcion, talla, codbarras2 in zip(codigoarticulo, descripcion, talla, codbarras2):
                    # Insertamos los datos
                    table.insert("", "end", values=(codigoarticulo, descripcion, talla, codbarras2))
    return

# Variable mensaje Directorio
pathmsg = StringVar()
# Mostrar Mensaje Directorio
pathlabe = Entry(tool_bar, state= "disabled",  textvariable=pathmsg, relief=RAISED, width=50).grid(row=0,  column=1,  padx=0,  pady=0,  ipadx=0)
# Exportamos la variable
global filepath        
# Funcion para buscar el directorio
def Directorio():
    # Reseteamos la tabla
    for item in table.get_children():
        table.delete(item)
        # Menu de elegir archivo y extension permitida
    global filepath        
        # Seleccionamos los tipos de datos permitidos
    file = filedialog.askopenfile(mode='r', filetypes=[('all files',"*.*") ])
        # Si un archivo
    if file:
            # Cogemos el valor del directorio y lo metemos en FilePath
        filepath = os.path.abspath(file.name)
            # Le damos el valor del directorio a la variable pathmsg
        pathmsg.set(filepath)

# Funcion para eliminar los duplicados del Excel
def eliminarduplicado():
    
    
    global duplicados, filas, df, fila, df_output
    # Leemos el excel
    df = pd.read_excel(filepath, header = 0, dtype='object')
    df_output = pd.read_excel(filepath, header = 0, dtype='object')
    dftable = pd.read_excel(filepath, header = 0, usecols = "A, B, C, H", dtype='object')
    # Guardando los datos del Excel 
    global named3, linkd3, apid3
    # Recorremos las tiendas seleccionadas

    for item in selected:
        # Cogemos los datos del diccionario
        print(selected)
        print(item)
        
        for datosD3 in d3:
            fila = df.iloc[7] 
            filas = df.shape[0]
            codigoarticulo = []
            ids = []
            descripcion = []
            talla = []
            codbarras2 = []
            # Guardamos los nombres del diccionario con los datos
            named3 = datosD3['name']
            if (named3 == item):            # Guardamos el link del diccionario con los datos
                linkd3 = datosD3['data']['link']
                # Guardamos la api del diccionario con los datos
                apid3 = datosD3['data']['api']
                # Ejecutamos API URL 
                apiurl()    

                # Solicitamos los datos de Prestashop para coger los datos de los Ean13 y posteriormente generar el fichero
                response = requests.get(url, params=schema_param)
                # Cogemos el contenido de la respuesta
                content = response.content
                # Pasamos los datos a JSON
                doc_json = xmltodict.parse(content)
                # Pasamos los datos a str
                json_data = json.dumps(doc_json)
                # Generamos diccionario con llos datos del JSON
                json_dict = json.loads(json_data)   
                # Cogemos los datos de los EAN de la API de Prestashop
                principal = json_dict['prestashop']['combinations']['combination']
                ean = principal
                # Generamos las variables de los duplicados y los no dupliucados
                duplicados = []
                noduplicados = []
                # Recorremos las filas
                for i in range(filas -1):
                    # Cogemos los datos y los guardamos en varias listas segun el tipo de dato     
                    codigoarticulo.append(df.iloc[i][0])
                    descripcion.append(df.iloc[i][1])
                    talla.append(df.iloc[i][2])
                    codbarras2.append(df.iloc[i][7])
                    # Recorremos los EAN
                    for eanAux in ean:
                        # Cogemos el valor de los Ean13
                        ean13 = eanAux.get('ean13')
                        # Añadimos los ean a la lista de los duplicados o no duplicados dependiendo de si lo estan o no
                        if ean13 is not None:
                            if ean13 != df.iloc[i][7]:
                                # Añadimos a la lista de noduplicados los Ean sin duplicar
                                noduplicados.append(ean13)
                            else:
                                # Añadimos a la Lista duplicados los Ean duplicados
                                duplicados.append(ean13)
                                # Añadimos a la Lista las Ids 
                                ids.append(df.iloc[i][0])
                        if i == 0 or i == filas -1:
                            continue
                print(duplicados)

                # Filtrar el DataFrame para eliminar las filas que contienen códigos de barras duplicados en PrestaShop
                df = df[~df['Cód Barras 2'].isin(duplicados)]
                # Generamos el fichero con el nombre sin duplicados 
                output_filepath = os.path.splitext(filepath)[0] + item + "_sin_duplicados.xlsx"
                # Guardar el DataFrame actualizado en el archivo XLSX original
                df.to_excel(output_filepath, engine='xlsxwriter')
                # Mensajes de comprobación
                print(filepath)
                print("Duplicados eliminados.")
                print("Nuevo tamaño del DataFrame:", df.shape[0])
                print("El archivo ha sido actualizado.")

                        # Limpiar la tabla antes de mostrar los datos actualizados
                for item in table.get_children():
                            # Reseteamos los datos de la tabla antes de mostrar los siguientes datos de la tabla
                    table.delete(item)                
                
                
                if duplicados:
                    # Cogemos los datos de duplicados y lo colocamos en el msg
                    doopemsg.set(duplicados)
                else:
                    # Si no hay duplicados lo colocamos en el msg
                    doopemsg.set("Sin duplicados")
                    #Generamos los elementos duplicados
                tableelement.append(duplicados)
                # Insertamos los datos a la tabla de Tkinter
                for codigoarticulo, descripcion, talla, codbarras2 in zip(codigoarticulo, descripcion, talla, codbarras2):
                    # Insertamos los datos
                    table.insert("", "end", values=(codigoarticulo, descripcion, talla, codbarras2))

            # Ejecutamos la funcion Analizar
    Analizar()

# Boton de buscar
Button(tool_bar,  text="Buscar", command=Directorio,  relief=RAISED).grid(row=0,  column=2,  padx=10,  pady=3)
# Boton de Analizar
Button(frame2,  text="Analizar", command=Analizar,  relief=RAISED).grid(row=0,  column=0,  padx=2,  pady=0)
# Boton de Eliminar duplicadps
Button(frame2,  text="Eliminar duplicados", command=eliminarduplicado,  relief=RAISED).grid(row=0,  column=1,  padx=2,  pady=0)

# Tabla con los datos
table = ttk.Treeview(win, columns=('#1', '#2', '#3', '#4'), selectmode=EXTENDED, show='headings')
# Formato de la tabla
table.grid(row=1, column=1, sticky="WE", padx=20)

# Generamos la columna Codigo Articulo
table.heading('#1', text='Código Artículo')
table.column("#1", anchor=CENTER, stretch=YES,width=100)
# Generamos la columna descripción
table.heading('#2', text='Descripción')
table.column("#2", anchor=CENTER, stretch=YES, width=100)
# Generamos la columna Talla
table.heading('#3', text='Talla')
table.column("#3", anchor=CENTER, stretch=YES, width=100)
# Generamos la columna Cod Barras 2 (ean13)
table.heading('#4', text='Cód Barras 2')
table.column("#4", anchor=CENTER, stretch=YES, width=100)


# Colocamos la marca de agua con el Nombre del creador
#html_label=HTMLLabel(win, html='<a style="font-size:.2em;" href="http://www.google.com">Hecho por Juanies</a>',  font=('Times', 2)).grid(row=2,  column=1,  padx=0,  pady=20)
#link1 = Label(win, text="Google Hyperlink", fg="blue", cursor="hand2").grid(row=2,  column=1,  padx=0,  pady=20)
#link1.bind("<Button-1>", lambda e: callback("http://www.google.com"))

def callback():
    webbrowser.open_new("https://github.com/Juanies")

year = datetime.today().year

your_variable_name = Button(text="Hecho por Juanito" + " " + str(year), command=callback)
your_variable_name.grid(column=1, row=3, pady=20)



win.mainloop()
