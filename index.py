# librerias
import pandas as pd 
import requests
import xmltodict
import json

from tkinter import *
from tkinter import ttk
from tkinter import filedialog
import os 

from ttkwidgets import CheckboxTreeview


# Interface
win = Tk()
win.geometry("650x450")
win.title("Comprobador EAN Prestashop")

# Tamaño maximo y minimo de la ventana
win.maxsize(650, 450)
win.minsize(650, 450)

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
    datashop = json.load(file)
# Nombre de todas las tiendas
nombres = datashop['tiendas']
nombres2 = []
# Diccionario 
d3 = []

# Creacion del menu
combo = CheckboxTreeview(win)
combo.grid(row=0, column=0, sticky="WE", padx=20)    
combo

s = []


# Añadiendo los elementos al Diccionario
for i in range(len(datashop['tiendas'])):
    nombrestiendas = datashop['tiendas'][i]['name']
    apistiendas = datashop['tiendas'][i]['api']
    linktiendas = datashop['tiendas'][i]['link']
    tienda = {
        nombrestiendas:{
            'link': linktiendas,
            'api': apistiendas,
        },
    }
    d3.append(tienda)
    nombres2.append(datashop['tiendas'][i]['name'])
    combo.insert("", "end", text=(datashop['tiendas'][i]['name']))




# Agregar elementos al CheckboxTreeview (código omitido) 
# Función para obtener el valor seleccionado
txt_output = Listbox(bd = 2, fg = "black", cursor = "target")
txt_output.grid(row=1,  column=0,  padx=0,  pady=20)
# Funcion coger elementos seleccionados
txt_output.insert(0, "Tiendas seleccionadas")

global selected
def get_selected_items(event=None):
    txt_output.delete(0, END)
    # Variable seleccionado
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
    print(selected)
    print()
    # Recorremos los valores Seleccionados

# Registrar el evento de selección en el CheckboxTreeview
combo.bind("<<TreeviewSelect>>", get_selected_items)

# Api de las tiendas
api = []

global url, schema_param

# Cambio de la API y URL segun la tienda elegida
def selection_changed(event):
    global url, schema_param
    selection = combo.get()
    api_value = d3[combo.current()].get(selection, {}).get('api')
    link_values = d3[combo.current()].get(selection, {}).get('link')
    if api_value is not None:
        print(api_value)
        api = "api/"
        api_url = f'{link_values}{api}'
        api_key = api_value
        resource = 'combinations'
        schema_param = 'display=[ean13]'
        url = f'{api_url}/{resource}?{schema_param}&ws_key={api_key}'
        # Continúa con el resto de tu código que utiliza url
        print(link_values)
    else:
        print("API no encontrada")

# Ejecutamos la funcion al cambiar de seleccion
combo.bind("<<ComboboxSelected>>", selection_changed)

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
    file = filedialog.askopenfile(mode='r', filetypes=[('all files',"*.*") ])
    # Si un archivo
    if file:
        # Cogemos el valor del directorio y lo metemos en FilePath
        filepath = os.path.abspath(file.name)
        # Le damos el valor del directorio a la variable pathmsg
        pathmsg.set(filepath)

# Variable Mensaje EAN duplicado
doopemsg = StringVar()
# Variable filas global
global filas
# Duplicados
tableelement = []

global duplicados
duplicados = []
# Elemento sin duplicar
ids = []
codigoarticulo = []
descripcion = []
talla = []
codbarras2 = [] 

filas = " "

# Analizamos el fichero
def Analizar():
    global duplicados, filas, df, fila, df_output
    # Leyendo el excel
    df = pd.read_excel(filepath, header = 0, dtype='object')
    df_output = pd.read_excel(filepath, header = 0, dtype='object')
    dftable = pd.read_excel(filepath, header = 0, usecols = "A, B, C, H", dtype='object')
    # Guardando los datos del Excel
    fila = df.iloc[7] 
    filas = df.shape[0]
    codigoarticulo = []
    descripcion = []
    talla = []
    codbarras2 = []

    # Pidiendo lo sdatos de la API de Prestashop
    response = requests.get(url, params=schema_param)
    content = response.content
    doc_json = xmltodict.parse(content, )
    json_data = json.dumps(doc_json)
    json_dict = json.loads(json_data)

    # Cogemos los datos de los EAN de la API de Prestashop
    ean = json_dict['prestashop']['combinations']['combination']
    duplicados = []
    noduplicados = []
    # Recorremos las filas
    for i in range(filas -1):     
        codigoarticulo.append(df.iloc[i][0])
        descripcion.append(df.iloc[i][1])
        talla.append(df.iloc[i][2])
        codbarras2.append(df.iloc[i][7])
        # Recorremos los EAN
        for eanAux in ean:
            ean13 = eanAux.get('ean13')
            if ean13 is not None:
                if ean13 != df.iloc[i][7]:
                    noduplicados.append(ean13)
                else:
                    duplicados.append(ean13)
                    ids.append(df.iloc[i][0])
            if i == 0 or i == filas -1:
                continue
    if duplicados:
        doopemsg.set(duplicados)
    else:
        doopemsg.set("Sin duplicados")
    tableelement.append(duplicados)
    for codigoarticulo, descripcion, talla, codbarras2 in zip(codigoarticulo, descripcion, talla, codbarras2):
        table.insert("", "end", values=(codigoarticulo, descripcion, talla, codbarras2))
    return

def eliminarduplicado():
    global df, table
    if df is None:
        print("No se ha cargado ningún archivo XLSX.")
        return

    if 'Cód Barras 2' not in df.columns:
        print("La columna 'Cód Barras 2' no existe en el DataFrame.")
        return

    # Filtrar el DataFrame para eliminar las filas que contienen códigos de barras duplicados en PrestaShop
    df = df[~df['Cód Barras 2'].isin(duplicados)]
    output_filepath = os.path.splitext(filepath)[0] + "_sin_duplicados.xlsx"
    
    # Guardar el DataFrame actualizado en el archivo XLSX original
    df.to_excel(output_filepath, engine='xlsxwriter')
    print(filepath)
    print("Duplicados eliminados.")
    print("Nuevo tamaño del DataFrame:", df.shape[0])
    print("El archivo ha sido actualizado.")

    # Limpiar la tabla antes de mostrar los datos actualizados
    for item in table.get_children():
        table.delete(item)
    Analizar()


# Botones para Buscar, Analizar y Eliminar duplicados
Button(tool_bar,  text="Buscar", command=Directorio,  relief=RAISED).grid(row=0,  column=2,  padx=10,  pady=3)
Button(frame2,  text="Analizar", command=Analizar,  relief=RAISED).grid(row=0,  column=0,  padx=2,  pady=0)
Button(frame2,  text="Eliminar duplicados", command=eliminarduplicado,  relief=RAISED).grid(row=0,  column=1,  padx=2,  pady=0)

# Tabla con los datos
table = ttk.Treeview(win, columns=('#1', '#2', '#3', '#4'), selectmode=EXTENDED, show='headings')
table.grid(row=1, column=1, sticky="WE", padx=20)

# Asignar encabezados de columna
table.heading('#1', text='Código Artículo')
table.column("#1", anchor=CENTER, stretch=YES,width=100)
table.heading('#2', text='Descripción')
table.column("#2", anchor=CENTER, stretch=YES, width=100)
table.heading('#3', text='Talla')
table.column("#3", anchor=CENTER, stretch=YES, width=100)
table.heading('#4', text='Cód Barras 2')
table.column("#4", anchor=CENTER, stretch=YES, width=100)



# Marca de agua Juan 2023
Label(win,  text="Hecho por Juanies#6389 - 2023",  relief=FLAT).grid(row=3,  column=1,  padx=0,  pady=20)


win.mainloop()