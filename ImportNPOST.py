import tkinter as tk
from tkinter import ttk, filedialog, messagebox,font
import pandas as pd
import base_datos as db
from PIL import Image, ImageTk
import os
from datetime import datetime


def centrar_ventana(root, ancho, alto):
    ancho_pantalla = root.winfo_screenwidth()
    alto_pantalla = root.winfo_screenheight()
    x = (ancho_pantalla // 2) - (ancho // 2)
    y = (alto_pantalla // 2) - (alto // 2)
    root.geometry(f'{ancho}x{alto}+{x}+{y}')
    
def registrar_error(texto_error, archivo_log=r'validar/appImport_log.txt'):
    try:
        # La fecha y hora del error, importante
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        # Formatear el mensaje de error, por si acaso
        mensaje_error = f"{timestamp} - ERROR - {texto_error}\n"
        # Escribir el mensaje de error en el archivo
        with open(archivo_log, 'a') as file:
            file.write(mensaje_error)
    except Exception as e:
        # En caso de error al registrar el error, imprimir el error en la consola
        print(f"Error al registrar el error: {str(e)}")

# Crear la ventana principal
root = tk.Tk()
root.title("NemeSys ImportNPOST")
ancho_ventana = 750
alto_ventana = 520
centrar_ventana(root, ancho_ventana, alto_ventana)

# Crear el Notebook (pestañas)
notebook = ttk.Notebook(root)
notebook.pack(fill=tk.BOTH, expand=True)
# Crear la pestaña de Cliente
frame_cliente = ttk.Frame(notebook)
notebook.add(frame_cliente, text='Cliente')
# Función para examinar archivo
def leer_columnascliente(ruta_archivo, listbox):
    df = pd.read_excel(ruta_archivo)  # Lee el archivo Excel
    columnas = df.columns.tolist()  # Obtiene la lista de columnas
    listbox.delete(0, tk.END)  # Borra cualquier contenido previo en el Listbox
    for col in columnas:
        listbox.insert(tk.END, col) 

def examinar_archivo():
    archivo = filedialog.askopenfilename(title="Seleccionar archivo", filetypes=[("Archivos de Excel", "*.xlsx *.xls")])
    if archivo:
        leer_columnascliente(archivo, listaexcel)
        entry_ruta.delete(0, tk.END)  # Borra cualquier contenido previo en el Entry
        entry_ruta.insert(0, archivo)  # Inserta la ruta del archivo en el Entry
        print(f"Archivo seleccionado: {archivo}")
        
#columna excel listbox solo observacion
groupboxexcel = tk.LabelFrame(frame_cliente, text="Campos Excel ", padx=1, pady=1)
groupboxexcel.place(x=295, y=93, width=100, height=240)
listaexcel = tk.Listbox(groupboxexcel)
listaexcel.place(x=5, y=20, width=80,height=190)

inicioframe = tk.LabelFrame(frame_cliente, padx=1, pady=1)
inicioframe.place(x=20, y=30, width=690, height=50)

# Crear el campo de entrada para mostrar la ruta del archivo
entry_ruta = tk.Entry(inicioframe)
entry_ruta.place(x=20, y=10, width=540, height=25)

# Crear el botón de examinar archivo
btn_examinar = tk.Button(inicioframe, width=12, height=1, text="Examinar", command=examinar_archivo)
btn_examinar.place(x=570, y=10)

def cargar_archivo_Cliente():
    datos = list(columnas_Clientes.values())
    for datos in datos:
        listacolumna1.insert(tk.END, datos)

timestamp = datetime.now().strftime('%m/%d/%y')
hora = datetime.now().strftime('%I:%M:%S %p')
CHEQUES = str(1)
TARJETA = str(1)
EFECTIVO = str(1)
ESTATUS_DEL = str(0)
PERM_VENTAS = 3
USUARIO_ID = 1
LIMIT_CRED = 0
DESCUENTO = 0
CRED_TIENDA = 0
NIVEL_PRECIO = 1
MONEDA = 'EFECTIVO'

columnas_Clientes = {
    "NOMBRE": "Nombre",
    "APELLIDO": "Apellido",
    "CEDULA": "Cedula\Rnc",
    "SEXO": "Sexo",
    "TELEFONO1": "Telefono",
}

# Lista de posiciones de las columnas
posiciones_cliente = []

def derechalistCliente():
    seleccion = listacolumna1.curselection()
    if seleccion:
        valor_seleccionado = listacolumna1.get(seleccion)
        listacolumna2.insert(tk.END, valor_seleccionado)
        listacolumna2.select_set(tk.END)
        listacolumna1.delete(seleccion)

        for key, value in columnas_Clientes.items():
            if value == valor_seleccionado:
                posiciones_cliente.append(key)
                break

def izquierdalistCliente():
    seleccion = listacolumna2.curselection()
    if seleccion:
        valor_seleccionado = listacolumna2.get(seleccion)
        listacolumna1.insert(tk.END, valor_seleccionado)
        listacolumna1.select_set(tk.END)
        listacolumna2.delete(seleccion)

        for key, value in columnas_Clientes.items():
            if value == valor_seleccionado:
                posiciones_cliente.remove(key)
                break

def limpiarCliente():
    listacolumna1.delete(0, tk.END)
    listacolumna2.delete(0, tk.END)
    posiciones_cliente.clear()
    entry_ruta.delete(0, tk.END)
    combobox1nfc.set('')
    combobox2Cond.set('')
    cargar_archivo_Cliente()




def cargar_archivo_Cliente():
    for key in columnas_Clientes.keys():
        listacolumna1.insert(tk.END, columnas_Clientes[key])
posicionexcel = None
# Función para insertar cliente
def insertarCliente(columnas, posiciones_cliente,):
    global posicionexcel
    posicionexcel += 1
    try:
        # Extraer las columnas de acuerdo a las posiciones dadas
        nombre = str(columnas[posiciones_cliente.index("NOMBRE")]).upper() if "NOMBRE" in posiciones_cliente and columnas[posiciones_cliente.index("NOMBRE")] else None
        apellido = str(columnas[posiciones_cliente.index("APELLIDO")]).upper() if "APELLIDO" in posiciones_cliente and columnas[posiciones_cliente.index("APELLIDO")] else None
        cedula = str(columnas[posiciones_cliente.index("CEDULA")]) if "CEDULA" in posiciones_cliente and columnas[posiciones_cliente.index("CEDULA")] else None
        sexo = str(columnas[posiciones_cliente.index("SEXO")]).upper() if "SEXO" in posiciones_cliente and columnas[posiciones_cliente.index("SEXO")] else None

        # Validar longitud de la cédula
        if cedula and not (9 <= len(cedula) <= 11):
            raise ValueError(f"Cédula inválida: {cedula}")

        # Asegurar que el teléfono se formatee correctamente
        telefono = str(columnas[posiciones_cliente.index("TELEFONO1")]) if "TELEFONO1" in posiciones_cliente and columnas[posiciones_cliente.index("TELEFONO1")] else None
        if telefono and len(telefono) == 10:  # Asumiendo que el teléfono debe tener 10 dígitos
            telefono_formateado = f"{telefono[:3]}-{telefono[3:6]}-{telefono[6:]}"
        else:
            telefono_formateado = "000-000-0000"

        print(nombre, apellido, cedula, sexo, telefono_formateado)

        cursor = db.conexiondb.cursor()
        query = """INSERT INTO CLIENTES (NOMBRE, APELLIDO, CEDULA, SEXO, TELEFONO1, MONEDA,
                                         FECHA, FECHA_INGRESO, HORA, CHEQUES, TARJETA, EFECTIVO,
                                         ESTATUS_DEL, USUARIO_ID, PERM_VENTAS, LIMIT_CRED, DESCUENTO, CRED_TIENDA, NIVEL_PRECIO, CONDICIONES_PAGO, NCF)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"""
        valores = [nombre, apellido, cedula, sexo, telefono_formateado, MONEDA, timestamp, timestamp, hora,
                   CHEQUES, TARJETA, EFECTIVO, ESTATUS_DEL, USUARIO_ID, PERM_VENTAS, LIMIT_CRED, DESCUENTO, CRED_TIENDA, NIVEL_PRECIO, result_condiciones, result_nfc]
        cursor.execute(query, valores)
        db.conexiondb.commit()
        print("Cliente insertado:", valores)

    except ValueError as ve:
        registrar_error(str(ve))
        global insertado
        insertado -= 1
        messagebox.showwarning("Error Cedula O Rnc Invaladio \n", str(ve),"\n")
        print(f"Cliente no insertado debido a error en cédula: {ve}")

    except Exception as e:
        registrar_error(str(e))
        print(f"Error al insertar cliente: {e}")

def actualizar_estado():
    global insertado 
    info=db.cantidadcliente()
    informacioncliente.config(text=info)

# Función para leer el archivo y procesar clientes
insertado=None

def insertcliente():
    iniciar()
    global insertado
    ruta = entry_ruta.get()
    if not ruta:
        print("No se ha seleccionado un archivo")
        return
    
    try:
        df = pd.read_excel(ruta, header=None, skiprows=1)
        
        #if df.shape[1] < 5:
        #    print("El archivo no tiene suficientes columnas")
        #    return
        
        insertado = 0
        for idx, row in df.iterrows():
            try:
                row = row.fillna(0).replace({pd.NA: 0, pd.NaT: 0, None: 0, 'nan': 0})
                columnas = row.tolist()
                str(columnas)
                str(posiciones_cliente)
                iniciar()
                insertarCliente(columnas, posiciones_cliente)
                print("Cliente insertado con éxito")
                insertado += 1
            except Exception as e:
                registrar_error(e)
                messagebox.showinfo("Error",f"Error al insertar cliente en la fila {idx}: {e}")
        
        messagebox.showinfo("Registro Cliente", f"Carga Exitosa \n cantidad insertada: {insertado}")
        
        detener()
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
    finally:
        actualizar_estado()
# Dummy call to display the lists        
groupboxlista = tk.LabelFrame(frame_cliente, text="Selecciones Columna de Excel ", padx=1, pady=1)
groupboxlista.place(x=400, y=93, width=310, height=240)
listacolumna1 = tk.Listbox(groupboxlista)
listacolumna1.place(x=5, y=20, width=100,height=190)
cargar_archivo_Cliente()
listacolumna2 = tk.Listbox(groupboxlista)
listacolumna2.place(x=200, y=20, width=100,height=190)

#CREACION COMBOBOX DE TRAKINGNCF,CONDICIONES_A_PAGAR
result_nfc = None
def on_combobox_select(event):
    global result_nfc
    selected_nfc = combobox1nfc.get()
    cursor = db.conexiondb.cursor()
    cursor.execute("SELECT NOMBRE FROM TRAKINGNCF WHERE NOMBRE = ?", (selected_nfc,))
    result = cursor.fetchone()
    if result is not None:
        result_nfc = result[0]
        print(result_nfc)
    else:
        print("No se encontró ningún resultado.")
        result_nfc = None
        
result_condiciones = None
def on_combobox_select_condiciones(event):
    global result_condiciones
    selected_nfc = combobox2Cond.get()
    cursor = db.conexiondb.cursor()
    cursor.execute("SELECT  codiciones FROM CONDICIONES_A_PAGAR where codiciones= ?", (selected_nfc,))
    result = cursor.fetchone()
    if result is not None:
        result_condiciones = result[0]
        print(result_condiciones)
    else:
        print("No se encontró ningún resultado.")
        result_condiciones = None

# Crear el Label dentro del GroupBox
groupbox = tk.LabelFrame(frame_cliente, text="Selecciones las Opciones ",padx=1, pady=1)
groupbox.place(x=20, y=93, width=270, height=240)
labeltrank = tk.Label(groupbox, text="NCF",font=("Arial", 12,"bold"))
labeltrank.place(x=6, y=20)
# Crear el primer ComboBox dentro del GroupBox
combobox1nfc = ttk.Combobox(groupbox, values=db.TRAKINGNCF())
combobox1nfc.place(x=60, y=20,width=200)
combobox1nfc.bind("<<ComboboxSelected>>", on_combobox_select)
# Crear el segundo ComboBox dentro del GroupBox
labelcondiciones = tk.Label(groupbox, text="Condiciones",font=("Arial", 12,"bold"))
labelcondiciones.place(x=5, y=50)
combobox2Cond = ttk.Combobox(groupbox, values=db.Condiciones())
combobox2Cond.place(x=120, y=50,width=140)
combobox2Cond.bind("<<ComboboxSelected>>", on_combobox_select_condiciones)
# Crear el tercero ComboBox dentro del GroupBox
labeltipo = tk.Label(groupbox, text="Tipo",font=("Arial", 12,"bold"))
labeltipo.place(x=6, y=80)
comboboxtipo = ttk.Combobox(groupbox, values='')
comboboxtipo.place(x=120, y=80,width=140)


imagenes ={}
# Imagen de los botones
def cargar_imagen(nombre, ruta):
    imagen = Image.open(ruta)
    imagen = imagen.resize((32, 32), Image.LANCZOS)
    imagen_tk = ImageTk.PhotoImage(imagen)
    imagenes[nombre] = imagen_tk  # Guardar referencia en el diccionario
    return imagen_tk


    

btnigual = cargar_imagen("igual", "Icono/igual.ico")
btncargar1 = cargar_imagen("cargar1", "Icono/CARGAR.ico")
btnlef = cargar_imagen("lef", "Icono/agregar.ico")
btnizq = cargar_imagen("izq", "Icono/right.ico")
btncerrar = cargar_imagen("cerrar", "Icono/CERRAR.ico")
btnlimpiar = cargar_imagen("limpiar", "Icono/BORRAR.ico")

# Botones para vincular información groupboxlista
btnderechacliente = tk.Button(groupboxlista, image=btnlef,command=derechalistCliente)
btnderechacliente.place(x=110, y=80, width=85,height=35)
btnizquierdacliente = tk.Button(groupboxlista, image=btnizq,command=izquierdalistCliente)
btnizquierdacliente.place(x=110, y=120,  width=85,height=35)


#botones de limpieza
botonfuncioncliente = tk.LabelFrame(frame_cliente, padx=1, pady=1)
botonfuncioncliente.place(x=400, y=348, width=310, height=45)
btnlimpiarcliente = tk.Button(botonfuncioncliente, image=btnlimpiar, command=limpiarCliente)
btnlimpiarcliente.place(x=10, y=1, width=90, height=35)
btncerrarcliente = tk.Button(botonfuncioncliente, image=btncerrar, command=frame_cliente.quit)
btncerrarcliente.place(x=110, y=1, width=90, height=35)
btncargarcliente = tk.Button(botonfuncioncliente, image=btncargar1, command=insertcliente)
btncargarcliente.place(x=210, y=1, width=90, height=35)
#mostrar informacion de la base datos tabla cliente
informacionBd = tk.LabelFrame(frame_cliente, padx=0, pady=0,)
informacionBd.place(x=20, y=348, width=270, height=45)
cantidadClienrte = tk.Label(informacionBd, text="Cantidad Cliente", font=("Arial", 9, "bold"))
cantidadClienrte.place(x=2, y=0)  
info=db.cantidadcliente()
informacioncliente=tk.Label(informacionBd,text=info,font=("Arial", 9))
informacioncliente.place(x=100,y=0,width=40)
estatuinf = tk.Label(informacionBd, text="Estatus", font=("Arial", 9, "bold"))
estatuinf.place(x=2, y=20)  
progressbar = ttk.Progressbar(informacionBd, orient=tk.HORIZONTAL, length=154, mode='indeterminate')

# Función para iniciar la barra de progreso
def iniciar():
    progressbar.start()
    progressbar.place(x=112, y=20)  
# Función para detener la barra de progreso
def detener():
    progressbar.stop()
    progressbar.place_forget()  # Ocultar la barra de progreso

# Función para ocultar la barra de progreso
def ocultar():
    progressbar.place_forget()


#inicio de la ventana inventario
# Crear la pestaña de Inventario
frame_inventario = ttk.Frame(notebook)
notebook.add(frame_inventario, text='Inventario')
# Mover aquí el código que estaba en la ventana principal a la pestaña de Inventario
def examinar_archivo_inventario():
    archivo = filedialog.askopenfilename(title="Seleccionar archivo", filetypes=[("Archivos de Excel", "*.xlsx *.xls")])
    if archivo:
        entrada_archivo.delete(0, tk.END)
        entrada_archivo.insert(0, archivo)

lf_archivo = ttk.LabelFrame(frame_inventario)
lf_archivo.grid(row=0, column=0, columnspan=3, padx=10, pady=5, sticky="ew")

entrada_archivo = ttk.Entry(lf_archivo, width=85, font=('Arial', 9))
entrada_archivo.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

boton_examinar = ttk.Button(lf_archivo, text="Examinar", command=examinar_archivo_inventario, width=15)
boton_examinar.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

lf_combos = ttk.LabelFrame(frame_inventario, text="")
lf_combos.grid(row=2, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
comboboxes = {}

for i, (label_text, combo_values) in enumerate([
    ("Departamento", db.Departamento_id()),
    ("Proveedor", db.Proveedor_id()),
    ("Impuesto Ventas *", db.Impuesto()),
    ("Impuesto Compra *", db.Impuesto()),
    ("Localidad *", db.Localidad()),
    ("Unidad *", db.UnidadMedida())
]):
    label = ttk.Label(lf_combos, text=label_text, font=("Arial", 12, "bold"))
    label.grid(row=i, column=0, padx=5, pady=5, sticky="w")

    if label_text in ["Departamento", "Proveedor"]:
        combo = ttk.Combobox(lf_combos, values=combo_values, width=30)  # Ajusta el width según lo necesario
    else:
        combo = ttk.Combobox(lf_combos, values=combo_values)
    
    combo.grid(row=i, column=1, padx=5, pady=5, sticky="ew")
    comboboxes[label_text] = combo

def obtener_valores():
    return {label: combo.get() for label, combo in comboboxes.items()}

def limpiar_comboboxes():
    for combo in comboboxes.values():
        combo.set('')


style = ttk.Style()
style.configure("TRadiobutton", font=("Arial", 9, "bold"))
lf_min_max = ttk.LabelFrame(frame_inventario, text="Valores")
lf_min_max.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
tipo_articulo = tk.StringVar(value="ARTICULO")
radio_servicio_min_max = ttk.Radiobutton(lf_min_max, text="Servicio", variable=tipo_articulo, value="SERVICIO", style="TRadiobutton")
radio_servicio_min_max.grid(row=0, column=0, padx=5, pady=5, sticky='w')
radio_articulo_min_max = ttk.Radiobutton(lf_min_max, text="Artículo", variable=tipo_articulo, value="ARTICULO", style="TRadiobutton")
radio_articulo_min_max.grid(row=1, column=0, padx=5, pady=5, sticky='w')
label_max = ttk.Label(lf_min_max, text="MAX", font=('Arial', 9, 'bold'))
label_max.grid(row=0, column=1, padx=5, pady=5)
maxv = tk.StringVar(value=100)
entrada_max = ttk.Entry(lf_min_max, width=5, font=('Arial', 9), textvariable=maxv)
entrada_max.grid(row=0, column=2, padx=5, pady=5)
label_min = ttk.Label(lf_min_max, text="MIN", font=('Arial', 9, 'bold'))
label_min.grid(row=1, column=1, padx=5, pady=5)
minx = tk.StringVar(value=1)
entrada_min = ttk.Entry(lf_min_max, width=5, font=('Arial', 9), textvariable=minx)
entrada_min.grid(row=1, column=2, padx=5, pady=5)
columnas_map = {
    "DEPARTAMENTO_ID": "Departamento",
    "PROVEEDOR_ID": "Proveedor",
    "DESC1": "Nombre Corto",
    "DESC2": "Nombre Largo",
    "DESC3": "Otra descripción 3",
    "DESC4": "Otra descripción 4",
    "ATRIB": "Atributo/Color",
    "TALLA": "Talla",
    "ALU": "Cod.Barra (ALU)",
    "UPC": "Cod.Barra (UPC)",
    "COSTO": "Costo",
#    "PRECIO": "Precio est",
    "PRECIO_IMP": "Precio Imp",
    "PRECIO_NIVEL2": "Nivel 2",
    "PRECIO_NIVEL3": "Nivel 3",
    "PRECIO_NIVEL4": "Nivel 4",
    "CODIGO_BARRA1": "Codigo barra 1",
    "CODIGO_BARRA2": "Codigo barra 2",
    "CODIGO_BARRA3": "Codigo barra 3",
    "CODIGO_BARRA4": "Codigo barra 4",
    "CODIGO_BARRA5": "Codigo barra 5"
}
posiciones_originales = []

def cargar_archivo():
    datos = list(columnas_map.values())
    for dato in datos:
        listbox1.insert(tk.END, dato)

def derechalist():
    seleccion = listbox1.curselection()
    if seleccion:
        valor_seleccionado = listbox1.get(seleccion)
        listbox2.insert(tk.END, valor_seleccionado)
        listbox2.select_set(tk.END)
        listbox1.delete(seleccion)

        for key, value in columnas_map.items():
            if value == valor_seleccionado:
                posiciones_originales.append(key)
                break
def izquierdalist():
    seleccion = listbox2.curselection()
    if seleccion:
        valor_seleccionado = listbox2.get(seleccion)
        listbox1.insert(tk.END, valor_seleccionado)
        listbox1.select_set(tk.END)
        listbox2.delete(seleccion)

        for key, value in columnas_map.items():
            if value == valor_seleccionado:
                posiciones_originales.remove(key)
                break
def limpiar():
    listbox1.delete(0, tk.END)
    listbox2.delete(0, tk.END)
    posiciones_originales.clear()
    entrada_archivo.delete(0, tk.END)
    limpiar_comboboxes()
    cargar_archivo()
    
    
#--------funciones posgrebar--------------
def iniciarinvetario():
    progressbarinventario.start()
    progressbarinventario.grid(row=1, column=1, padx=2, pady=2, sticky="ew")
# Función para detener la barra de progreso
def detenerinvetario():
    progressbarinventario.stop()
    progressbarinventario.grid_forget()
# Función para ocultar la barra de progreso
def ocultarinventario():
    progressbarinventario.grid_forget()   

def procesar_excel():
    try:
        iniciarinvetario()
        archivo = entrada_archivo.get()
        df = pd.read_excel(archivo, header=None, skiprows=1)
        columnas_numericas = ["COSTO", "precio", "PRECIO_IMP", "precio_nivel2", "precio_nivel3", "precio_nivel4"]
        for columna in columnas_numericas:
            if columna in df.columns:
                df[columna] = pd.to_numeric(df[columna], errors='coerce')
        cursor = db.conexiondb.cursor()
        cursor.execute("SELECT MAX(ITEM_ID) FROM EXISTENCIAS")
        max_codigo_item = cursor.fetchone()[0]
        codigo_item = max_codigo_item if max_codigo_item is not None else 999
        valores = obtener_valores()
        impuesto_ventas = valores["Impuesto Ventas *"]
        impuesto_compuesto = valores["Impuesto Compra *"]
        cursor.execute("SELECT ID FROM IMPUESTOS WHERE IMPUESTO = ?", (impuesto_ventas,))
        ventas_RES = cursor.fetchone()
        if ventas_RES is None:
            raise ValueError("Impuesto ventas no encontrado en la base de datos.")
        ventas = ventas_RES[0]
        cursor.execute("SELECT RAZON FROM IMPUESTOS WHERE IMPUESTO = ?", (impuesto_compuesto,))
        razon_re = cursor.fetchone()
        if razon_re is None:
            raise ValueError("Impuesto compra no encontrado en la base de datos.")
        razon = razon_re[0]
        impuesto_id = str(ventas)
        impuesto_porc = float(razon)
        impuesto_compra_id = int(ventas)
        iimpuesto_compra_porc = float(razon)
        localidad = valores["Localidad *"]
        unidad = valores["Unidad *"]
        tipoart = tipo_articulo.get()
        valormax = entrada_max.get()
        valormin = entrada_min.get()
        usID = '1'
        us = '1'
        estatus_del = '0'
        color = 'clWindow'
        columns = ','.join(posiciones_originales)
        placeholders = ','.join(['?' for _ in range(len(posiciones_originales))])
        cursor.execute("SELECT siglas FROM UNIDADES_MEDIDAS where descripcion = ?", (unidad,))
        und_res = cursor.fetchone()
        if und_res is None:
            raise ValueError("Unidad no encontrada en la base de datos.")
        und = und_res[0]
        departamento = valores.get("Departamento")
        proveedor = valores.get("Proveedor")
        cursor.execute("SELECT ID FROM DEPARTAMENTO WHERE NOMBRE_DEPARTAMENTO = ?", (departamento,))
        dep_res = cursor.fetchone()
        if dep_res is None:
            dep = None
        else:
            dep = dep_res[0]
        cursor.execute("SELECT ID FROM PROVEEDOR WHERE EMPRESA = ?", (proveedor,))
        pro_res = cursor.fetchone()
        if pro_res is None:
            pro = None
        else:
            pro = pro_res[0]
        departamentocombo="DEPARTAMENTO_ID"
        Proveedorcombo="PROVEEDOR_ID"
        df = procesar_columnas(df)
        
        if dep and pro is not None:
            valordep = str(dep)
            valorpro = str(pro)
            for idx in range(listbox1.size()):
                item = listbox1.get(idx)
                if item in ["Departamento", "Proveedor"]:
                    listbox1.itemconfig(idx, {'foreground': 'red'})
                else:
                    listbox1.itemconfig(idx, {'foreground': 'black'})
            indice_precio_imp = posiciones_originales.index("PRECIO_IMP")
            # Verificar las condiciones
            
            query = f"""
                        INSERT INTO INVENTARIO(
                        CODIGO_ITEM,
                        TIPO, 
                        CANT_MAX, 
                        CANT_MIN,
                        UND_MEDIDA,
                        LOCALIDAD,
                        USUARIO,
                        USUARIO_ID_MOD,
                        ESTATUS_DEL,
                        IMPUESTO_ID,
                        IMPUESTO_PORC,
                        FECHA_REG,
                        IMPUESTO_COMPRA_ID,
                        IIMPUESTO_COMPRA_PORC,
                        MOSTRAR_EN_COLOR,
                        ONDETAIL,
                        PRECIO,
                        {departamentocombo},
                        {Proveedorcombo},
                        {columns}) 
                        VALUES 
                        (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,{placeholders})"""
            timestamp = datetime.now().strftime('%m/%d/%y')
            insertados = 0
            for idx, row in df.iterrows():
                ondetail = db.ondetail()
                try:
                    actualizar_estadoinventario()
                    row = row.fillna(0).replace({pd.NA: 0, pd.NaT: 0, None: 0, 'nan': 0})
                    codigo_item += 1

                    if 'PRECIO_IMP' in posiciones_originales:
                        indice_precio_imp = posiciones_originales.index('PRECIO_IMP')
                        row[indice_precio_imp] = float(row[indice_precio_imp])
                        precioestadarvalor = row[indice_precio_imp]
                    else:
                        precioestadarvalor = 0  # Valor predeterminado si la columna no está presente

                    descuentoimp = (impuesto_porc / 100) + 1
                    precioest = precioestadarvalor / descuentoimp if precioestadarvalor is not None else 0

                    datos_fila = [
                        codigo_item, tipoart, 
                        int(valormax), int(valormin), 
                        und, localidad, 
                        usID, us, estatus_del, 
                        impuesto_id, impuesto_porc, timestamp, 
                        impuesto_compra_id, iimpuesto_compra_porc, 
                        color, ondetail, float(precioest), valordep, valorpro,
                    ]
                    datos_fila += list(row[:len(posiciones_originales)])

                    cursor.execute(query, datos_fila)
                    print(query, "\n", datos_fila)
                    insertados += 1

                except Exception as e:
                    if 'SQL error code = -303' in str(e):
                        print(f"Error de conversión detectado en la fila {idx+1}, valor problemático reemplazado por 0.")
                        registrar_error(e)
                        row = row.replace({pd.NA: 0, pd.NaT: 0, None: 0, 'nan': 0})
                        datos_fila = [
                            codigo_item, tipoart, 
                            int(valormax), int(valormin), 
                            und, localidad, 
                            usID, us, estatus_del, 
                            impuesto_id, impuesto_porc, timestamp, 
                            impuesto_compra_id, iimpuesto_compra_porc, 
                            color, ondetail, float(precioest), valordep, valorpro,
                        ]
                        datos_fila += list(row[:len(posiciones_originales)])
                        try:
                            cursor.execute(query, datos_fila)
                            print(query, "\n", datos_fila)
                            insertados += 1
                        except Exception as e:
                            print(f"Se cometió un error al reintentar la fila: {str(e)}")
                            registrar_error(e)
                            continue
                    else:
                        print(f"Se cometió un error al insertar la fila: {str(e)}")
                        registrar_error(e)
                        messagebox.showinfo("Error al Cargar la información", f"Validar el archivo en la línea {str(idx)}, \n Error captado es {str(e)}")
                        continue

            db.conexiondb.commit()
            db.secuenciaInventarioUpdate()
            db.secuenciasUpdate()
            messagebox.showinfo("Proceso Completo", f"Los datos se han insertado correctamente en la base de datos \n cantidad de registros insertados {str(insertados)}.")
            detenerinvetario()
        else:
            query = f"""
                    INSERT INTO INVENTARIO(
                    CODIGO_ITEM,
                    TIPO, 
                    CANT_MAX, 
                    CANT_MIN,
                    UND_MEDIDA,
                    LOCALIDAD,
                    USUARIO,
                    USUARIO_ID_MOD,
                    ESTATUS_DEL,
                    IMPUESTO_ID,
                    IMPUESTO_PORC,
                    FECHA_REG,
                    IMPUESTO_COMPRA_ID,
                    IIMPUESTO_COMPRA_PORC,
                    MOSTRAR_EN_COLOR,
                    ONDETAIL,
                    PRECIO,
                    {columns}) 
                    VALUES 
                    (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,{placeholders})"""
            timestamp = datetime.now().strftime('%m/%d/%y')
            insertados = 0
            for idx, row in df.iterrows():
                ondetail = db.ondetail()
                try:
                    row = row.fillna(0).replace({pd.NA: 0, pd.NaT: 0, None: 0, 'nan': 0})
                    codigo_item += 1
                    indice_precio_imp = posiciones_originales.index('PRECIO_IMP')
                    row[indice_precio_imp] = float(row[indice_precio_imp])
                    valorestadar=row[indice_precio_imp] 
                    descuentoimp = (impuesto_porc / 100) + 1
                    preciovalor = valorestadar / descuentoimp if valorestadar is not None else 0
                    
                    datos_fila = [
                        codigo_item, tipoart, 
                        int(valormax), int(valormin), 
                        und, localidad, 
                        usID, us, estatus_del, 
                        impuesto_id, impuesto_porc, timestamp, 
                        impuesto_compra_id, iimpuesto_compra_porc, 
                        color, ondetail,float(preciovalor)
                    ]
                    datos_fila += list(row[:len(posiciones_originales)])
                    cursor.execute(query, datos_fila)
                    print(query,"\n",datos_fila)
                    
                    insertados += 1
                    
                except Exception as e:
                    if 'SQL error code = -303' in str(e):
                        print(f"Error de conversión detectado en la fila {idx+1}, valor problemático reemplazado por 0.")
                        registrar_error(e)
                        row = row.replace({pd.NA: 0, pd.NaT: 0, None: 0, 'nan': 0})
                        datos_fila = [
                            codigo_item, tipoart, 
                            int(valormax), int(valormin), 
                            und, localidad, 
                            usID, us, estatus_del, 
                            impuesto_id, impuesto_porc, timestamp, 
                            impuesto_compra_id, iimpuesto_compra_porc, 
                            color, ondetail,float(preciovalor)
                        ]

                        datos_fila += list(row[:len(posiciones_originales)])
                        try:
                            cursor.execute(query, datos_fila)
                            print(query,"\n",datos_fila)
                            insertados += 1
                        except Exception as e:
                            print(f"Se cometió un error al reintentar la fila: {str(e)}")
                            registrar_error(e)
                            continue
                    else:
                        print(f"Se cometió un error al insertar la fila: {str(e)}")
                        registrar_error(e)
                        messagebox.showinfo("Error al Cargar la información", f"Validar el archivo en la línea {str(idx)}, \n Error captado es {str(e)}")
                        continue

            db.conexiondb.commit()
            db.secuenciaInventarioUpdate()
            db.secuenciasUpdate()

            messagebox.showinfo("Proceso Completo", f"Los datos se han insertado correctamente en la base de datos \n cantidad de registros insertados {str(insertados)}.")
            detenerinvetario()
    except Exception as e:
        messagebox.showerror("Error", f"Se ha producido un error al procesar el archivo en la funcion Procesar Excel: {str(e)}")
        registrar_error(e)
    finally:
        print("Datos insertados")
        db.inv_Unidad_medidad()
        db.Existencia()
        actualizar_estadoinventario()
        db.DepartamentoIdVacio()



def procesar_columnas(df):
    try:
        for i, row in df.iterrows():
            # Reemplazar valores NaN con valores por defecto en cada fila antes de la inserción
            row = row.fillna({'PROVEEDOR_ID': 'DEFAULT PROVEEDOR', 'DEPARTAMENTO_ID': 'DEFAULT DEPARTAMENTO'})
            # Verificar y obtener los valores de proveedor y departamento
            proveedor_nombre = 'DEFAULT PROVEEDOR'
            departamento_nombre = 'DEFAULT DEPARTAMENTO'
            try:
                if 'PROVEEDOR_ID' in posiciones_originales:
                    proveedor_nombre = row[posiciones_originales.index('PROVEEDOR_ID')]
                if 'DEPARTAMENTO_ID' in posiciones_originales:
                    departamento_nombre = row[posiciones_originales.index('DEPARTAMENTO_ID')]
            except KeyError as e:
                print(f"KeyError: {str(e)}")
                registrar_error(e)
                continue
            departamento_nombre.upper()
            proveedor_nombre.upper()
            # Insertar el proveedor y departamento si no existen y obtener sus IDs
            proveedor_id = db.InsertProvedor(proveedor_nombre) if proveedor_nombre != 'DEFAULT PROVEEDOR' else None
            departamento_id = db.InsertDepartamento(departamento_nombre) if departamento_nombre != 'DEFAULT DEPARTAMENTO' else None
            # Reemplazar los valores en el DataFrame
            if proveedor_id is not None:
                df.at[i, posiciones_originales.index('PROVEEDOR_ID')] = proveedor_id
            if departamento_id is not None:
                df.at[i, posiciones_originales.index('DEPARTAMENTO_ID')] = departamento_id
        
        df = df.astype(str)
        return df
    except Exception as e:
        print(f"Error al procesar las columnas: {str(e)}")
        registrar_error(e)
        return df

lf_listbox = ttk.LabelFrame(frame_inventario, text="Elemento a Insertar")
lf_listbox.grid(row=2, column=2, columnspan=2, rowspan=2, padx=10, pady=5, sticky="nsew") 
listbox1 = tk.Listbox(lf_listbox)
listbox1.grid(row=0, column=0, padx=7, pady=5, sticky="nsew")
cargar_archivo()
listbox2 = tk.Listbox(lf_listbox)
listbox2.grid(row=0, column=2, padx=7, pady=5, sticky="nsew")
btnlef = Image.open("Icono/agregar.ico")
btnlef = btnlef.resize((32, 32), Image.LANCZOS)
btnlef = ImageTk.PhotoImage(btnlef)
btnizq = Image.open("Icono/right.ico")
btnizq = btnizq.resize((32, 32), Image.LANCZOS)
btnizq = ImageTk.PhotoImage(btnizq)
btncerrar = Image.open("Icono/CERRAR.ico")
btncerrar = btncerrar.resize((32, 32), Image.LANCZOS)
btncerrar = ImageTk.PhotoImage(btncerrar)
btnlimpiar = Image.open("Icono/BORRAR.ico")
btnlimpiar = btnlimpiar.resize((32, 32), Image.LANCZOS)
btnlimpiar = ImageTk.PhotoImage(btnlimpiar)
btncargar = Image.open("Icono/CARGAR.ico")
btncargar = btncargar.resize((32, 32), Image.LANCZOS)
btncargar = ImageTk.PhotoImage(btncargar)

boton_frame = ttk.Frame(lf_listbox)
boton_frame.grid(row=0, column=1, padx=5, pady=5, sticky="ns")
lf_listbox.columnconfigure(0, weight=1)
lf_listbox.columnconfigure(2, weight=1)
lf_listbox.rowconfigure(0, weight=1)
boton_agregar = tk.Button(boton_frame, image=btnlef, command=derechalist)
boton_agregar.grid(row=0, column=0, pady=5)
boton_eliminar = tk.Button(boton_frame, image=btnizq, command=izquierdalist)
boton_eliminar.grid(row=1, column=0, pady=5)


def actualizar_estadoinventario():
    # Consulta a la base de datos
    inventario = db.conexiondb.cursor()
    inventario.execute("select count(*) from INVENTARIO")
    res = inventario.fetchall()[0][0]
    mostrarinventario = str(res)
    # Actualizar el texto del label
    resulinventario.config(text=mostrarinventario)    

#botones de la parte derechas
lf_botones_adicionales = ttk.LabelFrame(frame_inventario, text="")
lf_botones_adicionales.grid(row=4, column=2, columnspan=2, padx=10, pady=5, sticky="ew")
lf_botones_adicionales.columnconfigure(0, weight=1)
lf_botones_adicionales.columnconfigure(1, weight=1)
lf_botones_adicionales.columnconfigure(2, weight=1)
lf_botones_adicionales.rowconfigure(0, weight=1)
boton_limpiar = tk.Button(lf_botones_adicionales, image=btnlimpiar, command=limpiar)
boton_limpiar.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
boton_cerrar = tk.Button(lf_botones_adicionales, image=btncerrar, command=root.destroy)
boton_cerrar.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
boton_cargar = tk.Button(lf_botones_adicionales, image=btncargar, command=procesar_excel)
boton_cargar.grid(row=0, column=2, padx=5, pady=5, sticky="ew")
frame_inventario.columnconfigure(0, weight=1)
frame_inventario.columnconfigure(1, weight=1)
frame_inventario.columnconfigure(2, weight=1)
frame_inventario.columnconfigure(3, weight=1)
frame_inventario.rowconfigure(4, weight=1)
# Crear LabelFrame a la izquierda de lf_botones_adicionales
lf_botones_izquierda = ttk.LabelFrame(frame_inventario, text="")
lf_botones_izquierda.grid(row=4, column=0, columnspan=2, padx=10, pady=(0, 2), sticky="ew")
lf_botones_izquierda.columnconfigure(0, weight=1)
lf_botones_izquierda.columnconfigure(1, weight=1)
lf_botones_izquierda.columnconfigure(2, weight=1)
lf_botones_izquierda.rowconfigure(0, weight=1)
lf_botones_izquierda.rowconfigure(1, weight=1)


#consulta de cantidad de inventario
# Añadir contenido de proveedor al LabelFrame de la izquierda
caninventario = tk.Label(lf_botones_izquierda, text="Cantidad inventario", font=("Arial", 9, "bold"))
caninventario.grid(row=0, column=0, padx=2, pady=2, sticky="w")
inventario = db.conexiondb.cursor()
inventario.execute("select count(*) from INVENTARIO")
res = inventario.fetchall()[0][0]
mostrarinventario = str(res)
resulinventario=tk.Label(lf_botones_izquierda,text=mostrarinventario,font=("Arial", 9, "bold"))
resulinventario.grid(row=0, column=1, padx=2, pady=2, sticky="w")
estatuinventario = tk.Label(lf_botones_izquierda, text="Estatus", font=("Arial", 9, "bold"))
estatuinventario.grid(row=1, column=0, padx=2, pady=2, sticky="w")
progressbarinventario= ttk.Progressbar(lf_botones_izquierda, orient=tk.HORIZONTAL, length=200, mode='indeterminate')
# Crear LabelFrame a la derecha para los botones adicionales
lf_botones_adicionales.columnconfigure(0, weight=1)
lf_botones_adicionales.columnconfigure(1, weight=1)
lf_botones_adicionales.columnconfigure(2, weight=1)
lf_botones_adicionales.rowconfigure(0, weight=1)
#fin de la ventana inventario



#----VENTANA PROVEEDOR ----------------------------------------------------------------------
Proveedor = ttk.Frame(notebook)
notebook.add(Proveedor, text='Proveedor')
def leer_columnasproveedor(ruta_archivo, listbox):
    df = pd.read_excel(ruta_archivo)  # Lee el archivo Excel
    columnas = df.columns.tolist()  # Obtiene la lista de columnas
    listbox.delete(0, tk.END)  # Borra cualquier contenido previo en el Listbox
    for col in columnas:
        listbox.insert(tk.END, col)
def examinar_archivoProveedor():
    archivo = filedialog.askopenfilename(title="Seleccionar archivo", filetypes=[("Archivos de Excel", "*.xlsx *.xls")])
    if archivo:
        leer_columnasproveedor(archivo, listaexcelcliente)
        valorproveedor.delete(0, tk.END)  # Borra cualquier contenido previo en el Entry
        valorproveedor.insert(0, archivo)  # Inserta la ruta del archivo en el Entry
        print(f"Archivo seleccionado: {archivo}")
#columna excel listbox solo observacion
groupboxexcelcliente = tk.LabelFrame(Proveedor, text="Campos Excel ", padx=1, pady=1)
groupboxexcelcliente.place(x=295, y=93, width=100, height=240)
listaexcelcliente = tk.Listbox(groupboxexcelcliente)
listaexcelcliente.place(x=5, y=20, width=80,height=190)
Proveedorframe = tk.LabelFrame(Proveedor, padx=1, pady=1)
Proveedorframe.place(x=20, y=20, width=690, height=50)
# Crear el campo de entrada para mostrar la ruta del archivo
valorproveedor = tk.Entry(Proveedorframe)
valorproveedor.place(x=20, y=10, width=540, height=25)
# Crear el botón de examinar archivo
btn_examinar = tk.Button(Proveedorframe, width=12, height=1, text="Examinar", command=examinar_archivoProveedor)
btn_examinar.place(x=570, y=10)
#funciones necesarias
columnas_Proveedor = {
    "EMPRESA": "Nombre",
    "APELLIDO": "Apellido",
    "DIRECCION1": "Direccion",
    "RNC": "RNC",
    "TELEFONO1": "Celular",
    "TELEFONO2": "Telefono",
    "FAX" : "Fax",
    "MAIL": "Mail",
}

# Lista de posiciones de las columnas
posiciones_proveedor = []

def derechalistProveedor():
    seleccion = listacolumna1prove.curselection()
    if seleccion:
        valor_seleccionado = listacolumna1prove.get(seleccion)
        listacolumna2prove.insert(tk.END, valor_seleccionado)
        listacolumna2prove.select_set(tk.END)
        listacolumna1prove.delete(seleccion)

        for key, value in columnas_Proveedor.items():
            if value == valor_seleccionado:
                posiciones_proveedor.append(key)
                break

def izquierdalistProveedor():
    seleccion = listacolumna2prove.curselection()
    if seleccion:
        valor_seleccionado = listacolumna2prove.get(seleccion)
        listacolumna1prove.insert(tk.END, valor_seleccionado)
        listacolumna1prove.select_set(tk.END)
        listacolumna2prove.delete(seleccion)

        for key, value in columnas_Proveedor.items():
            if value == valor_seleccionado:
                posiciones_proveedor.remove(key)
                break

def limpiarProveedor():
    listacolumna1prove.delete(0, tk.END)
    listacolumna2prove.delete(0, tk.END)
    posiciones_proveedor.clear()
    valorproveedor.delete(0, tk.END)
    combobox2Condi.set('')
    combobox1md.set('')
    
    cargar_archivo_Proveedor()

def cargar_archivo_Proveedor():
    for key in columnas_Proveedor.keys():
        listacolumna1prove.insert(tk.END, columnas_Proveedor[key])


result_md = None
def on_combobox_select(event):
    global result_md
    selected_md = combobox1md.get()
    cursor = db.conexiondb.cursor()
    cursor.execute("SELECT NOMBRE_MONEDA FROM MONEDA WHERE NOMBRE_MONEDA = ?", (selected_md,))
    result = cursor.fetchone()
    if result is not None:
        result_md = result[0]
        print(result_md)
    else:
        print("No se encontró ningún resultado.")
        result_md = None
result_condi = None

def on_combobox_select_condiciones(event):
    global result_condi
    selected_cond = combobox2Condi.get()
    cursor = db.conexiondb.cursor()
    cursor.execute("SELECT  codiciones FROM CONDICIONES_A_PAGAR where codiciones= ?", (selected_cond,))
    result = cursor.fetchone()
    if result is not None:
        result_condi = result[0]
        print(result_condi)
    else:
        print("No se encontró ningún resultado.")
        result_condi = None
#------------------------------------------------------------------------------------------------        
groupbox = tk.LabelFrame(Proveedor,text="Selecciones las Opciones ",padx=1, pady=1)
groupbox.place(x=20, y=93, width=260, height=240)
labelcondiciones = tk.Label(groupbox, text="Condiciones",font=("Arial", 12,"bold"))
labelcondiciones.place(x=5, y=30)
# Crear el segundo ComboBox dentro del GroupBox
combobox2Condi = ttk.Combobox(groupbox, values=db.Condiciones())
combobox2Condi.place(x=110, y=30,width=140)
combobox2Condi.bind("<<ComboboxSelected>>", on_combobox_select_condiciones)
# Crear el Label dentro del GroupBox
labeltrank = tk.Label(groupbox, text="Moneda",font=("Arial", 12,"bold"))
labeltrank.place(x=5, y=60)
# Crear el primer ComboBox dentro del GroupBox
combobox1md = ttk.Combobox(groupbox, values=db.Moneda())
combobox1md.place(x=110, y=60,width=140)
combobox1md.bind("<<ComboboxSelected>>", on_combobox_select)
#------------------------------------------------------------------------------------------------
def actualizarproveedor():
    consulta=db.conexiondb.cursor()
    consulta.execute('select count (*) from proveedor')
    resultadoc=consulta.fetchall()[0][0]
    informacionproveedor.config(text=str(resultadoc))
#------------------------------------------------------------------------------------------------
def insertarProveedor(row, posiciones_proveedor):
    try:
        cursor = db.conexiondb.cursor()
        emp = row[posiciones_proveedor.index("EMPRESA")]
        cursor.execute("SELECT ID FROM PROVEEDOR WHERE EMPRESA = ?", (emp.upper(),))
        result = cursor.fetchone()
        if result:
            messagebox.showinfo("Error Proveedor Existente", f"Proveedor con : \n {emp} ya existe. No se insertará.")
            return False
        
        # Obtener el máximo ID y aumentar en uno
        cursor.execute("SELECT MAX(ID) FROM PROVEEDOR")
        COD_PRO = cursor.fetchone()[0]
        if COD_PRO is None:
            COD_PRO = 100  # Si no hay registros, comenzar con 100
        else:
            COD_PRO += 1  # Aumentar en uno el valor máximo
        
        fecha = timestamp
        columns = ','.join(posiciones_proveedor)
        placeholders = ','.join(['?' for _ in range(len(posiciones_proveedor))])
        # Formatear correctamente la consulta SQL
        sql = f"INSERT INTO PROVEEDOR (COD_PROVEEDOR, FECHA, CONDICIONES_PAGO, MONEDA,NOMBRE, {columns}) VALUES (?, ?, ?, ?,?, {placeholders})"
        # Construir los valores a insertar
        values = [str(COD_PRO), fecha, result_condi, result_md,emp.upper()] + row
        print(sql, '\n', values)
        cursor.execute(sql, values)
        db.conexiondb.commit()
        return True
        
    except Exception as e:
        registrar_error(str(e))
        print(f"Error al insertar proveedor: {e}")
        return False
# Función para leer el archivo y procesar proveedores
insertado = 0
#--------------------------------------------------------------------------------------------------
def insertexcelProvedor():
    global insertado 
    ruta = valorproveedor.get()
    if not ruta:
        print("No se ha seleccionado un archivo")
        return
    try:
        df = pd.read_excel(ruta, header=None, skiprows=1)
        for idx, row in df.iterrows():
            try:
                row = row.fillna(0).replace({pd.NA: 0, pd.NaT: 0, None: 0, 'nan': 0}).tolist()
                
                # Verificar si la fila tiene suficientes elementos según posiciones_proveedor
                if insertarProveedor(row, posiciones_proveedor):
                    print("Proveedor insertado con éxito")
                    iniciarpro()
                    insertado += 1
                else:
                    insertado -= 1
            except Exception as e:
                registrar_error(str(e))
                messagebox.showinfo("Error", f"Error al insertar proveedor en la fila {idx}: {e}")
        
        messagebox.showinfo("Registro Proveedor", f"La cantidad de proveedores registrados es: \n {str(insertado)}")
        detenerpro()
        actualizarproveedor()
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
    finally:
        db.Proveedormayus()
        actualizarproveedor()

#listado columna de proveedor-----------------------------------------------------------------------
groupboxlista = tk.LabelFrame(Proveedor, text="Selecciones Columna de Excel ", padx=1, pady=1)
groupboxlista.place(x=400, y=93, width=310, height=240)
listacolumna1prove = tk.Listbox(groupboxlista)
listacolumna1prove.place(x=5, y=20, width=100,height=190)
cargar_archivo_Proveedor()  
listacolumna2prove = tk.Listbox(groupboxlista)
listacolumna2prove.place(x=200, y=20, width=100,height=190)
btnderechaproveedor = tk.Button(groupboxlista, image=btnlef,command=derechalistProveedor)
btnderechaproveedor.place(x=110, y=80, width=85,height=35)
btnizquierdaproveedor = tk.Button(groupboxlista, image=btnizq,command=izquierdalistProveedor)
btnizquierdaproveedor.place(x=110, y=120,  width=85,height=35)

#botonfuciones proveedor------------------------------------------------------------------------------    
botonfuncionproveedor = tk.LabelFrame(Proveedor, padx=1, pady=1)
botonfuncionproveedor.place(x=400, y=348, width=310, height=45)
btnlimpiarproveedor = tk.Button(botonfuncionproveedor, image=btnlimpiar, command=limpiarProveedor)
btnlimpiarproveedor.place(x=10, y=1, width=90, height=35)
btncerrarproveedor = tk.Button(botonfuncionproveedor, image=btncerrar, command=root.destroy)
btncerrarproveedor.place(x=110, y=1, width=90, height=35)
btncargarproveedor = tk.Button(botonfuncionproveedor, image=btncargar1, command=insertexcelProvedor)
btncargarproveedor.place(x=210, y=1, width=90, height=35)

#mostrar informacion proveedor--------------------------------------------------------------------------
infoproveedor = tk.LabelFrame(Proveedor, padx=0, pady=0,)
infoproveedor.place(x=20, y=348, width=260, height=45)
canproveedor= tk.Label(infoproveedor, text="Cantidad Proveedor", font=("Arial", 9, "bold"))
canproveedor.place(x=2, y=0)
consulta=db.conexiondb.cursor()
consulta.execute('select count (*) from proveedor')
resultado1=consulta.fetchall()[0][0]
informacionproveedor=tk.Label(infoproveedor,text=str(resultado1),font=("Arial", 9))
informacionproveedor.place(x=140,y=0,width=40)
estatuproveedor = tk.Label(infoproveedor, text="Estatus", font=("Arial", 9, "bold"))
estatuproveedor.place(x=2, y=20)  
reproveedor=tk.Label(infoproveedor,text='',font=("Arial", 9))
reproveedor.place(x=120,y=20)
progressbarproveedor = ttk.Progressbar(infoproveedor, orient=tk.HORIZONTAL, length=154, mode='indeterminate')
#---------------------------------------------------------------------------------------------------------
def iniciarpro():
    progressbarproveedor.start()
    progressbarproveedor.place(x=110, y=20)  
# Función para detener la barra de progreso
def detenerpro():
    progressbarproveedor.stop()
    progressbarproveedor.place_forget()  # Ocultar la barra de progreso
# Función para ocultar la barra de progreso
def ocultarpro():
    progressbarproveedor.place_forget()
#----VENTANA PROVEEDOR FINAL CODIGO----------------------------------------------------------------------

def update_time():
    current_time = datetime.now().strftime("%H:%M:%S")
    status_label3.config(text=f"JamenSoft {current_time}")
    root.after(1000, update_time)  # Actualiza cada 1000 ms (1 segundo)
statusbar = tk.Label(root, text="Nemesys Version 1.4", bd=1, relief=tk.SUNKEN, anchor=tk.W)
statusbar.place(x=0, y=500, width=750)
status_label3 = tk.Label(statusbar, text="JamenSoft", bd=1, relief="sunken", anchor="w")
status_label3.place(relx=1.0, rely=0.5, anchor="e")

update_time() 
root.iconbitmap(r"Icono/Import32.ico")
root.resizable(width=0, height=0)
root.mainloop()


