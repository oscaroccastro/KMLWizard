import os
import utm
import simplekml
import webbrowser
import configparser
import csv

from tkinter import *
from tkinter import messagebox, filedialog
from ttkbootstrap import Window, Bootstyle, ttk
from tkhtmlview import HTMLLabel

# Ruta del archivo de configuración
config_file = "config.ini"

# Variables globales
hemisferio_config = ""
zona_config = ""
descripcion_kml = ""
archivo_temporal = "temp.html"
agregar_html_kml = False
tamanio_linea_config = 1.0
color_linea_config = ""
color_poligono_config = ""

# Funciones de configuración
def guardar_configuracion():
    global zona_variable, hemisferio_variable  # Añadir esta línea
    zona = zona_variable.get()
    hemisferio = hemisferio_variable.get()
    config = configparser.ConfigParser()
    config["Configuracion"] = {"zona": zona, "hemisferio": hemisferio}
    with open(config_file, "w") as f:
        config.write(f)
    messagebox.showinfo("Información", "La configuración se ha guardado correctamente.")

def cargar_configuracion():
    global zona_config, hemisferio_config  # Añadir esta línea
    config = configparser.ConfigParser()
    if os.path.exists(config_file):
        config.read(config_file)
        zona_config = config.get("Configuracion", "zona", fallback="")
        hemisferio_config = config.get("Configuracion", "hemisferio", fallback="")
    else:
        guardar_configuracion()

# Función para convertir coordenadas
def convertir_coordenadas():
    if not hemisferio_variable.get() or not zona_variable.get() or not coordenadas_x_text.get("1.0", END).strip() or not coordenadas_y_text.get("1.0", END).strip():
        messagebox.showerror("Error", "Por favor, verifique todos los campos.")
        return
    zona = int(zona_variable.get())
    hemisferio = hemisferio_variable.get()
    coordenadas_decimales = []
    for x, y in zip(coordenadas_x_text.get("1.0", END).strip().split('\n'), coordenadas_y_text.get("1.0", END).strip().split('\n')):
        lat, lon = utm.to_latlon(float(x), float(y), zona, hemisferio)
        coordenadas_decimales.append((lat, lon))
    guardar_coordenadas(coordenadas_decimales)

# Función para guardar coordenadas en un archivo de texto
def guardar_coordenadas(coordenadas):
    archivo = filedialog.asksaveasfile(defaultextension=".txt", filetypes=[("Archivos de texto", "*.txt")])
    if archivo:
        try:
            for lat, lon in coordenadas:
                archivo.write(f"Latitud: {lat}\tLongitud: {lon}\n")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")
        finally:
            archivo.close()

# Función para cargar coordenadas desde un archivo CSV
def cargar_coordenadas():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos CSV", "*.csv")])
    if archivo:
        coordenadas_x = []
        coordenadas_y = []
        try:
            with open(archivo, "r") as f:
                reader = csv.reader(f)
                next(reader)  # Saltar la primera línea (encabezados)
                for row in reader:
                    x = float(row[0].replace(",", "."))  # Reemplazar coma por punto si es necesario
                    y = float(row[1].replace(",", "."))  # Reemplazar coma por punto si es necesario
                    coordenadas_x.append(x)
                    coordenadas_y.append(y)
            
            # Limpiar los campos actuales
            coordenadas_x_text.delete("1.0", END)
            coordenadas_y_text.delete("1.0", END)

            # Agregar las coordenadas al campo de texto
            for x, y in zip(coordenadas_x, coordenadas_y):
                coordenadas_x_text.insert(END, f"{x}\n")
                coordenadas_y_text.insert(END, f"{y}\n")

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo: {e}")

# Función para generar un archivo KML
def generar_kml():
    if not hemisferio_variable.get() or not zona_variable.get() or not coordenadas_x_text.get("1.0", END).strip() or not coordenadas_y_text.get("1.0", END).strip():
        messagebox.showerror("Error", "Por favor, verifique todos los campos.")
        return
    coordenadas_decimales = [(utm.to_latlon(float(x), float(y), int(zona_variable.get()), hemisferio_variable.get())) for x, y in zip(coordenadas_x_text.get("1.0", END).strip().split('\n'), coordenadas_y_text.get("1.0", END).strip().split('\n'))]
    kml = simplekml.Kml()
    tipo_geometria = tipo_geometria_variable.get()
    nombre = nombre_entry.get()
    archivo_kml = f"{nombre}.kml"
    
    if descripcion_kml:
        descripcion = simplekml.CDATA(f"<b>{nombre}</b><br><br>{descripcion_kml}") if agregar_html_kml else descripcion_kml
        kml.document.newdescription(descripcion)
    
    if tipo_geometria == "Punto":
        for lat, lon in coordenadas_decimales:
            kml.newpoint(name=nombre, coords=[(lon, lat)])
    elif tipo_geometria == "Polilínea":
        linea = kml.newlinestring(name=nombre, coords=[(lon, lat) for lat, lon in coordenadas_decimales])
    elif tipo_geometria == "Polígono":
        poligono = kml.newpolygon(name=nombre, outerboundaryis=[(lon, lat) for lat, lon in coordenadas_decimales])
    
    try:
        kml.save(archivo_kml)
        messagebox.showinfo("Información", f"El archivo KML se ha generado correctamente en '{archivo_kml}'.")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo generar el archivo KML: {e}")

# Función para generar tabla HTML
def generar_tabla_html():
    if not hemisferio_variable.get() or not zona_variable.get() or not coordenadas_x_text.get("1.0", END).strip() or not coordenadas_y_text.get("1.0", END).strip():
        messagebox.showerror("Error", "Por favor, complete todos los campos.")
        return
    coordenadas_decimales = [(utm.to_latlon(float(x), float(y), int(zona_variable.get()), hemisferio_variable.get())) for x, y in zip(coordenadas_x_text.get("1.0", END).strip().split('\n'), coordenadas_y_text.get("1.0", END).strip().split('\n'))]
    contenido_html = f'''
    <html>
    <head>
        <style>
            table, th, td {{ border: 1px solid black; border-collapse: collapse; }}
            th, td {{ padding: 5px; }}
        </style>
    </head>
    <body>
        <h2>{nombre_entry.get()}</h2>
        <table>
            <tr><th colspan="2">Coordenadas UTM</th></tr>
            <tr><th>X</th><th>Y</th></tr>
    '''
    for lat, lon in coordenadas_decimales:
        x, y = utm.from_latlon(lat, lon)[:2]
        contenido_html += f'<tr><td>{x:.4f}</td><td>{y:.4f}</td></tr>'
    contenido_html += '</table></body></html>'
    mostrar_html(archivo_temporal, contenido_html)

# Función para mostrar el contenido HTML
def mostrar_html(archivo, contenido):
    with open(archivo, 'w') as f:
        f.write(contenido)
    webbrowser.open(archivo)

# Funciones para manejar el código HTML
def mostrar_codigo_html():
    try:
        with open(archivo_temporal, "r") as f:
            contenido_html = f.read()
        codigo_html_ventana = Toplevel(root)
        codigo_html_ventana.title("Código HTML")
        codigo_html_text = Text(codigo_html_ventana, height=20, width=80)
        codigo_html_text.pack()
        codigo_html_text.insert(END, contenido_html)
        codigo_html_text.config(state=DISABLED)
        ttk.Button(codigo_html_ventana, text="Copiar Código", command=lambda: copiar_contenido(contenido_html)).pack()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo mostrar el código HTML: {e}")

def copiar_contenido(contenido):
    root.clipboard_clear()
    root.clipboard_append(contenido)
    messagebox.showinfo("Información", "El contenido se ha copiado al portapapeles.")

# Función para actualizar la etiqueta del nombre según el tipo de geometría
def actualizar_nombre_etiqueta():
    tipo_geometria = tipo_geometria_variable.get()
    if tipo_geometria == "Punto":
        nombre_label.config(text="Nombre de los puntos:")
    elif tipo_geometria == "Polilínea":
        nombre_label.config(text="Nombre de la polilínea:")
    elif tipo_geometria == "Polígono":
        nombre_label.config(text="Nombre del polígono:")

# Función para mostrar y guardar propiedades KML
def mostrar_propiedades_kml():
    propiedades_ventana = Toplevel(root)
    propiedades_ventana.title("Propiedades KML")
    agregar_html = BooleanVar()
    tamanio_linea = DoubleVar()
    color_linea = StringVar()
    color_poligono = StringVar()

    def guardar_propiedades():
        global agregar_html_kml, tamanio_linea_config, color_linea_config, color_poligono_config
        agregar_html_kml = agregar_html.get()
        tamanio_linea_config = tamanio_linea.get()
        color_linea_config = color_linea.get()
        color_poligono_config = color_poligono.get()
        propiedades_ventana.destroy()

    ttk.Label(propiedades_ventana, text="Agregar código HTML:").pack()
    ttk.Checkbutton(propiedades_ventana, variable=agregar_html, bootstyle="success-round-toggle").pack()
    ttk.Label(propiedades_ventana, text="Tamaño de la línea:").pack()
    ttk.Scale(propiedades_ventana, variable=tamanio_linea, from_=0.1, to=10, resolution=0.1, orient=HORIZONTAL).pack()
    ttk.Label(propiedades_ventana, text="Color de la línea:").pack()
    ttk.Entry(propiedades_ventana, textvariable=color_linea).pack()
    ttk.Label(propiedades_ventana, text="Color del polígono:").pack()
    ttk.Entry(propiedades_ventana, textvariable=color_poligono).pack()
    ttk.Button(propiedades_ventana, text="Guardar", command=guardar_propiedades).pack()

# Función para mostrar información de "Acerca de"
def mostrar_acerca_de():
    ventana_acerca_de = Toplevel(root)
    ventana_acerca_de.title("Acerca de")
    descripcion = ("Conversor de Coordenadas - Versión: 1.6\n\n"
                   "Créditos:\n\nO. Contreras\nChatGPT\nEcosoluciones Ambientales\n\n"
                   "Visita el Repositorio de este Programa")
    ttk.Label(ventana_acerca_de, text=descripcion).pack(padx=10, pady=10)
    def abrir_enlace():
        webbrowser.open("https://github.com/Ecosoluciones-Ambientales/KMLConverter")
    ttk.Button(ventana_acerca_de, text="Abrir enlace", command=abrir_enlace).pack(side=LEFT, padx=5)
    ventana_acerca_de.bind("<Escape>", lambda event: ventana_acerca_de.destroy())
    ventana_acerca_de.focus_set()

# Función para mostrar la ayuda
def mostrar_help():
    ventana_ayuda = Toplevel(root)
    ventana_ayuda.title("Ayuda")
    with open("manual.html", "r") as f:
        contenido_html = f.read()
    HTMLLabel(ventana_ayuda, html=contenido_html).pack(expand=True, fill=BOTH)

# Función para salir de la aplicación
def salir():
    root.quit()

# Crear la ventana principal
root = Window(title="KMLWizard V.1.6", themename="darkly")
root.resizable(False, False)

# Inicializar las variables globales de Tkinter
zona_variable = StringVar(root)
hemisferio_variable = StringVar(root)

# Cargar la configuración
cargar_configuracion()

# Crear el menú superior
menu_superior = Menu(root)
root.config(menu=menu_superior)

menu_archivo = Menu(menu_superior, tearoff=0)
menu_superior.add_cascade(label="Archivo", menu=menu_archivo)
menu_archivo.add_command(label="Cargar coordenadas", command=cargar_coordenadas)
menu_archivo.add_separator()
menu_archivo.add_command(label="Guardar Coordenadas", command=convertir_coordenadas)
menu_archivo.add_command(label="Guardar configuración", command=guardar_configuracion)
menu_archivo.add_separator()
menu_archivo.add_command(label="Preferencias", command=guardar_configuracion)
menu_archivo.add_separator()
menu_archivo.add_command(label="Salir", command=salir)

menu_editor = Menu(menu_superior, tearoff=0)
menu_superior.add_cascade(label="Editor", menu=menu_editor)
menu_editor.add_command(label="Previsualizar Tabla", command=generar_tabla_html)
menu_editor.add_command(label="Mostrar Código HTML", command=mostrar_codigo_html)
menu_editor.add_separator()
menu_editor.add_command(label="Propiedades KML", command=mostrar_propiedades_kml)

menu_informacion = Menu(menu_superior, tearoff=0)
menu_superior.add_cascade(label="Información", menu=menu_informacion)
menu_informacion.add_command(label="Acerca de", command=mostrar_acerca_de)
menu_informacion.add_command(label="Help", command=mostrar_help)

# Frame para los menús desplegables
frame_menus = ttk.Frame(root)
frame_menus.pack(pady=10)

zona_label = ttk.Label(frame_menus, text="Zona:")
zona_label.pack(side=LEFT, padx=5)
zona_variable = StringVar(root)
zona_variable.set(zona_config)
zona_menu = ttk.OptionMenu(frame_menus, zona_variable, zona_config, *range(1, 61))
zona_menu.pack(side=LEFT, padx=5)

hemisferio_label = ttk.Label(frame_menus, text="Hemisferio:")
hemisferio_label.pack(side=LEFT, padx=5)
hemisferio_variable = StringVar(root)
hemisferio_variable.set(hemisferio_config)
hemisferio_menu = ttk.OptionMenu(frame_menus, hemisferio_variable, hemisferio_config, "Norte", "Sur")
hemisferio_menu.pack(side=LEFT, padx=5)

coordenadas_x_label = ttk.Label(root, text="Coordenadas X:")
coordenadas_x_label.pack(pady=5)
coordenadas_x_text = Text(root, height=5, width=30)
coordenadas_x_text.pack(pady=5)

coordenadas_y_label = ttk.Label(root, text="Coordenadas Y:")
coordenadas_y_label.pack(pady=5)
coordenadas_y_text = Text(root, height=5, width=30)
coordenadas_y_text.pack(pady=5)

tipo_geometria_label = ttk.Label(root, text="Tipo de geometría:")
tipo_geometria_label.pack(pady=5)
tipo_geometria_variable = StringVar()
tipo_geometria_frame = ttk.Frame(root)
tipo_geometria_frame.pack(pady=5)
ttk.Radiobutton(tipo_geometria_frame, text="Punto", variable=tipo_geometria_variable, value="Punto", command=actualizar_nombre_etiqueta).pack(side=LEFT, padx=5)
ttk.Radiobutton(tipo_geometria_frame, text="Polilínea", variable=tipo_geometria_variable, value="Polilínea", command=actualizar_nombre_etiqueta).pack(side=LEFT, padx=5)
ttk.Radiobutton(tipo_geometria_frame, text="Polígono", variable=tipo_geometria_variable, value="Polígono", command=actualizar_nombre_etiqueta).pack(side=LEFT, padx=5)

nombre_label = ttk.Label(root, text="Nombre del polígono:")
nombre_label.pack(pady=5)
nombre_entry = ttk.Entry(root)
nombre_entry.pack(pady=5)

kml_boton = ttk.Button(root, text="Generar KML", command=generar_kml, bootstyle="primary")
kml_boton.pack(pady=5)
tabla_html_boton = ttk.Button(root, text="Previsualizar HTML", command=generar_tabla_html, bootstyle="secondary")
tabla_html_boton.pack(pady=5)

root.geometry("400x520")
root.mainloop()
