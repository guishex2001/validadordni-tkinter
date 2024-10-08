import openpyxl
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox


# Función para cargar los datos del archivo Excel
def cargar_base_datos(ruta_archivo):
    wb = openpyxl.load_workbook(ruta_archivo)
    hoja = wb.active
    base_datos = {}

    # Leer filas y guardar en un diccionario {DNI: (Tipo, Name)}
    for fila in hoja.iter_rows(min_row=2, values_only=True):
        dni, tipo, name = fila
        # Convertir el DNI a entero y luego a cadena, eliminando espacios
        dni_str = str(int(dni)).strip()
        base_datos[dni_str] = (tipo, name)
    return base_datos


# Función para extraer el DNI entre la 4ta y 5ta comilla
def extraer_dni(texto):
    # Encuentra todas las posiciones de las comillas
    indices_comillas = [i for i, char in enumerate(texto) if char == '"']

    # Asegurarse de que hay al menos 5 comillas para obtener la cuarta y quinta
    if len(indices_comillas) >= 5:
        # Extraer el texto entre la 4ta y 5ta comilla
        start = indices_comillas[3] + 1
        end = indices_comillas[4]
        dni_potencial = texto[start:end].strip()

        # Verificar si el contenido es un número válido (7 a 8 dígitos)
        if re.fullmatch(r'\d{7,8}', dni_potencial):
            return dni_potencial

    return None


# Función para validar el DNI
def validar_dni(dni, base_datos):
    if dni in base_datos:
        tipo, name = base_datos[dni]
        return f"El DNI {dni} corresponde a {name} ({tipo})"
    else:
        return f"El DNI {dni} no se encuentra en la base de datos"


# Función para cargar el archivo de Excel y actualizar la base de datos
def cargar_archivo():
    ruta_archivo = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos de Excel", "*.xlsx")]
    )
    if ruta_archivo:
        try:
            global base_datos
            base_datos = cargar_base_datos(ruta_archivo)
            messagebox.showinfo("Éxito", "Archivo cargado correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar el archivo: {e}")


# Función para procesar el código escaneado
def procesar_codigo(event=None):  # Acepta un parámetro event para manejar el bind de "Enter"
    texto_escaneado = entry_codigo.get()
    if not texto_escaneado:
        messagebox.showwarning("Advertencia", "Por favor, ingrese un código para validar.")
        return

    dni = extraer_dni(texto_escaneado)
    if dni:
        resultado = validar_dni(dni, base_datos)
    else:
        resultado = "Error al extraer el DNI"

    # Mostrar resultado en el área de texto
    text_resultado.delete(1.0, tk.END)
    text_resultado.insert(tk.END, resultado)

    # Limpiar el campo de entrada para el siguiente escaneo
    entry_codigo.delete(0, tk.END)


# Configuración de la interfaz gráfica
root = tk.Tk()
root.title("Validador de DNI")
root.geometry("800x600")

# Botón para cargar el archivo Excel
btn_cargar = tk.Button(root, text="Cargar archivo Excel", command=cargar_archivo)
btn_cargar.pack(pady=10)

# Campo de entrada para el código
label_codigo = tk.Label(root, text="Escanea o ingresa el código:")
label_codigo.pack()
entry_codigo = tk.Entry(root, width=50)
entry_codigo.pack(pady=5)

# Vincular la tecla "Enter" a la función procesar_codigo
entry_codigo.bind('<Return>', procesar_codigo)

# Botón para procesar el código
btn_procesar = tk.Button(root, text="Validar DNI", command=procesar_codigo)
btn_procesar.pack(pady=10)

# Área de texto para mostrar el resultado
text_resultado = tk.Text(root, height=5, width=50)
text_resultado.pack(pady=10)

# Iniciar la aplicación
base_datos = {}
root.mainloop()
