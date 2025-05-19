import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook

#creammos un archivo de excel
libro = Workbook()
hoja = libro.active
hoja.append(["ID", "N° Serie", "Cantidad", "Descripción", "Lugar"])


# Guardar el archivo de Excel
# libro.save("inventario.xlsx")

def recibir_datos():
    try:
        id = int(entry_id.get())
        serie = int(entry_serie.get())
        cantidad = int(entry_cantidad.get())
        descripcion = entry_descripcion.get()
        lugar = entry_lugar.get()

        entradas = [id, serie, cantidad, descripcion, lugar]

        valido, mensaje = validar_entrada(entradas)
        
        if not valido:
            messagebox.showerror("Error", mensaje)
            return
        messagebox.showinfo("Validación", "Datos válidos.")
        #validacion a guardar_datos 
        return entradas, True
    
    except ValueError:
        messagebox.showerror("Error", "ID, Serie y Cantidad deben ser números enteros.")

     # Limpiar las entradas
    entry_id.delete(0, tk.END)
    entry_serie.delete(0, tk.END)
    entry_cantidad.delete(0, tk.END)
    entry_obs.delete(0, tk.END)
    entry_descripcion.delete(0, tk.END)
    entry_lugar.delete(0, tk.END)

def validar_entrada(entrada):

    id, serie, cantidad, descripcion, lugar = entrada

    # Validar que no haya campos vacíos
    if not id or not serie or not cantidad or not descripcion or not lugar:
        return False, "Todos los campos son obligatorios."
    
    # Validar que ID, Serie y Cantidad sean enteros
    if not isinstance(id, int) or not isinstance(serie, int) or not isinstance(cantidad, int):
        return False, "ID, Serie y Cantidad deben ser números enteros."

    # Validar que Descripción y Lugar sean strings
    if not isinstance(descripcion, str) or not isinstance(lugar, str):
        return False, "Descripción y Lugar deben ser texto."

    return True, "Datos válidos."

# Función para guardar los datos en el archivo de Excel
def guardar_datos():
    datos , estado = recibir_datos()
    if datos:
        id, serie, cantidad, descripcion, lugar = datos
        # Agregar los datos a la hoja de Excel
        hoja.append([id, serie, cantidad, descripcion, lugar])
        
        # Guardar el archivo de Excel
        libro.save("inventario.xlsx")

        # Muestrar un mensaje de éxito  
        messagebox.showinfo("Éxito", "Datos guardados correctamente.")
    else:
        messagebox.showerror("Error", "No se pudieron guardar los datos.")
    
def retirar_datos():
    pass
def buscar_datos():
    pass
def eliminar_datos():
    pass

# Crear la ventana principal
root = tk.Tk()
root.title("Inventario de sistemas ARSA")
root.configure(bg="#4B6587")

lebal_style = {
    "bg": "#4B6587",
    "fg": "white",
    "font": ("Arial", 12)
}
entry_style = {
    "bg": "#d3d3d3",
    "fg": "black",
    "font": ("Arial", 12)
}

# Crear etiquetas y entradas para cada campo

label_id = tk.Label(root, text="ID", **lebal_style)
label_id.grid(row=0, column=0, padx=10, pady=5)
entry_id = tk.Entry(root, **entry_style)
entry_id.grid(row=0, column=1, padx=10, pady=5)

label_serie = tk.Label(root, text="N° Serie", **lebal_style)
label_serie.grid(row=1, column=0, padx=10, pady=5)
entry_serie = tk.Entry(root, **entry_style)
entry_serie.grid(row=1, column=1, padx=10, pady=5)

label_cantidad = tk.Label(root, text="Cantidad", **lebal_style)
label_cantidad.grid(row=2, column=0, padx=10, pady=5)
entry_cantidad = tk.Entry(root, **entry_style)
entry_cantidad.grid(row=2, column=1, padx=10, pady=5)

label_descripcion = tk.Label(root, text="Descripcion", **lebal_style)
label_descripcion.grid(row=3, column=0, padx=10, pady=5)
entry_descripcion = tk.Entry(root, **entry_style)
entry_descripcion.grid(row=3, column=1, padx=10, pady=5)

label_obs = tk.Label(root, text="Observacion", **lebal_style)
label_obs.grid(row=5, column=0, padx=10, pady=5)
entry_obs = tk.Entry(root, **entry_style)
entry_obs.grid(row=5, column=1, padx=10, pady=5)

label_lugar = tk.Label(root, text="Lugar", **lebal_style)
label_lugar.grid(row=6, column=0, padx=10, pady=5)
entry_lugar = tk.Entry(root, **entry_style)
entry_lugar.grid(row=6, column=1, padx=10, pady=5)


# Crear el botón para guardar los datos sin parentesis
boton_guardar = tk.Button(root, text="Guardar", command=guardar_datos, bg="#6d8299", fg="white", font=("Arial", 12))
boton_guardar.grid(row=7, column=0, padx=5, pady=5, sticky="ew")

# Botón para retirar datos
boton_retirar = tk.Button(root, text="Retirar", command=retirar_datos, bg="#6d8299", fg="white", font=("Arial", 12))
boton_retirar.grid(row=7, column=1, padx=5, pady=5, sticky="ew")

# Botón para buscar datos
boton_buscar = tk.Button(root, text="Buscar", command=buscar_datos, bg="#6d8299", fg="white", font=("Arial", 12))
boton_buscar.grid(row=7, column=2, padx=5, pady=5, sticky="ew")

# Botón para eliminar datos
boton_eliminar = tk.Button(root, text="Eliminar", command=eliminar_datos, bg="#6d8299", fg="white", font=("Arial", 12))
boton_eliminar.grid(row=7, column=3, padx=5, pady=5, sticky="ew")

# Configurar las columnas para que ocupen espacio proporcionalmente
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_columnconfigure(2, weight=1)
root.grid_columnconfigure(3, weight=1)


# sirve para que no se cierre la ventana
root.mainloop()