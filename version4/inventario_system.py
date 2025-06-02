import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
from pathlib import Path
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

class InventarioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventario de sistemas")
        self.root.geometry("1280x720")
        self.root.configure(bg="#f0f0f0")
        
        # Archivo Excel
        self.archivo_excel = "inventario.xlsx"
        self.inicializar_excel()
        
        # Variables para los campos
        self.id_var = tk.IntVar()
        self.descripcion_var = tk.StringVar()
        self.serie_var = tk.StringVar()
        self.observaciones_var = tk.StringVar()
        self.lugar_var = tk.StringVar()
        self.cantidad_var = tk.StringVar()

        # variables para campos de busqueda
        self.id_buscar_var = tk.StringVar()
        self.descripcion_buscar_var = tk.StringVar()
        self.eleccion = tk.StringVar(value="")

        # Crear interfaz
        self.crear_interfaz()
        
        # Cargar datos iniciales en la tabla
        self.actualizar_tabla()
    
    def inicializar_excel(self):
        # Verificar si el archivo existe
        if not os.path.exists(self.archivo_excel):
            # Crear un nuevo archivo Excel con las columnas necesarias
            df = pd.DataFrame(columns=[
                "ID", "Descripcion", "Serie", 
                "Observaciones", "Lugar", "Cantidad"
            ])
            df.to_excel(self.archivo_excel, index=False)
            print(f"Archivo {self.archivo_excel} creado correctamente.")
        else:
            try:
                # Verificar que el archivo es accesible y tiene el formato correcto
                pd.read_excel(self.archivo_excel)
                print(f"Archivo {self.archivo_excel} cargado correctamente.")
            except (InvalidFileException, Exception) as e:
                messagebox.showerror("Error", f"El archivo de inventario está dañado o no es accesible: {str(e)}")
                # Crear un backup y un nuevo archivo
                if os.path.exists(self.archivo_excel):
                    os.rename(self.archivo_excel, f"{self.archivo_excel}.bak")
                df = pd.DataFrame(columns=[
                    "ID", "Descripcion", "Serie", 
                    "Observaciones", "Lugar", "Cantidad"
                ])
                df.to_excel(self.archivo_excel, index=False)
                
    def crear_interfaz(self):
        # Frame principal
        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Frame horizontal para inputs y sugerencias
        top_frame = tk.Frame(main_frame, bg="#f0f0f0")
        top_frame.pack(fill=tk.X, padx=10, pady=10)

        # Frame para los campos de entrada (izquierda)
        input_frame = tk.LabelFrame(top_frame, text="Datos del Producto", bg="#f0f0f0", font=("Arial", 12))
        input_frame.grid(row=0, column=0, sticky="nw")

        # Descripción + evento de sugerencias
        tk.Label(input_frame, text="Descripcion", bg="#f0f0f0", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        entry = tk.Entry(input_frame, textvariable=self.descripcion_var, font=("Arial", 12), width=30)
        entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        entry.bind("<KeyRelease>", self.tabla_sugerencias)

        self.crear_campo(input_frame, "N°. Serie:", self.serie_var, 2)
        self.crear_campo(input_frame, "Observaciones:", self.observaciones_var, 3)
        self.crear_campo(input_frame, "Lugar:", self.lugar_var, 4)
        self.crear_campo(input_frame, "Cantidad:", self.cantidad_var, 5)

        # Sugerencias (a la derecha)
        self.sugerencia_frame = tk.LabelFrame(top_frame, text="Sugerencia de descripción", bg="#f0f0f0", font=("Arial", 12))
        self.sugerencia_frame.grid(row=0, column=1, padx=20, sticky="ne")

        self.listbox = tk.Listbox(self.sugerencia_frame, font=("Arial", 11), height=10, width=30)
        self.listbox.pack(padx=10, pady=10, fill=tk.BOTH)
        self.listbox.bind("<ButtonRelease-1>", self.tabla_sugerencias)

        # Campo de búsqueda
        busqueda_frame = tk.LabelFrame(main_frame, text="Campo búsqueda", bg="#f0f0f0", font=("Arial", 12))
        busqueda_frame.pack(fill=tk.X, padx=10, pady=10)

        self.crear_campo(busqueda_frame, "Buscar por ID:", self.id_buscar_var, 0)

        # Buscar por descripción + evento de sugerencia
        tk.Label(busqueda_frame, text="Buscar por descripcion:", bg="#f0f0f0", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        buscar_entry = tk.Entry(busqueda_frame, textvariable=self.descripcion_buscar_var, font=("Arial", 12), width=30)
        buscar_entry.grid(row=1, column=1, padx=10, pady=5, sticky="w")
        buscar_entry.bind("<KeyRelease>", self.tabla_sugerencias2)  # <-- NUEVO

        # Botones
        button_frame = tk.Frame(main_frame, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        self.crear_boton(button_frame, "Ingreso Producto", self.agregar_producto, "#2ecc71", 0)
        self.crear_boton(button_frame, "Agregar Stock", self.tipo_agregar, "#9b59b6", 1)
        self.crear_boton(button_frame, "Retirar", self.tipo_retiro, "#e74c3c", 2)
        self.crear_boton(button_frame, "Eliminar", self.tipo_eliminacion, "#e67e22", 4)
        self.crear_boton(button_frame, "Buscar", self.tipo_busqueda, "#3498db", 5)

        # Tabla
        table_frame = tk.LabelFrame(main_frame, text="Inventario Actual", bg="#f0f0f0", font=("Arial", 14))
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(table_frame, columns=("ID", "Descripcion", "Serie", "Observaciones", "Lugar", "Cantidad"), show="headings")

        self.tree.heading("ID", text="ID")
        self.tree.heading("Descripcion", text="Descripcion")
        self.tree.heading("Serie", text="No. Serie")
        self.tree.heading("Observaciones", text="Observaciones")
        self.tree.heading("Lugar", text="Lugar")
        self.tree.heading("Cantidad", text="Cantidad")

        self.tree.column("ID", width=80)
        self.tree.column("Descripcion", width=200)
        self.tree.column("Serie", width=100)
        self.tree.column("Observaciones", width=150)
        self.tree.column("Lugar", width=100)
        self.tree.column("Cantidad", width=80)

        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.bind("<ButtonRelease-1>", self.seleccionar_item)
    
    # Crear campos de entrada
    def crear_campo(self, parent, texto, variable, fila):
        tk.Label(parent, text=texto, bg="#f0f0f0", font=("Arial", 12)).grid(row=fila, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(parent, textvariable=variable, font=("Arial", 12), width=30).grid(row=fila, column=1, padx=10, pady=5, sticky="w")
    
    # Crear campo de búsqueda
    def crear_campo_busqueda(self, parent, texto, variable, fila):
        tk.Label(parent, text=texto, bg="#f0f0f0", font=("Arial", 14)).grid(row=fila, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(parent, textvariable=variable, font=("Arial", 14), width=30).grid(row=fila, column=1, padx=10, pady=5, sticky="w")
    
    # Crear botones con comandos y estilos
    def crear_boton(self, parent, texto, comando, color, columna):
        tk.Button(
            parent, 
            text=texto, 
            command=comando, 
            bg=color, 
            fg="black", 
            font=("Arial", 10, "bold"),
            width=14,
            height=2,
            relief=tk.RAISED
        ).grid(row=0, column=columna, padx=5, pady=5)
    
    #limpia campos de entrada
    def limpiar_campos(self):
        self.id_var.set("")
        self.descripcion_var.set("")
        self.serie_var.set("")
        self.observaciones_var.set("")
        self.lugar_var.set("")
        self.cantidad_var.set("")
    
    # seguir sugerencias de descripcion y muestra en la lista en campos de entrada
    def tabla_sugerencias(self, event):
        # Obtener el texto ingresado
        texto = self.descripcion_var.get().strip().lower()
        
        # Limpiar la lista de sugerencias
        self.listbox.delete(0, tk.END)
        
        if texto:
            try:
                # Cargar datos desde Excel
                df = pd.read_excel(self.archivo_excel)
                
                # Filtrar las descripciones que contengan el texto ingresado
                coincidencias = df[df["Descripcion"].str.contains(texto, case=False, na=False)]
                
                # Agregar coincidencias a la lista de sugerencias
                for descripcion in coincidencias["Descripcion"].unique():
                    self.listbox.insert(tk.END, descripcion)
                    
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar las sugerencias: {str(e)}")
                
     # seguir sugerencias de descripcion y muestra en la lista
    
    #seguir sugerencias de descripcion y muestra en la lista en campos de busqueda
    def tabla_sugerencias2(self, event):
        # Obtener el texto ingresado
        texto = self.descripcion_buscar_var.get().strip().lower()
        
        # Limpiar la lista de sugerencias
        self.listbox.delete(0, tk.END)
        
        if texto:
            try:
                # Cargar datos desde Excel
                df = pd.read_excel(self.archivo_excel)
                
                # Filtrar las descripciones que contengan el texto ingresado
                coincidencias = df[df["Descripcion"].str.contains(texto, case=False, na=False)]
                
                # Agregar coincidencias a la lista de sugerencias
                for descripcion in coincidencias["Descripcion"].unique():
                    self.listbox.insert(tk.END, descripcion)
                    
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo cargar las sugerencias: {str(e)}")

    #actualizar la tabla de inventario
    def actualizar_tabla(self):

        # Limpiar tabla existente
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        try:
            # Cargar datos desde Excel
            df = pd.read_excel(self.archivo_excel)
            
            # Agregar datos a la tabla
            for index, row in df.iterrows():
                self.tree.insert("", tk.END, values=(
                    row.get("ID", ""), 
                    row.get("Descripcion", ""), 
                    row.get("Serie", ""), 
                    row.get("Observaciones", ""), 
                    row.get("Lugar", ""), 
                    row.get("Cantidad", "")
                ))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar los datos: {str(e)}")
    
    # Seleccionar sugerencia de la lista del tabla previ
    def seleccionar_item(self, event):
        # Obtener el ítem seleccionado
        seleccion = self.tree.selection()
        if seleccion:
            item = self.tree.item(seleccion[0])
            valores = item['values']
            
            # Actualizar los campos con los valores seleccionados
            self.id_buscar_var.set(valores[0])
            self.descripcion_var.set(valores[1])
            self.serie_var.set(valores[2])
            self.observaciones_var.set(valores[3])
            self.lugar_var.set(valores[4])
            self.cantidad_var.set(valores[5])
    
    # Validar datos de entrada
    def validar_datos(self):
        
        # Validar que no haya campos vacíos
        if not self.serie_var.get().strip() or not self.cantidad_var or not self.descripcion_var.get().strip() or not self.lugar_var.get().strip():
           messagebox.showwarning("Advertencia", "El campo los campos son obligatorios.")
           return False
        
        # Validar que Descripcion y Lugar sean strings
        if not isinstance(self.descripcion_var.get(), str) or not isinstance(self.lugar_var.get(), str) or not isinstance(self.serie_var.get(), str):
            messagebox.showwarning("Descripcion, Lugar y N° Serie deben ser texto.")
            return False
        
        # Validar que Serie y Cantidad sean enteros
        try:
            cantidad = int(self.cantidad_var.get())

            # Validar que Cantidad y Serie sean positivos
            if cantidad < 0:
                messagebox.showwarning("Advertencia", "La cantidad y N° Serie no pueden ser negativos.")
                return False
            
            return True
        except ValueError:
            messagebox.showwarning("Advertencia", "Serie y Cantidad deben ser números enteros.")
            return False

    # Validar datos de búsqueda
    def validar_datos_busqueda(self):

        # Validar que no haya campos vacíos
        if not self.id_buscar_var.get():
            messagebox.showwarning("Advertencia", "El campo ID es obligatorio.")
            return False
        
        #validar que descripcion sea un string
        if not isinstance(self.descripcion_buscar_var.get(), str):
            messagebox.showwarning("Advertencia", "El campo Descripcion debe ser texto.")
            return False
        
        # Validar que ID sea un entero
        try:
            id_producto = int(self.id_buscar_var.get())
            if id_producto <= 0:
                messagebox.showwarning("Advertencia", "El ID debe ser un número entero positivo.")
                return False
        except ValueError:
            messagebox.showwarning("Advertencia", "El ID debe ser un número entero.")
            return False
        
        return True
    
    # Solicita al usuario como desea eliminar el producto
    def tipo_eliminacion(self):
    
        # Crear ventana de diálogo
        ventana = tk.Toplevel()
        ventana.title("Tipo de Eliminación")
        ventana.geometry("400x200")
        ventana.grab_set()  # Bloquea la ventana principal hasta que se cierre esta

        etiqueta = tk.Label(ventana, text="¿Cómo deseas eliminar el producto?")
        etiqueta.pack(pady=10)

        # Botones con comandos que SETEAN el valor en self.eleccion
        btn_id = tk.Button(ventana, text="Por ID", width=20, command=lambda: [self.eleccion.set("id"), ventana.destroy()])
        btn_id.pack(pady=2)

        btn_descripcion = tk.Button(ventana, text="Por Descripcion", width=20, command=lambda: [self.eleccion.set("descripcion"), ventana.destroy()])
        btn_descripcion.pack(pady=2)

        btn_cancelar = tk.Button(ventana, text="Cancelar", width=20, command=lambda: [self.eleccion.set("cancelar"), ventana.destroy()])
        btn_cancelar.pack(pady=2)

        # Espera hasta que el usuario cierre el cuadro de diálogo
        ventana.wait_window()

        # Evaluar la opción después de cerrar el cuadro
        eleccion = self.eleccion.get()
        if eleccion == "id":
            print("Eliminando por ID")
            self.eliminar_id()
            
        elif eleccion == "descripcion":
            print("Eliminando por Descripcion")
            self.eliminar_descripcion()

        elif eleccion == "cancelar":
            ventana.destroy()
            print("Operación cancelada")
    
    # solicita al usuario como desea retirar el producto
    def tipo_retiro(self):
        # Crear ventana de diálogo
        ventana = tk.Toplevel()
        ventana.title("Tipo de Eliminación")
        ventana.geometry("400x200")
        ventana.grab_set()  # Bloquea la ventana principal hasta que se cierre esta

        etiqueta = tk.Label(ventana, text="¿Cómo deseas Retirar el producto?")
        etiqueta.pack(pady=10)

        # Botones con comandos que SETEAN el valor en self.eleccion
        btn_id = tk.Button(ventana, text="Por ID", width=20, command=lambda: [self.eleccion.set("id"), ventana.destroy()])
        btn_id.pack(pady=2)

        btn_descripcion = tk.Button(ventana, text="Por Descripcion", width=20, command=lambda: [self.eleccion.set("descripcion"), ventana.destroy()])
        btn_descripcion.pack(pady=2)

        btn_cancelar = tk.Button(ventana, text="Cancelar", width=20, command=lambda: [self.eleccion.set("cancelar"), ventana.destroy()])
        btn_cancelar.pack(pady=2)

        # Espera hasta que el usuario cierre el cuadro de diálogo
        ventana.wait_window()

        # Evaluar la opción después de cerrar el cuadro
        eleccion = self.eleccion.get()
        if eleccion == "id":
            print("Retiro por ID")
            self.retiro_id()
            
        elif eleccion == "descripcion":
            print("Retiro por Descripcion")
            self.retiro_descripcion()

        elif eleccion == "cancelar":
            ventana.destroy()
            print("Operación cancelada")

    # Solicita al usuario como desea buscar el producto
    def tipo_busqueda(self):
        
        # Crear ventana de diálogo
        ventana = tk.Toplevel()
        ventana.title("Buscar producto")
        ventana.geometry("400x200")
        ventana.grab_set()  # Bloquea la ventana principal hasta que se cierre esta

        etiqueta = tk.Label(ventana, text="¿Cómo deseas buscar el producto?")
        etiqueta.pack(pady=10)

        # Botones con comandos que SETEAN el valor en self.eleccion
        btn_id = tk.Button(ventana, text="Por ID", width=20, command=lambda: [self.eleccion.set("id"), ventana.destroy()])
        btn_id.pack(pady=2)

        btn_descripcion = tk.Button(ventana, text="Por Descripcion", width=20, command=lambda: [self.eleccion.set("descripcion"), ventana.destroy()])
        btn_descripcion.pack(pady=2)

        btn_cancelar = tk.Button(ventana, text="Cancelar", width=20, command=lambda: [self.eleccion.set("cancelar"), ventana.destroy()])
        btn_cancelar.pack(pady=2)

        # Espera hasta que el usuario cierre el cuadro de diálogo
        ventana.wait_window()

        # Evaluar la opción después de cerrar el cuadro
        eleccion = self.eleccion.get()
        if eleccion == "id":
            print("Buscar por ID")
            self.buscar_ID()
        elif eleccion == "descripcion":
            print("Buscar por Descripcion")
            self.buscar_decripcion()
        elif eleccion == "cancelar":
            print("Operación cancelada")

        print("Salió de la ventana")
    
    # Solicita al usuario como desea agregar stock
    def tipo_agregar(self):
        
        # Crear ventana de diálogo
        ventana = tk.Toplevel()
        ventana.title("Buscar producto")
        ventana.geometry("400x200")
        ventana.grab_set()  # Bloquea la ventana principal hasta que se cierre esta

        etiqueta = tk.Label(ventana, text="¿Cómo deseas agregar stock?")
        etiqueta.pack(pady=10)

        # Botones con comandos que SETEAN el valor en self.eleccion
        btn_id = tk.Button(ventana, text="Por ID", width=20, command=lambda: [self.eleccion.set("id"), ventana.destroy()])
        btn_id.pack(pady=2)

        btn_descripcion = tk.Button(ventana, text="Por Descripcion", width=20, command=lambda: [self.eleccion.set("descripcion"), ventana.destroy()])
        btn_descripcion.pack(pady=2)

        btn_cancelar = tk.Button(ventana, text="Cancelar", width=20, command=lambda: [self.eleccion.set("cancelar"), ventana.destroy()])
        btn_cancelar.pack(pady=2)

        # Espera hasta que el usuario cierre el cuadro de diálogo
        ventana.wait_window()

        # Evaluar la opción después de cerrar el cuadro
        eleccion = self.eleccion.get()
        if eleccion == "id":
            print("Buscar por ID")
            self.agregar_id()
        elif eleccion == "descripcion":
            print("Buscar por Descripcion")
            self.agregar_descripcion()
        elif eleccion == "cancelar":
            ventana.destroy()
            print("Operación cancelada")

    #busca el producto por ID
    def buscar_ID(self):

        if not self.id_buscar_var.get():
            messagebox.showwarning("Advertencia", "El campo ID es obligatorio.")
            return
            
        try:
            id_producto = int(self.id_buscar_var.get())
            
            # Cargar datos actuales
            tabla_excel = pd.read_excel(self.archivo_excel)
            
            # Buscar producto
            producto = tabla_excel[tabla_excel["ID"] == id_producto]
            
            if len(producto) > 0:
                # Seleccionar en la tabla
                for item in self.tree.get_children():
                    valores = self.tree.item(item, "values")
                    if valores[0] == id_producto:
                        self.tree.selection_set(item)
                        self.tree.focus(item)
                        self.tree.see(item)
                        break

                messagebox.showinfo("Encontrado", f"Producto {id_producto} encontrado\nDescripcion: {producto.iloc[0]['Descripcion']}\nSerie: {producto.iloc[0]['Serie']}\nObservaciones: {producto.iloc[0]['Observaciones']}\nLugar: {producto.iloc[0]['Lugar']}\nCantidad: {producto.iloc[0]['Cantidad']}")
                
            else:
                messagebox.showinfo("No encontrado", f"El producto con ID {id_producto} no existe en el inventario.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al buscar el producto: {str(e)}")

    #busca el producto por descripcion
    def buscar_decripcion(self):

        if not self.descripcion_buscar_var.get():
            messagebox.showwarning("Advertencia", "El campo Descripcion es obligatorio.")
            return
            
        try:
            descripcion_input = self.descripcion_buscar_var.get().strip().lower()

            # Cargar datos actuales
            tabla_excel = pd.read_excel(self.archivo_excel)
            
            # Verificar si el producto ya existe por ID
         
            producto_existente = tabla_excel[tabla_excel["Descripcion"].str.lower() == descripcion_input.lower()]
            
            if not producto_existente.empty:
                indice = producto_existente.index[0]
                messagebox.showinfo("Encontrado", f"Producto encontrado\nID: {producto_existente.iloc[0]['ID']}\nDescripcion: {producto_existente.iloc[0]['Descripcion']}\nSerie: {producto_existente.iloc[0]['Serie']}\nObservaciones: {producto_existente.iloc[0]['Observaciones']}\nLugar: {producto_existente.iloc[0]['Lugar']}\nCantidad: {producto_existente.iloc[0]['Cantidad']}")
            else:
                messagebox.showinfo("No encontrado", f"El producto con decripcion {descripcion_input} no encontrado en el inventario.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al buscar el producto: {str(e)}")
        
    #esto eliminar por id
    def eliminar_id(self):

        if not self.validar_datos_busqueda():
            return
            
        try:
            id_producto = int(self.id_buscar_var.get())
            
            # Cargar datos actuales
            tabla_excel = pd.read_excel(self.archivo_excel)
            
            # Buscar producto
            producto = tabla_excel[tabla_excel["ID"] == id_producto]
            
            if len(producto) > 0:

                # Confirmación
                confirmacion = messagebox.askyesno("Confirmar eliminación", 
                                               f"¿Está seguro de eliminar el producto con ID {id_producto}?")
                if not confirmacion:
                    return

                # Eliminar producto
                tabla_excel = tabla_excel[tabla_excel["ID"] != id_producto]
                
                # Guardar cambios
                tabla_excel.to_excel(self.archivo_excel, index=False)
                
                messagebox.showinfo("Éxito", f"Producto {id_producto} eliminado correctamente.")
                self.actualizar_tabla()
                self.limpiar_campos()
            else:
                messagebox.showinfo("No encontrado", f"El producto con ID {id_producto} no existe en el inventario.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al eliminar el producto: {str(e)}")
    
    #esto eliminar por id
    def eliminar_descripcion(self):

        if not self.descripcion_buscar_var.get():
            messagebox.showwarning("Advertencia", "El campo Descripcion es obligatorio.")
            return
            
        try:
            descripcion_input = self.descripcion_buscar_var.get().strip().lower()

            # Cargar datos actuales
            tabla_excel = pd.read_excel(self.archivo_excel)
            
            # Verificar si el producto ya existe por ID
            producto_existente = tabla_excel[tabla_excel["Descripcion"].str.lower() == descripcion_input.lower()]
            
            
            if not producto_existente.empty:

                # Confirmación
                confirmacion = messagebox.askyesno("Confirmar eliminación", 
                                               f"¿Está seguro de eliminar el producto con descripcion {producto_existente.iloc[0]['Descripcion']}?")
                if not confirmacion:
                    return

                # Eliminar producto
                tabla_excel = tabla_excel[tabla_excel["Descripcion"] != producto_existente.iloc[0]['Descripcion']]
                
                # Guardar cambios
                tabla_excel.to_excel(self.archivo_excel, index=False)
                
                messagebox.showinfo("Éxito", f"Producto {producto_existente.iloc[0]['Descripcion']} eliminado correctamente.")
                self.actualizar_tabla()
                self.limpiar_campos()
            else:
                messagebox.showinfo("No encontrado", f"El producto con ID {descripcion_input} no existe en el inventario.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al eliminar el producto: {str(e)}")

    # Agregar producto nuevo o actualizar cantidad
    def agregar_producto(self):

        if not self.validar_datos():
            return
        
        try:
            descripcion_input = self.descripcion_var.get().strip().lower()
            cantidad = int(self.cantidad_var.get())

            # Cargar datos actuales
            tabla_excel = pd.read_excel(self.archivo_excel)
            
            # Verificar si el producto ya existe por ID
            producto_existente = tabla_excel[tabla_excel["Descripcion"].str.lower() == descripcion_input.lower()]
            
            if not producto_existente.empty:

                # Confirmación
                confirmacion = messagebox.askyesno("Producto Existente", 
                                               f"¿Desea agregar cantidad al producto?-\nDescripcion: {descripcion_input}\nCantidad actual: {producto_existente['Cantidad'].values[0]}")
                if not confirmacion:
                    return
                
                # Producto existe, actualizar cantidad
                indice = producto_existente.index[0]
                tabla_excel.at[indice, "Cantidad"] += cantidad
                messagebox.showinfo("Éxito", f"Producto con ID {indice+1} agregado a la cantidad.")
            else:
                id_producto = int(self.crear_id())

                # Agregar nuevo producto
                nuevo_producto = {
                    "ID": id_producto,
                    "Descripcion": self.descripcion_var.get().lower(),
                    "Serie": self.serie_var.get().lower(),
                    "Observaciones": self.observaciones_var.get().lower(),
                    "Lugar": self.lugar_var.get(),
                    "Cantidad": cantidad
                }

                tabla_excel = pd.concat([tabla_excel, pd.DataFrame([nuevo_producto])], ignore_index=True)
                messagebox.showinfo("Éxito", f"Producto {id_producto} agregado correctamente.")
            
            # Guardar cambios
            tabla_excel.to_excel(self.archivo_excel, index=False)
            
            # Actualizar tabla
            self.actualizar_tabla()
            self.limpiar_campos()
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo agregar el producto: {str(e)}")
    
    # Retirar producto pro id
    def retiro_id(self):
        
        if not self.validar_datos_busqueda():
            return
        
        try:
            id_producto = int(self.id_buscar_var.get())
            
            # Cargar datos actuales
            tabla_excel = pd.read_excel(self.archivo_excel)
            
            # Verificar si el producto existe
            producto_existente = tabla_excel[tabla_excel["ID"] == id_producto]
            
            if len(producto_existente) > 0:
                indice = tabla_excel[tabla_excel["ID"] == id_producto].index[0]
                cantidad_actual = int(tabla_excel.at[indice, "Cantidad"])
                
                 # Confirmación
                confirmacion = messagebox.askyesno("Confirmar Retirar", 
                                               f"¿Está seguro de retirar stock del producto con descripcion {producto_existente.iloc[0]['Descripcion']}?")
                if not confirmacion:
                    return

                # Solicitar cantidad a retirar
                #cantidad_retirar = simpledialog.askinteger(f"Cantidad actual {cantidad_actual}", "Ingrese la cantidad a retirar:", minvalue=1)
                cantidad_retirar = simpledialog.askinteger(title=f"Cantidad actual: {cantidad_actual}", prompt="Ingrese la cantidad a Retirar:", parent=self.root)
                
                if cantidad_retirar is None:
                    return
                
                if cantidad_retirar > cantidad_actual:
                    messagebox.showwarning("Advertencia", f"No hay suficiente cantidad disponible. Disponible: {cantidad_actual}")
                    return
                
                # Actualizar cantidad
                nueva_cantidad = cantidad_actual - cantidad_retirar
                tabla_excel.at[indice, "Cantidad"] = nueva_cantidad
                
                if nueva_cantidad == 0:
                    resultado = messagebox.askyesno("Confirmar", "La cantidad ha llegado a cero. ¿Desea eliminar el producto del inventario?")
                    if resultado:
                        tabla_excel = tabla_excel.drop(indice)
                
                # Guardar cambios
                tabla_excel.to_excel(self.archivo_excel, index=False)
                
                messagebox.showinfo("Éxito", f"Se retiraron {cantidad_retirar} unidades del producto {id_producto}.")
                self.actualizar_tabla()
                self.limpiar_campos()
            else:
                messagebox.showwarning("No encontrado", f"El producto con ID {id_producto} no existe en el inventario.")
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo retirar el producto: {str(e)}")
    
    #Retirar producto por descripcion
    def retiro_descripcion(self):

        if not self.descripcion_buscar_var.get():
            messagebox.showwarning("Advertencia", "El campo Descripcion es obligatorio.")
            return
        
        try:

            descripcion_input = self.descripcion_buscar_var.get().strip().lower()

            # Cargar datos actuales
            tabla_excel = pd.read_excel(self.archivo_excel)
            
            # Verificar si el producto ya existe por ID
            producto_existente = tabla_excel[tabla_excel["Descripcion"].str.lower() == descripcion_input.lower()]
            
            if not producto_existente.empty:

                # Confirmación
                confirmacion = messagebox.askyesno("Confirmar Retirar", 
                                               f"¿Está seguro de retirar stock del producto con descripcion {producto_existente.iloc[0]['Descripcion']}?")
                if not confirmacion:
                    return
                
                # obtiene el id del producto si exite
                indice = producto_existente.index[0]

                cantidad_actual =  int(tabla_excel.at[indice, "Cantidad"])

                # Solicitar cantidad a retirar
                #cantidad_retirar = simpledialog.askinteger(f"Cantidad actual {cantidad_actual}", "Ingrese la cantidad a retirar:", minvalue=1)
                cantidad_retirar = simpledialog.askinteger(title=f"Cantidad actual: {cantidad_actual}", prompt="Ingrese la cantidad a Retirar:", parent=self.root)
                
                if cantidad_retirar is None:
                    return
                
                if cantidad_retirar > cantidad_actual:
                    messagebox.showwarning("Advertencia", f"No hay suficiente cantidad disponible. Disponible: {cantidad_actual}")
                    return
                
                # Actualizar cantidad
                nueva_cantidad = cantidad_actual - cantidad_retirar
                tabla_excel.at[indice, "Cantidad"] = nueva_cantidad
                
                if nueva_cantidad == 0:
                    resultado = messagebox.askyesno("Confirmar", "La cantidad ha llegado a cero. ¿Desea eliminar el producto del inventario?")
                    if resultado:
                        tabla_excel = tabla_excel.drop(indice)
                
                # Guardar cambios
                tabla_excel.to_excel(self.archivo_excel, index=False)
                
                messagebox.showinfo("Éxito", f"Se retiraron {cantidad_retirar} unidades del producto.")

                self.actualizar_tabla()
                self.limpiar_campos()
            else:
                messagebox.showwarning("No encontrado", f"El producto con descripcion {descripcion_input} no existe en el inventario.")
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo retirar el producto: {str(e)}")
    
    # Agregar stock de un producto existente por ID
    def agregar_id(self):

        if not self.validar_datos_busqueda():
            return
            
        try:
            id_producto = int(self.id_buscar_var.get())
            
            # Cargar datos actuales
            tabla_excel = pd.read_excel(self.archivo_excel)
            
            # Buscar producto
            producto = tabla_excel[tabla_excel["ID"] == id_producto]
            
            if len(producto) > 0:
                indice = tabla_excel[tabla_excel["ID"] == id_producto].index[0]
                cantidad_actual = int(tabla_excel.at[indice, "Cantidad"])
                
                # Solicitar cantidad a retirar
                #cantidad_agregar = simpledialog.askinteger(f"Cantidad actual {cantidad_actual}", "Ingrese la cantidad a agregar:", minvalue=1)
                cantidad_agregar = simpledialog.askinteger(title=f"Cantidad actual: {cantidad_actual}", prompt="Ingrese la cantidad a agregar:", parent=self.root)
                
                if cantidad_agregar is None and not isinstance(cantidad_agregar, int):
                    messagebox.showwarning("Advertencia", f"No se ingresó una cantidad válida, { cantidad_agregar } debe ser un número entero.")
                    return
                
                # Actualizar cantidad
                nueva_cantidad = cantidad_actual + cantidad_agregar
                tabla_excel.at[indice, "Cantidad"] = nueva_cantidad
                
                
                # Guardar cambios
                tabla_excel.to_excel(self.archivo_excel, index=False)
                
                messagebox.showinfo("Éxito", f"Se agregaron {cantidad_agregar} unidades del producto {id_producto}.")

                self.actualizar_tabla()
                self.limpiar_campos()
            else:
                messagebox.showwarning("No encontrado", f"El producto con ID {id_producto} no existe en el inventario.")
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo agregar el producto: {str(e)}")

        # Agregar stock de un producto existente

    # Agregar stock de un producto existente por descripcion
    def agregar_descripcion(self):

            if not self.descripcion_buscar_var.get():
                messagebox.showwarning("Advertencia", "El campo Descripcion es obligatorio.")
                return
            
            try:
                descripcion_input = self.descripcion_buscar_var.get().strip().lower()
            
                # Cargar datos actuales
                tabla_excel = pd.read_excel(self.archivo_excel)
            
                # Buscar producto
                producto = tabla_excel[tabla_excel["Descripcion"].str.lower() == descripcion_input.lower()]
            
                if not producto.empty:
                
                    indice = producto.index[0]
                    cantidad_actual = int(tabla_excel.at[indice, "Cantidad"])
                
                    # Solicitar cantidad a retirar
                    #cantidad_agregar = simpledialog.askinteger(f"Cantidad actual {cantidad_actual}", "Ingrese la cantidad a agregar:", minvalue=1)
                    cantidad_agregar = simpledialog.askinteger(title=f"Cantidad actual: {cantidad_actual}", prompt="Ingrese la cantidad a agregar:", parent=self.root)
                
                    if cantidad_agregar is None and not isinstance(cantidad_agregar, int):
                        messagebox.showwarning("Advertencia", f"No se ingresó una cantidad válida, { cantidad_agregar } debe ser un número entero.")
                        return
                
                    # Actualizar cantidad
                    nueva_cantidad = cantidad_actual + cantidad_agregar
                    tabla_excel.at[indice, "Cantidad"] = nueva_cantidad
                

                    # Guardar cambios
                    tabla_excel.to_excel(self.archivo_excel, index=False)
                
                    messagebox.showinfo("Éxito", f"Se agregaron {cantidad_agregar} unidades del producto {descripcion_input}.")
                    self.actualizar_tabla()
                    self.limpiar_campos()
                else:
                    messagebox.showwarning("No encontrado", f"El producto con ID {descripcion_input} no existe en el inventario.")
                
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo agregar el producto: {str(e)}")

    #crea un id unico para cada producto
    def crear_id(self):
        try:
            # Cargar datos actuales
            df = pd.read_excel(self.archivo_excel)
            
            # Generar nuevo ID
            if df.empty:
                nuevo_id = 1
            else:
                nuevo_id = df["ID"].max() + 1
            
            return nuevo_id
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar ID: {str(e)}")

def main():
    root = tk.Tk()
    app = InventarioApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
