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
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")
        
        # Archivo Excel
        self.archivo_excel = "inventario.xlsx"
        self.inicializar_excel()
        
        # Variables para los campos
        self.id_var = tk.IntVar()
        self.descripcion_var = tk.StringVar()
        self.serie_var = tk.IntVar()
        self.observaciones_var = tk.StringVar()
        self.lugar_var = tk.StringVar()
        self.cantidad_var = tk.IntVar()

        # variables para campos de busqueda
        self.id_buscar_var = tk.StringVar()
        
        # Crear interfaz
        self.crear_interfaz()
        
        # Cargar datos iniciales en la tabla
        self.actualizar_tabla()
    
    def inicializar_excel(self):
        # Verificar si el archivo existe
        if not os.path.exists(self.archivo_excel):
            # Crear un nuevo archivo Excel con las columnas necesarias
            df = pd.DataFrame(columns=[
                "ID", "Descripción", "Serie", 
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
                    "ID", "Descripción", "Serie", 
                    "Observaciones", "Lugar", "Cantidad"
                ])
                df.to_excel(self.archivo_excel, index=False)
                
    def crear_interfaz(self):
        # Frame principal
        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Frame para los campos de entrada
        input_frame = tk.LabelFrame(main_frame, text="Datos del Producto", bg="#f0f0f0", font=("Arial", 12))
        input_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Crear campos de entrada
        self.crear_campo(input_frame, "Descripción:", self.descripcion_var, 1)
        self.crear_campo(input_frame, "No. Serie:", self.serie_var, 2)
        self.crear_campo(input_frame, "Observaciones:", self.observaciones_var, 3)
        self.crear_campo(input_frame, "Lugar:", self.lugar_var, 4)
        self.crear_campo(input_frame, "Cantidad:", self.cantidad_var, 5)
        
         # Frame para los campos de busqueda
        input_frame = tk.LabelFrame(main_frame, text="Campo busqueda", bg="#f0f0f0", font=("Arial", 12))
        input_frame.pack(fill=tk.X, padx=10, pady=10)

        self.crear_campo(input_frame, "buscar ID:", self.id_buscar_var, 0)

        # Frame para botones
        button_frame = tk.Frame(main_frame, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Botones de acción
        self.crear_boton(button_frame, "Ingreso Producto", self.agregar_producto, "#2ecc71", 0)
        self.crear_boton(button_frame, "Agregar Stock", self.agregar_stock, "#9b59b6", 1)
        self.crear_boton(button_frame, "Retirar", self.retirar_producto, "#e74c3c", 2)
        self.crear_boton(button_frame, "Eliminar", self.eliminar_producto, "#e67e22", 4)
        self.crear_boton(button_frame, "Buscar", self.buscar_producto, "#3498db", 5)

        # Crear tabla de visualización
        table_frame = tk.LabelFrame(main_frame, text="Inventario Actual", bg="#f0f0f0", font=("Arial", 12))
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Configurar tabla
        self.tree = ttk.Treeview(table_frame, columns=("ID", "Descripción", "Serie", "Observaciones", "Lugar", "Cantidad"), show="headings")
        
        # Configurar columnas
        self.tree.heading("ID", text="ID")
        self.tree.heading("Descripción", text="Descripción")
        self.tree.heading("Serie", text="No. Serie")
        self.tree.heading("Observaciones", text="Observaciones")
        self.tree.heading("Lugar", text="Lugar")
        self.tree.heading("Cantidad", text="Cantidad")
        
        # Ajustar anchos de columnas
        self.tree.column("ID", width=80)
        self.tree.column("Descripción", width=200)
        self.tree.column("Serie", width=100)
        self.tree.column("Observaciones", width=150)
        self.tree.column("Lugar", width=100)
        self.tree.column("Cantidad", width=80)
        
        # Barra de desplazamiento
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        # Empaquetar tabla y scrollbar
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Evento de selección de la tabla
        self.tree.bind("<ButtonRelease-1>", self.seleccionar_item)
    
    # Crear campos de entrada
    def crear_campo(self, parent, texto, variable, fila):
        tk.Label(parent, text=texto, bg="#f0f0f0", font=("Arial", 12)).grid(row=fila, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(parent, textvariable=variable, font=("Arial", 12), width=30).grid(row=fila, column=1, padx=10, pady=5, sticky="w")
    
    # Crear campo de búsqueda
    def crear_campo_busqueda(self, parent, texto, variable, fila):
        tk.Label(parent, text=texto, bg="#f0f0f0", font=("Arial", 12)).grid(row=fila, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(parent, textvariable=variable, font=("Arial", 12), width=30).grid(row=fila, column=1, padx=10, pady=5, sticky="w")
    
    def crear_boton(self, parent, texto, comando, color, columna):
        tk.Button(
            parent, 
            text=texto, 
            command=comando, 
            bg=color, 
            fg="white", 
            font=("Arial", 10, "bold"),
            width=12,
            height=2,
            relief=tk.RAISED
        ).grid(row=0, column=columna, padx=5, pady=5)
    
    def limpiar_campos(self):
        self.id_var.set("")
        self.descripcion_var.set("")
        self.serie_var.set("")
        self.observaciones_var.set("")
        self.lugar_var.set("")
        self.cantidad_var.set("")
    
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
                    row.get("Descripción", ""), 
                    row.get("Serie", ""), 
                    row.get("Observaciones", ""), 
                    row.get("Lugar", ""), 
                    row.get("Cantidad", "")
                ))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar los datos: {str(e)}")
    
    def seleccionar_item(self, event):
        # Obtener el ítem seleccionado
        seleccion = self.tree.selection()
        if seleccion:
            item = self.tree.item(seleccion[0])
            valores = item['values']
            
            # Actualizar los campos con los valores seleccionados
            self.id_var.set(valores[0])
            self.descripcion_var.set(valores[1])
            self.serie_var.set(valores[2])
            self.observaciones_var.set(valores[3])
            self.lugar_var.set(valores[4])
            self.cantidad_var.set(valores[5])
    
    # Validar datos de entrada
    def validar_datos(self):
        
        # Validar que no haya campos vacíos
        if not self.serie_var or not self.cantidad_var or not self.descripcion_var or not self.lugar_var:
           messagebox.showwarning("Advertencia", "El campo Descripción es obligatorio.")
           return False
        
        # Validar que ID, Serie y Cantidad sean enteros
        if not isinstance(self.serie_var.get(), int) or not isinstance(self.cantidad_var.get(), int):
            messagebox.showwarning( "Advertencia","Serie y Cantidad deben ser números enteros.")
            return False
        
        # Validar que Descripción y Lugar sean strings
        if not isinstance(self.descripcion_var.get(), str) or not isinstance(self.lugar_var.get(), str):
            messagebox.showwarning("Descripción y Lugar deben ser texto.")
            return False
        
        # Validar que la cantidad sea un número
        try:
            cantidad = self.cantidad_var.get()
            if cantidad < 0:
                messagebox.showwarning("Advertencia", "La cantidad no puede ser negativa.")
                return False
        except ValueError:
            messagebox.showwarning("Advertencia", "La cantidad debe ser un número entero.")
            return False
        
        return True
    
    # Validar datos de búsqueda
    def validar_datos_busqueda(self):
        # Validar que no haya campos vacíos
        if not self.id_buscar_var.get():
            messagebox.showwarning("Advertencia", "El campo ID es obligatorio.")
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

    def agregar_producto(self):

        if not self.validar_datos():
            return
        
        try:
            id_producto = int(self.crear_id())
            cantidad = int(self.cantidad_var.get())
            
            # Cargar datos actuales
            tabla_excel = pd.read_excel(self.archivo_excel)
            
            # Verificar si el producto ya existe por ID
            producto_existente = tabla_excel[tabla_excel["ID"] == id_producto]
            
            if len(producto_existente) > 0:
                # Producto existe, actualizar cantidad
                indice = tabla_excel[tabla_excel["ID"] == id_producto].index[0]
                tabla_excel.at[indice, "Cantidad"] = int(tabla_excel.at[indice, "Cantidad"]) + cantidad
                messagebox.showinfo("Éxito", f"Se actualizó la cantidad del producto {id_producto}.")
            else:
                # Agregar nuevo producto
                nuevo_producto = {
                    "ID": id_producto,
                    "Descripción": self.descripcion_var.get(),
                    "Serie": self.serie_var.get(),
                    "Observaciones": self.observaciones_var.get(),
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
    
    def retirar_producto(self):

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
                
                # Solicitar cantidad a retirar
                cantidad_retirar = simpledialog.askinteger(f"Cantidad actual {cantidad_actual}", "Ingrese la cantidad a retirar:", minvalue=1)
                
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

    def agregar_stock(self):

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
                
                # Solicitar cantidad a retirar
                cantidad_agregar = simpledialog.askinteger(f"Cantidad actual {cantidad_actual}", "Ingrese la cantidad a retirar:", minvalue=1)
                
                if cantidad_agregar is None:
                    return
                
                if cantidad_agregar > cantidad_actual:
                    messagebox.showwarning("Advertencia", f"No hay suficiente cantidad disponible. Disponible: {cantidad_actual}")
                    return
                
                # Actualizar cantidad
                nueva_cantidad = cantidad_actual + cantidad_agregar
                tabla_excel.at[indice, "Cantidad"] = nueva_cantidad
                
                if nueva_cantidad == 0:
                    resultado = messagebox.askyesno("Confirmar", "La cantidad ha llegado a cero. ¿Desea eliminar el producto del inventario?")
                    if resultado:
                        tabla_excel = tabla_excel.drop(indice)
                
                # Guardar cambios
                tabla_excel.to_excel(self.archivo_excel, index=False)
                
                messagebox.showinfo("Éxito", f"Se retiraron {cantidad_agregar} unidades del producto {id_producto}.")
                self.actualizar_tabla()
                self.limpiar_campos()
            else:
                messagebox.showwarning("No encontrado", f"El producto con ID {id_producto} no existe en el inventario.")
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo retirar el producto: {str(e)}")
    
    def buscar_producto(self):

        if not self.validar_datos_busqueda():
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

                messagebox.showinfo("Encontrado", f"Producto {id_producto} encontrado\nDescripción: {producto.iloc[0]['Descripción']}\nSerie: {producto.iloc[0]['Serie']}\nObservaciones: {producto.iloc[0]['Observaciones']}\nLugar: {producto.iloc[0]['Lugar']}\nCantidad: {producto.iloc[0]['Cantidad']}")
                
            else:
                messagebox.showinfo("No encontrado", f"El producto con ID {id_producto} no existe en el inventario.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al buscar el producto: {str(e)}")
    
    def eliminar_producto(self):

        if not self.validar_datos_busqueda():
            return
            
        try:
            id_producto = self.id_buscar_var.get()
            
            # Confirmación
            confirmacion = messagebox.askyesno("Confirmar eliminación", 
                                               f"¿Está seguro de eliminar el producto con ID {id_producto}?")
            if not confirmacion:
                return
            
            # Cargar datos actuales
            tabla_excel = pd.read_excel(self.archivo_excel)
            
            # Verificar si el producto existe
            producto_existente = tabla_excel[tabla_excel["ID"] == id_producto]
            
            if len(producto_existente) > 0:
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
    
    def compactar_hoja(self):
        try:
            # Cargar datos actuales
            df = pd.read_excel(self.archivo_excel)
            
            # Eliminar filas vacías
            df_compacto = df.dropna(how='all')
            
            # Guardar cambios
            df_compacto.to_excel(self.archivo_excel, index=False)
            
            messagebox.showinfo("Éxito", "Hoja compactada correctamente.")
            self.actualizar_tabla()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error al compactar la hoja: {str(e)}")

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
