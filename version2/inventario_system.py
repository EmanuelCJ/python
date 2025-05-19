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
        self.root.title("Sistema de Inventario")
        self.root.geometry("800x600")
        self.root.configure(bg="#f0f0f0")
        
        # Archivo Excel
        self.archivo_excel = "inventario.xlsx"
        self.inicializar_excel()
        
        # Variables para los campos
        self.id_var = tk.StringVar()
        self.descripcion_var = tk.StringVar()
        self.serie_var = tk.StringVar()
        self.observaciones_var = tk.StringVar()
        self.lugar_var = tk.StringVar()
        self.cantidad_var = tk.StringVar()
        
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
        self.crear_campo(input_frame, "ID:", self.id_var, 0)
        self.crear_campo(input_frame, "Descripción:", self.descripcion_var, 1)
        self.crear_campo(input_frame, "No. Serie:", self.serie_var, 2)
        self.crear_campo(input_frame, "Observaciones:", self.observaciones_var, 3)
        self.crear_campo(input_frame, "Lugar:", self.lugar_var, 4)
        self.crear_campo(input_frame, "Cantidad:", self.cantidad_var, 5)
        
        # Frame para botones
        button_frame = tk.Frame(main_frame, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Botones de acción
        self.crear_boton(button_frame, "Agregar", self.agregar_producto, "#2ecc71", 0)
        self.crear_boton(button_frame, "Retirar", self.retirar_producto, "#e74c3c", 1)
        self.crear_boton(button_frame, "Buscar", self.buscar_producto, "#3498db", 2)
        self.crear_boton(button_frame, "Eliminar", self.eliminar_producto, "#e67e22", 3)
        self.crear_boton(button_frame, "Compactar", self.compactar_hoja, "#9b59b6", 4)
        
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
    
    def crear_campo(self, parent, texto, variable, fila):
        tk.Label(parent, text=texto, bg="#f0f0f0", font=("Arial", 10)).grid(row=fila, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(parent, textvariable=variable, font=("Arial", 10), width=30).grid(row=fila, column=1, padx=10, pady=5, sticky="w")
    
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
    
    def validar_datos(self):
        # Validar que los campos necesarios no estén vacíos
        if not self.id_var.get().strip():
            messagebox.showwarning("Advertencia", "El campo ID es obligatorio.")
            return False
        
        if not self.descripcion_var.get().strip():
            messagebox.showwarning("Advertencia", "El campo Descripción es obligatorio.")
            return False
        
        # Validar que la cantidad sea un número
        try:
            cantidad = int(self.cantidad_var.get())
            if cantidad < 0:
                messagebox.showwarning("Advertencia", "La cantidad no puede ser negativa.")
                return False
        except ValueError:
            messagebox.showwarning("Advertencia", "La cantidad debe ser un número entero.")
            return False
        
        return True
    
    def agregar_producto(self):
        if not self.validar_datos():
            return
        
        try:
            id_producto = self.id_var.get().strip()
            cantidad = int(self.cantidad_var.get())
            
            # Cargar datos actuales
            df = pd.read_excel(self.archivo_excel)
            
            # Verificar si el producto ya existe por ID
            producto_existente = df[df["ID"] == id_producto]
            
            if len(producto_existente) > 0:
                # Producto existe, actualizar cantidad
                indice = df[df["ID"] == id_producto].index[0]
                df.at[indice, "Cantidad"] = int(df.at[indice, "Cantidad"]) + cantidad
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
                df = pd.concat([df, pd.DataFrame([nuevo_producto])], ignore_index=True)
                messagebox.showinfo("Éxito", f"Producto {id_producto} agregado correctamente.")
            
            # Guardar cambios
            df.to_excel(self.archivo_excel, index=False)
            
            # Actualizar tabla
            self.actualizar_tabla()
            self.limpiar_campos()
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo agregar el producto: {str(e)}")
    
    def retirar_producto(self):
        if not self.id_var.get().strip():
            messagebox.showwarning("Advertencia", "Debe especificar un ID para retirar producto.")
            return
            
        try:
            id_producto = self.id_var.get().strip()
            
            # Cargar datos actuales
            df = pd.read_excel(self.archivo_excel)
            
            # Verificar si el producto existe
            producto_existente = df[df["ID"] == id_producto]
            
            if len(producto_existente) > 0:
                indice = df[df["ID"] == id_producto].index[0]
                cantidad_actual = int(df.at[indice, "Cantidad"])
                
                # Solicitar cantidad a retirar
                cantidad_retirar = simpledialog.askinteger("Retirar", "Ingrese la cantidad a retirar:", minvalue=1)
                
                if cantidad_retirar is None:
                    return
                
                if cantidad_retirar > cantidad_actual:
                    messagebox.showwarning("Advertencia", f"No hay suficiente cantidad disponible. Disponible: {cantidad_actual}")
                    return
                
                # Actualizar cantidad
                nueva_cantidad = cantidad_actual - cantidad_retirar
                df.at[indice, "Cantidad"] = nueva_cantidad
                
                if nueva_cantidad == 0:
                    resultado = messagebox.askyesno("Confirmar", "La cantidad ha llegado a cero. ¿Desea eliminar el producto del inventario?")
                    if resultado:
                        df = df.drop(indice)
                
                # Guardar cambios
                df.to_excel(self.archivo_excel, index=False)
                
                messagebox.showinfo("Éxito", f"Se retiraron {cantidad_retirar} unidades del producto {id_producto}.")
                self.actualizar_tabla()
                self.limpiar_campos()
            else:
                messagebox.showwarning("No encontrado", f"El producto con ID {id_producto} no existe en el inventario.")
                
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo retirar el producto: {str(e)}")
    
    def buscar_producto(self):
        if not self.id_var.get().strip():
            messagebox.showwarning("Advertencia", "Debe especificar un ID para buscar.")
            return
            
        try:
            id_producto = self.id_var.get().strip()
            
            # Cargar datos actuales
            df = pd.read_excel(self.archivo_excel)
            
            # Buscar producto
            producto = df[df["ID"] == id_producto]
            
            if len(producto) > 0:
                # Seleccionar en la tabla
                for item in self.tree.get_children():
                    valores = self.tree.item(item, "values")
                    if valores[0] == id_producto:
                        self.tree.selection_set(item)
                        self.tree.focus(item)
                        self.tree.see(item)
                        break
                
                # Actualizar campos
                fila = producto.iloc[0]
                self.id_var.set(fila["ID"])
                self.descripcion_var.set(fila["Descripción"])
                self.serie_var.set(fila["Serie"])
                self.observaciones_var.set(fila["Observaciones"])
                self.lugar_var.set(fila["Lugar"])
                self.cantidad_var.set(str(fila["Cantidad"]))
                
                messagebox.showinfo("Encontrado", f"Producto {id_producto} encontrado.")
            else:
                messagebox.showinfo("No encontrado", f"El producto con ID {id_producto} no existe en el inventario.")
                
        except Exception as e:
            messagebox.showerror("Error", f"Error al buscar el producto: {str(e)}")
    
    def eliminar_producto(self):
        if not self.id_var.get().strip():
            messagebox.showwarning("Advertencia", "Debe especificar un ID para eliminar.")
            return
            
        try:
            id_producto = self.id_var.get().strip()
            
            # Confirmación
            confirmacion = messagebox.askyesno("Confirmar eliminación", 
                                               f"¿Está seguro de eliminar el producto con ID {id_producto}?")
            if not confirmacion:
                return
            
            # Cargar datos actuales
            df = pd.read_excel(self.archivo_excel)
            
            # Verificar si el producto existe
            producto_existente = df[df["ID"] == id_producto]
            
            if len(producto_existente) > 0:
                # Eliminar producto
                df = df[df["ID"] != id_producto]
                
                # Guardar cambios
                df.to_excel(self.archivo_excel, index=False)
                
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

def main():
    root = tk.Tk()
    app = InventarioApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
