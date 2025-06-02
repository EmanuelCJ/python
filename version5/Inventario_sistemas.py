import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
import pandas as pd
#esta libreria es para manejar archivos csv y excel muy grandes
import dask.dataframe as dd

class InventarioSistemasApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("Inventario de Sistemas")
        self.root.geometry("800x600")
        self.root.configure(bg="#1E1E1E")

    # Archivo Excel
        self.archivo_csv = "inventario.csv"
        self.inicializar_excel()

        # Variables para los campos de entrada
        self.id_var = tk.IntVar()
        self.descripcion_var = tk.StringVar()
        self.serie_var = tk.StringVar()
        self.observaciones_var = tk.StringVar()
        self.lugar_var = tk.StringVar()
        self.cantidad_var = tk.StringVar()

        # variables para campos de busqueda
        self.id_buscar_var = tk.StringVar()
        self.descripcion_buscar_var = tk.StringVar()

        # Variable para la elección de búsqueda
        self.eleccion = tk.StringVar(value="")  # Valor por defecto

        # Crear interfaz
        self.crear_interfaz()

        #Cargar datos iniciales en la tabla
        self.actualizar_tabla()

    def inicializar_excel(self):
        # Verificar si el archivo existe
        if not os.path.exists(self.archivo_csv):
            # Crear un nuevo archivo Excel con las columnas necesarias
            df = pd.DataFrame(columns=[
                "ID", "Descripcion", "Serie", 
                "Observaciones", "Lugar", "Cantidad"
            ])
            #
            df.to_csv(self.archivo_csv, index=False)
            print(f"Archivo {self.archivo_csv} creado correctamente.")
        else:
            try:
                # Verificar que el archivo es accesible y tiene el formato correcto
                ddf = dd.read_csv(self.archivo_csv)
                ddf.head()

                print(f"Archivo {self.archivo_csv} cargado correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"El archivo de inventario está dañado o no es accesible: {str(e)}")
                # Crear un backup y un nuevo archivo
                if os.path.exists(self.archivo_csv):
                    os.rename(self.archivo_csv, f"{self.archivo_csv}.bak")
                df = dd.DataFrame(columns=[
                    "ID", "Descripcion", "Serie", 
                    "Observaciones", "Lugar", "Cantidad"
                ])
                df.to_csv(self.archivo_csv, index=False)

    def crear_interfaz(self):

        # Frame principal
        main_frame = tk.Frame(self.root, bg="#1E1E1E")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Crear tabla de visualización
        table_frame = tk.LabelFrame(main_frame, text="Vista previa del Inventario Actual",  font=("Arial", 16, "bold"), bg="#fefefe", fg="#2c3e50")
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

         # Frame para botones
        button_frame = tk.Frame(main_frame, bg="#1E1E1E")
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Botones de acción
        self.crear_boton(button_frame, "Ingreso Producto", self.agregar_producto, "#2ecc71", 0)
        self.crear_boton(button_frame, "Agregar Stock", self.agregar_stock, "#9b59b6", 1)
        self.crear_boton(button_frame, "Retirar", self.retirar_producto, "#e74c3c", 2)
        self.crear_boton(button_frame, "Eliminar", self.eliminar_producto, "#e67e22", 4)
        self.crear_boton(button_frame, "Buscar", self.tipo_busqueda, "#3498db", 5)

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

      # Crear campos de entrada
    
    def crear_campo(self, parent, texto, variable, fila):
        tk.Label(parent, text=texto, bg="#f0f0f0", font=("Arial", 12)).grid(row=fila, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(parent, textvariable=variable, font=("Arial", 12), width=30).grid(row=fila, column=1, padx=10, pady=5, sticky="w")
    

    def agregar_producto(self):

        ventana = tk.Toplevel()
        ventana.title("Agregar nuevo producto")
        ventana.geometry("600x300")
        ventana.grab_set()

        # Frame para los campos de entrada
        input_frame = tk.LabelFrame(ventana, text="Datos del Producto",  font=("Arial", 12, "bold"), bg="#fefefe", fg="#2c3e50")
        input_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Crear campos de entrada
        self.crear_campo(input_frame, "Descripción:", self.descripcion_var, 1)
        self.crear_campo(input_frame, "N°. Serie:", self.serie_var, 2)
        self.crear_campo(input_frame, "Observaciones:", self.observaciones_var, 3)
        self.crear_campo(input_frame, "Lugar:", self.lugar_var, 4)
        self.crear_campo(input_frame, "Cantidad:", self.cantidad_var, 5)

        # Frame de botones
        button_frame = tk.Frame(ventana, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        self.crear_boton(button_frame, "Nuevo Producto", self.nuevo_producto, "#2ecc71", 0)
        self.crear_boton(button_frame, "Cancelar", ventana.destroy, "#e74c3c", 1)

    def agregar_stock(self):
        ventana = tk.Toplevel()
        ventana.title("Agregar stock de producto")
        ventana.geometry("600x300")
        ventana.grab_set()
        # Frame para los campos de entrada
        input_frame = tk.LabelFrame(ventana, text="Campos de Agregado",  font=("Arial", 12, "bold"), bg="#fefefe", fg="#2c3e50")
        input_frame.pack(fill=tk.X, padx=10, pady=10)

        # Crear campos de entrada
        self.crear_campo(input_frame, "Por ID:", self.id_buscar_var, 1)
        self.crear_campo(input_frame, "Por Descripcion:", self.descripcion_buscar_var, 2)

        # Frame de botones
        button_frame = tk.Frame(ventana, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        self.crear_boton(button_frame, "Agregar Stock", self.agregar_producto, "#2ecc71", 0)
        self.crear_boton(button_frame, "Cancelar", ventana.destroy, "#e74c3c", 1)


        # Espera hasta que el usuario cierre el cuadro de diálogo
        ventana.wait_window()

    def retirar_producto(self):
        ventana = tk.Toplevel()
        ventana.title("Retirar stock de producto")
        ventana.geometry("600x300")
        ventana.grab_set()

        # Frame para los campos de entrada
        input_frame = tk.LabelFrame(ventana, text="Campos de Retiro", font=("Arial", 12, "bold"), bg="#fefefe", fg="#2c3e50")
        input_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Crear campos de entrada
        self.crear_campo(input_frame, "Por ID:", self.id_buscar_var, 1)
        self.crear_campo(input_frame, "Por Descripcion:", self.descripcion_buscar_var, 2)
        

        # Frame de botones
        button_frame = tk.Frame(ventana, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, padx=10, pady=10)

        self.crear_boton(button_frame, "Retirar Stock", self.agregar_producto, "#2ecc71",1)
        self.crear_boton(button_frame, "Cancelar", ventana.destroy, "#e74c3c", 2)

        # Espera hasta que el usuario cierre el cuadro de diálogo
        ventana.wait_window()

    def eliminar_producto(self):
        ventana = tk.Toplevel()
        ventana.title("Eliminar un producto")
        ventana.geometry("600x300")
        ventana.grab_set()

        # Frame para los campos de entrada
        input_frame = tk.LabelFrame(ventana, text="Campos de Eliminación", font=("Arial", 12, "bold"), bg="#fefefe", fg="#2c3e50")
        input_frame.pack(fill=tk.X, padx=10, pady=10)

        # Crear campos de entrada
        self.crear_campo(input_frame, "Por ID:", self.id_buscar_var, 1)
        self.crear_campo(input_frame, "Por Descripcion:", self.descripcion_buscar_var, 2)
        
        button_frame = tk.Frame(ventana, bg="#f0f0f0")
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.crear_boton(button_frame, "Eliminar Producto", self.nuevo_producto, "#2ecc71", 0)
        self.crear_boton(button_frame, "Cancelar", ventana.destroy, "#e74c3c", 1)

        # Espera hasta que el usuario cierre el cuadro de diálogo
        ventana.wait_window()

    def tipo_busqueda(self):
        ventana = tk.Toplevel()
        ventana.title("Buscar producto")
        ventana.geometry("600x300")
        ventana.grab_set()
        
         # Frame para los campos de entrada
        input_frame = tk.LabelFrame(ventana, text="Campos de busqueda", font=("Arial", 12, "bold"), bg="#fefefe", fg="#2c3e50")
        input_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Crear campos de entrada
        self.crear_campo(input_frame, "Por ID:", self.id_buscar_var, 1)
        self.crear_campo(input_frame, "Por Descripcion:", self.descripcion_buscar_var, 2)

        # Frame de búsqueda
        frame_busqueda = tk.Frame(ventana, bg="#f0f0f0")
        frame_busqueda.pack(fill=tk.X, padx=10, pady=10)

         # Botones de acción
        self.crear_boton(frame_busqueda, "Por ID", self.buscar_id, "#2ecc71", 0)
        self.crear_boton(frame_busqueda, "Por Descripcion", self.buscar_descripcion, "#9b59b6", 1)
        self.crear_boton(frame_busqueda, "Cancelar", ventana.destroy, "#e74c3c", 3)

        
        # Espera hasta que el usuario cierre el cuadro de diálogo
        ventana.wait_window()

    def buscar_id(self):
        pass

    def buscar_descripcion(self):
        pass

    def nuevo_producto(self):

        if not self.validar_datos():
            return

        try:
            id_producto = int(self.crear_id())
            cantidad = int(self.cantidad_var.get())
            serie = int(self.serie_var.get())

            # Cargar datos actuales con Dask data frame
            ddf = dd.read_csv(self.archivo_csv)
            df = ddf.compute()  # Convertir a pandas para manipulación inmediata

            # Verificar si el producto ya existe por ID
            resultado = df[df["Descripcion"].str.contains(self.descripcion_var, case=False, na=False)]

            if not resultado.empty:
                # Producto existe, actualizar cantidad
                indice = resultado.index[0]
                df.at[indice, "Cantidad"] += cantidad
                messagebox.showinfo("Éxito", f"Se actualizó la cantidad del producto {id_producto}.")
            else:
                # Agregar nuevo producto
                nuevo_producto = {
                    "ID": id_producto,
                    "Descripción": self.descripcion_var.get().lower(),
                    "Serie": serie,
                    "Observaciones": self.observaciones_var.get().lower(),
                    "Lugar": self.lugar_var.get(),
                    "Cantidad": cantidad
                }
                df = pd.concat([df, pd.DataFrame([nuevo_producto])], ignore_index=True)
                messagebox.showinfo("Éxito", f"Producto {id_producto} agregado correctamente.")

                # Guardar cambios (sobrescribe el archivo)
                df.to_csv(self.archivo_csv, index=False)

                # Actualizar tabla e interfaz
                self.actualizar_tabla()
                self.limpiar_campos()

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {e}")

    def crear_id(self):
        try:
            # Cargar datos actuales
            df = dd.read_csv(self.archivo_csv).compute()
            
            # Generar nuevo ID
            if df.empty:
                nuevo_id = 1
            else:
                nuevo_id = df["ID"].max() + 1
            
            return nuevo_id
        except Exception as e:
            messagebox.showerror("Error", f"Error al generar ID: {str(e)}")
    
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
            df = dd.read_csv(self.archivo_csv).compute()
            
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
            self.id_buscar_var.set(valores[0])
            self.descripcion_var.set(valores[1])
            self.serie_var.set(valores[2])
            self.observaciones_var.set(valores[3])
            self.lugar_var.set(valores[4])
            self.cantidad_var.set(valores[5])

    # Validar datos de entrada
    def validar_datos(self):
        
        # Validar que no haya campos vacíos
        if not self.serie_var or not self.cantidad_var or not self.descripcion_var.get().strip() or not self.lugar_var.get().strip():
           messagebox.showwarning("Advertencia", "El campo los campos son obligatorios.")
           return False
        
        # Validar que Descripción y Lugar sean strings
        if not isinstance(self.descripcion_var.get(), str) or not isinstance(self.lugar_var.get(), str) or not isinstance(self.serie_var.get(), str):
            messagebox.showwarning("Descripción y Lugar deben ser texto.")
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
        
def main():
    root = tk.Tk()
    app = InventarioSistemasApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()