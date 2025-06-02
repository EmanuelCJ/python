import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

# Ruta específica donde guardar el Excel
RUTA_DIRECTORIO = r"C:\Users\arsas\Desktop\sisStock"
os.makedirs(RUTA_DIRECTORIO, exist_ok=True)
EXCEL_FILE = os.path.join(RUTA_DIRECTORIO, "stock_productos.xlsx")

# Clase Producto
class Producto:
    _id_counter = 1

    def __init__(self, descripcion, cantidad):
        self.ID = Producto._id_counter
        Producto._id_counter += 1
        self.Descripcion = descripcion
        self.Cantidad = cantidad

    def sumarCantidad(self, cantidad):
        self.Cantidad += cantidad

    def restarCantidad(self, cantidad):
        if cantidad <= self.Cantidad:
            self.Cantidad -= cantidad
        else:
            raise ValueError("Cantidad a restar supera el stock disponible")

    def to_dict(self):
        return {"ID": self.ID, "Descripcion": self.Descripcion, "Cantidad": self.Cantidad}


# Clase Stock
class Stock:
    def __init__(self):
        self.listaProducto = []
        self.cargarDesdeExcel()

    def agregarProducto(self, producto):
        self.listaProducto.append(producto)
        self.guardarEnExcel()

    def quitarProducto(self, producto_id):
        self.listaProducto = [p for p in self.listaProducto if p.ID != producto_id]
        self.guardarEnExcel()

    def listarProductos(self):
        return "\n".join(str(p.__dict__) for p in self.listaProducto)

    def guardarEnExcel(self):
        try:
            print(f"Guardando en: {EXCEL_FILE}")
            df = pd.DataFrame([p.to_dict() for p in self.listaProducto])
            df.to_excel(EXCEL_FILE, index=False)
            print("Archivo guardado correctamente.")
        except Exception as e:
            print(f"Error al guardar el archivo: {e}")

    def cargarDesdeExcel(self):
        if os.path.exists(EXCEL_FILE):
            df = pd.read_excel(EXCEL_FILE)
            self.listaProducto = []
            for _, row in df.iterrows():
                prod = Producto(row['Descripcion'], int(row['Cantidad']))
                prod.ID = int(row['ID'])
                self.listaProducto.append(prod)
            if not df.empty:
                Producto._id_counter = df['ID'].max() + 1

    def buscarProducto(self, producto_id):
        for producto in self.listaProducto:
            if producto.ID == producto_id:
                return producto
        return None


# Interfaz gráfica con Tkinter
class App:
    def __init__(self, root):
        self.stock = Stock()
        root.title("Sistema de Stock")

        self.tabControl = ttk.Notebook(root)
        self.frame_nuevo = ttk.Frame(self.tabControl)
        self.frame_admin = ttk.Frame(self.tabControl)

        self.tabControl.add(self.frame_nuevo, text='Cargar Producto Nuevo')
        self.tabControl.add(self.frame_admin, text='Administrar Elementos')
        self.tabControl.pack(expand=1, fill="both")

        self.init_frame_nuevo()
        self.init_frame_admin()

    def init_frame_nuevo(self):
        tk.Label(self.frame_nuevo, text="Descripción:").grid(row=0, column=0)
        self.descripcion_entry = tk.Entry(self.frame_nuevo)
        self.descripcion_entry.grid(row=0, column=1)

        tk.Label(self.frame_nuevo, text="Cantidad:").grid(row=1, column=0)
        self.cantidad_entry = tk.Entry(self.frame_nuevo)
        self.cantidad_entry.grid(row=1, column=1)

        tk.Button(self.frame_nuevo, text="Agregar Producto", command=self.agregar_producto).grid(row=2, column=0, columnspan=2)

    def init_frame_admin(self):
        tk.Label(self.frame_admin, text="ID Producto:").grid(row=0, column=0)
        self.id_admin_entry = tk.Entry(self.frame_admin)
        self.id_admin_entry.grid(row=0, column=1)

        tk.Label(self.frame_admin, text="Cantidad:").grid(row=1, column=0)
        self.cant_admin_entry = tk.Entry(self.frame_admin)
        self.cant_admin_entry.grid(row=1, column=1)

        tk.Button(self.frame_admin, text="Sumar", command=self.sumar_elemento).grid(row=2, column=0)
        tk.Button(self.frame_admin, text="Restar", command=self.restar_elemento).grid(row=2, column=1)
        tk.Button(self.frame_admin, text="Listar Productos", command=self.listar_productos).grid(row=3, column=0, columnspan=2)

        self.text_area = tk.Text(self.frame_admin, height=10, width=50)
        self.text_area.grid(row=4, column=0, columnspan=2)

    def agregar_producto(self):
        desc = self.descripcion_entry.get()
        try:
            cant = int(self.cantidad_entry.get())
            producto = Producto(desc, cant)
            self.stock.agregarProducto(producto)
            messagebox.showinfo("Éxito", f"Producto agregado con ID {producto.ID}")
            self.descripcion_entry.delete(0, tk.END)
            self.cantidad_entry.delete(0, tk.END)
        except ValueError:
            messagebox.showerror("Error", "Cantidad debe ser un número entero")

    def sumar_elemento(self):
        try:
            id_ = int(self.id_admin_entry.get())
            cant = int(self.cant_admin_entry.get())
            prod = self.stock.buscarProducto(id_)
            if prod:
                prod.sumarCantidad(cant)
                self.stock.guardarEnExcel()
                messagebox.showinfo("Éxito", f"Se sumaron {cant} unidades al producto ID {id_}")
            else:
                messagebox.showerror("Error", "Producto no encontrado")
        except ValueError:
            messagebox.showerror("Error", "Entradas inválidas")

    def restar_elemento(self):
        try:
            id_ = int(self.id_admin_entry.get())
            cant = int(self.cant_admin_entry.get())
            prod = self.stock.buscarProducto(id_)
            if prod:
                try:
                    prod.restarCantidad(cant)
                    self.stock.guardarEnExcel()
                    messagebox.showinfo("Éxito", f"Se restaron {cant} unidades al producto ID {id_}")
                except ValueError as ve:
                    messagebox.showerror("Error", str(ve))
            else:
                messagebox.showerror("Error", "Producto no encontrado")
        except ValueError:
            messagebox.showerror("Error", "Entradas inválidas")

    def listar_productos(self):
        self.text_area.delete(1.0, tk.END)
        productos = self.stock.listarProductos()
        self.text_area.insert(tk.END, productos)


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
