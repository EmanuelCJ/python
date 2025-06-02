import unittest
import pandas as pd
import os
from inventario_system import InventarioApp

class TestInventarioApp(unittest.TestCase):
    def setUp(self):
        self.archivo = "inventario.xlsx"
        
        # Crear archivo con 1000 artículos de prueba
        data = [{
            "ID": i + 1,
            "Descripcion": f"producto_{i}",
            "Serie": f"serie_{i}",
            "Observaciones": f"obs_{i}",
            "Lugar": f"lugar_{i}",
            "Cantidad": i + 10
        } for i in range(10000)]

        df = pd.DataFrame(data)
        df.to_excel(self.archivo, index=False)

    def test_archivo_creado(self):
        self.assertTrue(os.path.exists(self.archivo), "El archivo no fue creado correctamente.")

    def test_cantidad_productos(self):
        df = pd.read_excel(self.archivo)
        self.assertEqual(len(df), 10000, "No hay 1000 productos en el inventario.")

    def test_buscar_producto_existente(self):
        df = pd.read_excel(self.archivo)
        producto = df[df["Descripcion"] == "producto_500"]
        self.assertFalse(producto.empty, "No se encontró producto_500.")

    def test_agregar_nuevo_producto(self):
        df = pd.read_excel(self.archivo)
        nuevo_producto = {
            "ID": 10001,
            "Descripcion": "nuevo_producto",
            "Serie": "NP123",
            "Observaciones": "nuevo ingreso",
            "Lugar": "almacen",
            "Cantidad": 20
        }
        df = pd.concat([df, pd.DataFrame([nuevo_producto])], ignore_index=True)
        df.to_excel(self.archivo, index=False)

        df_actualizado = pd.read_excel(self.archivo)
        self.assertEqual(len(df_actualizado), 10001, "No se agregó el nuevo producto correctamente.")

    def test_eliminar_producto(self):
        df = pd.read_excel(self.archivo)
        df = df[df["Descripcion"] != "producto_10"]
        df.to_excel(self.archivo, index=False)

        df_post = pd.read_excel(self.archivo)
        self.assertFalse((df_post["Descripcion"] == "producto_10").any(), "No se eliminó producto_10.")

    # def tearDown(self):
    #     # Limpiar después de las pruebas
    #     if os.path.exists(self.archivo):
    #         os.remove(self.archivo)

if __name__ == '__main__':
    unittest.main()
