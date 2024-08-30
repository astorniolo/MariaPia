import pandas as pd
import mysql.connector
from datetime import datetime

# Conexión a la base de datos MySQL
conn = mysql.connector.connect(
    host='localhost',
    user='root',
    password='hepatalgina',
    database='mariapia'
)
cursor = conn.cursor()

# Leer el archivo Excel
archivo_excel = 'C:\workspace\PyCN\Model\INSUMOS.xls'
hojas = pd.ExcelFile(archivo_excel).sheet_names  # Obtener los nombres de las hojas

for hoja in hojas:
    if hoja != "Hoja interna":

        df = pd.read_excel(archivo_excel, sheet_name=hoja)
        
        for index, row in df.iterrows():
            nombre = row['NOMBRE']
            volumen = row['VOLUMEN']
            precio_por_mayor = row['PRECIO POR']
            descripcion = row['DESCRIPCION']
            if (hoja=="Barras Glicerina" or hoja=="Tintura" or hoja=="Hierbas" or  hoja=="Laboratorio" or hoja=="Comestible" ):
                presentacion=""
            else:    
                presentacion = row['PRESENTACION']
            tipo = hoja
            
            nombre_producto=nombre
            tipo_producto=tipo

            # Verificar si el producto ya existe en la base de datos
            cursor.execute("SELECT id FROM Productos WHERE nombre_producto = %s AND tipo_producto = %s", (nombre_producto, tipo_producto))
            resultado = cursor.fetchone()

            if resultado:
                # Producto ya existe, actualizar precio
                cursor.execute("""
                    UPDATE Productos 
                    SET precio = %s, fecha_actualizacion = %s 
                    WHERE id = %s
                """, (precio_producto, datetime.now().date(), resultado[0]))
            else:
                # Producto no existe, insertarlo
                cursor.execute("""
                    INSERT INTO Productos (tipo_producto, nombre_producto, precio, fecha_actualizacion)
                    VALUES (%s, %s, %s, %s)
                """, (tipo_producto, nombre_producto, precio_producto, datetime.now().date()))

        conn.commit()  # Aplicar cambios en la base de datos

# Cerrar la conexión a la base de datos
cursor.close()
conn.close()
        


