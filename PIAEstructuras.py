import random as rd
import sys
import datetime
import sqlite3 
from sqlite3 import Error
from datetime import (date, datetime, timedelta)
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

def Crear_tabla():
    try:
        conn = sqlite3.connect('Ventas_DelSol.db') 
        mi_cursor = conn.cursor()
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS Productos (idProducto INTEGER PRIMARY KEY AUTOINCREMENT, nombreProducto TEXT NOT NULL, precioProducto INT NOT NULL, existe TEXT DEFAULT 'Activo')")
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS Sucursales (idSucursal INTEGER PRIMARY KEY AUTOINCREMENT, nombreSucursal TEXT NOT NULL, direccionSucursal TEXT NOT NULL, telefonoSucursal INT NOT NULL, existe TEXT DEFAULT 'Activo')")
        mi_cursor.execute("CREATE TABLE IF NOT EXISTS Ventas (idVenta INTEGER PRIMARY KEY AUTOINCREMENT, producto TEXT NOT NULL, sucursal TEXT NOT NULL, cantidadProducto INT NOT NULL, costoProducto INT NOT NULL, costoTotal INT NOT NULL, fecha TEXT NOT NULL, existe TEXT DEFAULT 'Activo')")
        print("La base de datos Ventas_DelSol y las tablas Productos, Sucursales y Ventas se han creado correctamente")
    except sqlite3.Error as e:
        print(e)
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

Crear_tabla()

def mostrar_menu():
    while True:
        print("\nBase de Datos VENTAS DELSOL:")
        print("1. Registrar")
        print("2. Borrar")
        print("3. Editar")
        print("4. Leer")
        print("5. Exportar a Excel")
        print("6. Salir del Programa")
        
        opcion = input("Selecciona una opción: ")

        if opcion == "1":
            print(" ")
            mostrar_menu_registrar()
        elif opcion == "2":
            print(" ")
            mostrar_menu_borrar()
        elif opcion == "3":
            print(" ")
            editar_registro()
        elif opcion == "4":
            print(" ")
            mostrar_menu_leer()
        elif opcion == "5":
            print(" ")
            exportar_a_excel()
        elif opcion == "6":
            print(" ")
            print(f"Hasta luego 😃")
            break
        else:
            print("ERROR: Opción no válida. Intentar nuevamente.")

def mostrar_menu_registrar():
    while True:
        print("Menú de Registro:")
        print("1. Registrar Producto")
        print("2. Registrar Sucursal")
        print("3. Registrar Venta")
        print("4. Volver al Menú Principal")
        
        opcion = input("Selecciona una opción: ")

        if opcion == "1":
            registrar_producto()
        elif opcion == "2":
            registrar_sucursal()
        elif opcion == "3":
            registrar_venta()
        elif opcion == "4":
            break
        else:
            print("ERROR: Opción no válida. Intentar nuevamente.")

def mostrar_menu_borrar():
    while True:
        print("Menú de Borrar:")
        print("1. Borrar Producto")
        print("2. Borrar Sucursal")
        print("3. Borrar Venta")
        print("4. Volver al Menú Principal")

        opcion = input("Selecciona una opción: ")

        if opcion == "1":
            borrar_registro("Productos")
        elif opcion == "2":
            borrar_registro("Sucursales")
        elif opcion == "3":
            borrar_registro("Ventas")
        elif opcion == "4":
            break
        else:
            print("ERROR: Opción no válida. Intentar nuevamente.")

def borrar_registro(tabla):
    try:
        conn = sqlite3.connect('Ventas_DelSol.db')
        mi_cursor = conn.cursor()

        print(f"\nBorrar {tabla}:")
        id_registro = input(f"Ingrese el ID de {tabla} que desea borrar: ")

        # Verificar si el registro existe
        if tabla == 'Productos':
            mi_cursor.execute(f"SELECT existe FROM {tabla} WHERE idProducto = ?", (id_registro,))
        elif tabla == 'Sucursales':
            mi_cursor.execute(f"SELECT existe FROM {tabla} WHERE idSucursal = ?", (id_registro,))
        elif tabla == 'Ventas':
            mi_cursor.execute(f"SELECT existe FROM {tabla} WHERE idVenta= ?", (id_registro,))
        else:
            print('Error: tabla no existe')

        resultado = mi_cursor.fetchone()

        if resultado:
            existe_actual = resultado[0]

            if existe_actual == 'Si': # Cambiar el existe a 'No'
                if tabla == 'Productos':
                    mi_cursor.execute(f"UPDATE Productos SET existe = 'No' WHERE idProducto = ?", (id_registro,))
                    print(f"Producto con ID {id_registro} borrado correctamente.")
                elif tabla == 'Sucursales':
                    mi_cursor.execute(f"UPDATE Sucursales SET existe = 'No' WHERE idSucursal = ?", (id_registro,))
                    print(f"Sucursal con ID {id_registro} borrada correctamente.")
                elif tabla == 'Ventas':
                    mi_cursor.execute(f"UPDATE Ventas SET existe = 'No' WHERE idVenta = ?", (id_registro,))
                    print(f"Venta con ID {id_registro} borrada correctamente.")
                else:
                    print('Error: tabla no existe')
            else:
                # El registro ya está Inactivo, dar la opción de reactivarlo
                reactivar = input(f"El registro ya está Inactivo. ¿Reactivar? (S/N): ").upper()
                if reactivar == 'S':
                    if tabla == 'Productos':
                        mi_cursor.execute(f"UPDATE Productos SET existe = 'Si' WHERE idProducto = ?", (id_registro,))
                        print(f"Producto con ID {id_registro} reactivado correctamente.")    
                    elif tabla == 'Sucursales':
                        mi_cursor.execute(f"UPDATE Sucursales SET existe = 'Si' WHERE idSucursal = ?", (id_registro,))
                        print(f"Sucursal con ID {id_registro} reactivado correctamente.")
                    elif tabla == 'Ventas':
                        mi_cursor.execute(f"UPDATE Ventas SET existe = 'Si' WHERE idVenta = ?", (id_registro,))
                        print(f"Ventas con ID {id_registro} reactivado correctamente.")
                    else:
                        print('Error: tabla no existe')
                else:
                    print(f"El registro con ID {id_registro} no fue reactivado.")

        else:
            print(f"No existe {tabla} con ID {id_registro}.")

        conn.commit()

    except sqlite3.Error as e:
        print(e)
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

def editar_registro():
    try:
        conn = sqlite3.connect('Ventas_DelSol.db')
        mi_cursor = conn.cursor()
        texto_editar = input("Escribe el comando sqlite a realizar: ")
        mi_cursor.execute(f"{texto_editar}")
        print("Comando ejecutado en sqlite con éxito.")
        conn.commit()

    except sqlite3.Error as e:
        print(e)
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

def mostrar_menu_leer():
    while True:
        print("Menú de Leer:")
        print("1. Ver todos los Productos")
        print("2. Ver todas las Sucursales")
        print("3. Ver todas las Ventas")
        print("4. Volver al Menú Principal")

        opcion = input("Selecciona una opción: ")

        if opcion == "1":
            leer_tabla("Productos")
        elif opcion == "2":
            leer_tabla("Sucursales")
        elif opcion == "3":
            leer_tabla("Ventas")
        elif opcion == "4":
            break
        else:
            print("ERROR: Opción no válida. Intentar nuevamente.")

def leer_tabla(tabla):
    try:
        conn = sqlite3.connect('Ventas_DelSol.db')
        mi_cursor = conn.cursor()

        # Realizar un SELECT * en la tabla seleccionada
        mi_cursor.execute(f"SELECT * FROM {tabla}")
        filas = mi_cursor.fetchall()

        if filas:
            # Mostrar los resultados
            print(f"\nRegistros en la tabla {tabla}:")
            for fila in filas:
                print(fila)
        else:
            print(f"No hay registros en la tabla {tabla}.")

    except sqlite3.Error as e:
        print(e)
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")
    finally:
        if 'conn' in locals():
            conn.close()
    print(" ")

def exportar_a_excel():
    while True:
        print("Menú de Exportar a Excel:")
        print("1. Exportar Productos a Excel")
        print("2. Exportar Sucursales a Excel")
        print("3. Exportar Ventas a Excel")
        print("4. Volver al Menú Principal")

        opcion = input("Selecciona una opción: ")

        if opcion == "1":
            exportar_excel_tabla("Productos")
        elif opcion == "2":
            exportar_excel_tabla("Sucursales")
        elif opcion == "3":
            exportar_excel_tabla("Ventas")
        elif opcion == "4":
            break
        else:
            print("ERROR: Opción no válida. Intentar nuevamente.")
            print(" ")

def exportar_excel_tabla(tabla):
    try:
        conn = sqlite3.connect('Ventas_DelSol.db')
        df = pd.read_sql_query(f"SELECT * FROM {tabla}", conn)

        if not df.empty:
            # Crear un nuevo libro de Excel y escribir los datos
            libro = Workbook()
            hoja = libro.active

            for fila in dataframe_to_rows(df, index=False, header=True):
                hoja.append(fila)

            # Guardar el archivo Excel
            libro.save(f'{tabla}.xlsx')
            print(f"Datos de la tabla {tabla} exportados a {tabla}.xlsx correctamente.")
        else:
            print(f"No hay registros en la tabla {tabla}.")

    except sqlite3.Error as e:
        print(e)
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")
    finally:
        if 'conn' in locals():
            conn.close()
    print(" ")

def validar_precio(mensaje):
    while True:
        try:
            precio = float(input(mensaje))
            if precio >= 0:
                return precio
            else:
                print("ERROR: El precio debe ser un número positivo.")
        except ValueError:
            print("ERROR: Ingrese un número válido.")

def validar_fecha(mensaje):
    while True:
        try:
            fecha = input(mensaje)
            datetime.strptime(fecha, '%d/%m/%Y')
            return fecha
        except ValueError:
            print("ERROR: Formato de fecha incorrecto. Ingrese la fecha en formato DD/MM/AAAA.")

def registrar_producto():
    try:
        conn = sqlite3.connect('Ventas_DelSol.db') 
        mi_cursor = conn.cursor()
        nombre_producto = input("Ingrese el nombre del producto: ")
        precio_producto = validar_precio("Ingrese el precio del producto: ")
        mi_cursor.execute("INSERT INTO Productos (nombreProducto, precioProducto, existe) VALUES (?, ?, 'Si')", (nombre_producto, precio_producto))
        print("Producto registrado correctamente.")
        conn.commit()
    except sqlite3.Error as e:
        print(e)
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

def registrar_sucursal():
    try:
        conn = sqlite3.connect('Ventas_DelSol.db') 
        mi_cursor = conn.cursor()
        nombre_sucursal = input("Ingrese el nombre de la sucursal: ")
        direccion_sucursal = input("Ingrese la dirección de la sucursal: ")
        telefono_sucursal = int(input("Ingrese el teléfono de la sucursal: "))
        mi_cursor.execute("INSERT INTO Sucursales (nombreSucursal, direccionSucursal, telefonoSucursal, existe) VALUES (?, ?, ?, 'Si')", (nombre_sucursal, direccion_sucursal, telefono_sucursal))
        print("Sucursal registrada correctamente.")
        conn.commit()
    except sqlite3.Error as e:
        print(e)
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

def registrar_venta():
    try:
        conn = sqlite3.connect('Ventas_DelSol.db') 
        mi_cursor = conn.cursor()
        producto = input("Ingrese el nombre del producto vendido: ")
        sucursal = input("Ingrese el nombre de la sucursal: ")
        cantidad_producto = int(input("Ingrese la cantidad de productos vendidos: "))
        costo_producto = validar_precio("Ingrese el costo unitario del producto: ")
        costo_total = cantidad_producto * costo_producto
        fecha = validar_fecha("Ingrese la fecha de la venta (formato: DD/MM/AAAA): ")
        mi_cursor.execute("INSERT INTO Ventas (producto, sucursal, cantidadProducto, costoProducto, costoTotal, fecha, existe) VALUES (?, ?, ?, ?, ?, ?, 'Si')", (producto, sucursal, cantidad_producto, costo_producto, costo_total, fecha))
        print("Venta registrada correctamente.")
        conn.commit()
    except sqlite3.Error as e:
        print(e)
    except Exception as e:
        print(f"Se produjo el siguiente error: {e}")
    finally:
        if 'conn' in locals():
            conn.close()

mostrar_menu()
