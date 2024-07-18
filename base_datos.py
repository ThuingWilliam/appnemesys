import pyodbc as db
import os
from datetime import datetime
import pandas as pd
paramConexion = 'dsn=NEMESYSDB;us=SYSDBA;cont=masterkey'
conexiondb = db.connect(paramConexion)

def registrar_error(texto_error, archivo_log=r'validar/bdError_log.txt'):

    try:
        # La fecha y hora del error, importante
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Formatear el mensaje de error, por si acaso
        mensaje_error = f"{timestamp} - ERROR - {texto_error}\n"
        
        # Escribir el mensaje de error en el archivo
        with open(archivo_log, 'a') as file:
            file.write(mensaje_error)
    except Exception as e:
        # En caso de error al registrar el error, imprimir el error en la consola
        print(f"Error al registrar el error: {str(e)}")

def Departamento_id():
    cursor = conexiondb.cursor()
    cursor.execute("SELECT DISTINCT NOMBRE_DEPARTAMENTO FROM departamento")
    resultado = cursor.fetchall()
    cursor.close()
    return [dep[0] for dep in resultado]

def Proveedor_id():
    cursor = conexiondb.cursor()
    cursor.execute("SELECT DISTINCT EMPRESA FROM proveedor")
    resultado = cursor.fetchall()
    cursor.close()
    return [pro[0] for pro in resultado]

def Localidad():
    cursor = conexiondb.cursor()
    cursor.execute("SELECT DISTINCT ubicacion FROM LOCALIDAD")
    resultado = cursor.fetchall()
    cursor.close()
    return [loc[0] for loc in resultado]

def UnidadMedida():
    cursor = conexiondb.cursor()
    cursor.execute("SELECT DISTINCT descripcion FROM UNIDADES_MEDIDAS")
    resultado = cursor.fetchall()
    cursor.close()
    return [uni[0] for uni in resultado]

def Impuesto():
    cursor = conexiondb.cursor()
    cursor.execute("SELECT DISTINCT impuesto FROM IMPUESTOS")
    resultado = cursor.fetchall()
    cursor.close()
    return [imp[0] for imp in resultado]

def itemmax():
    cursor = conexiondb.cursor()
    cursor.execute("SELECT MAX(codigo_item) FROM inventario")
    resultado = cursor.fetchone()
    cursor.close()
    return resultado[0] if resultado else None

def crear_proveedor_default():
    cursor = None
    try:
        cursor = conexiondb.cursor()
        # Verificar si existe el proveedor con nombre 'PROVEEDOR DEFAULT'
        cursor.execute("SELECT COUNT(*) FROM PROVEEDOR WHERE EMPRESA = 'PROVEEDOR DEFAULT'")
        count = cursor.fetchone()[0]
        # Verificar si existe el proveedor con ID=100
        cursor.execute("SELECT COUNT(*) FROM PROVEEDOR WHERE ID = 100")
        id_100_exists = cursor.fetchone()[0]
        
        if count == 0:
            # Crear el proveedor con ID=100 si no existe
            if id_100_exists == 0:
                id_proveedor="101"
                id_proveedor=str(id_proveedor)
                fecha = datetime.now().strftime('%m/%d/%Y')
                cursor.execute(f"INSERT INTO PROVEEDOR (COD_PROVEEDOR, EMPRESA, MONEDA, FECHA) VALUES ({str(id_proveedor)}, 'PROVEEDOR DEFAULT', 'EFECTIVO', '{fecha}')")
                print(f'Proveedor PROVEEDOR DEFAULT creado con COD_PROVEEDOR={str(id_proveedor)}')
            else:
                # Si ID=100 ya existe, buscar el siguiente ID disponible
                cursor.execute("SELECT MAX(ID) FROM PROVEEDOR")
                max_id = cursor.fetchone()[0]
                id_proveedor = max_id + 1
                fecha = datetime.now().strftime('%m/%d/%Y')
                cursor.execute(f"INSERT INTO PROVEEDOR (COD_PROVEEDOR, EMPRESA, MONEDA, FECHA) VALUES ({str(id_proveedor)}, 'PROVEEDOR DEFAULT', 'EFECTIVO', '{fecha}')")
                print(f'Proveedor PROVEEDOR DEFAULT creado con COD_PROVEEDOR={str(id_proveedor)}')
        else:
            # Obtener el ID del proveedor existente
            cursor.execute("SELECT ID FROM PROVEEDOR WHERE EMPRESA = 'PROVEEDOR DEFAULT'")
            id_proveedor = cursor.fetchone()[0]
            print(f'El proveedor PROVEEDOR DEFAULT ya existe con ID={str(id_proveedor)}')

        # SE PONDRAN LOS NOMBRE DE LA BD EN MAYUSCULA PROVEEDOR Y EMPRESA
        cursor.execute(f"UPDATE DEPARTAMENTO SET NOMBRE_DEPARTAMENTO = UPPER(NOMBRE_DEPARTAMENTO)")
        cursor.execute(f"UPDATE INVENTARIO SET PROVEEDOR_ID = {str(id_proveedor)} WHERE PROVEEDOR_ID IS NULL")
        cursor.execute(f"UPDATE PROVEEDOR SET EMPRESA = UPPER(EMPRESA)")
        conexiondb.commit()  # Guarda los cambios en la base de datos
        print('Actualización exitosa')
    except Exception as e:
        print(f'Error: {e}')
        registrar_error(f"esta es la funcion crear_proveedor_default: {e}")
        conexiondb.rollback()  # Revierte los cambios en caso de error
    finally:
        if cursor is not None:
            cursor.close()

def InsertProvedor(proveedor):
    cursor = conexiondb.cursor()
    proveedor = proveedor.upper()
    moneda = "EFECTIVO"
    timestamp = datetime.now().strftime('%m/%d/%y')

    # Verificar si PROVEEDOR DEFAULT existe
    cursor.execute('SELECT COUNT(*) FROM PROVEEDOR WHERE EMPRESA = ?', ("PROVEEDOR DEFAULT",))
    existe_default = cursor.fetchone()[0]

    if not existe_default:
        # Crear PROVEEDOR DEFAULT con COD_PROVEEDOR 100
        cursor.execute("INSERT INTO PROVEEDOR (COD_PROVEEDOR, EMPRESA, MONEDA, FECHA) VALUES (?, ?, ?, ?)",
                       (100, "PROVEEDOR DEFAULT", moneda, timestamp))
        conexiondb.commit()
        print('PROVEEDOR DEFAULT insertado')

    if not proveedor:
        proveedor = "PROVEEDOR DEFAULT"

    # Verificar si el proveedor ya existe
    cursor.execute('SELECT COUNT(*) FROM PROVEEDOR WHERE EMPRESA = ?', (proveedor,))
    existe = cursor.fetchone()[0]

    if existe:
        print(f'El proveedor "{proveedor}" ya existe.')
        cursor.execute('SELECT ID FROM PROVEEDOR WHERE EMPRESA = ?', (proveedor,))
        existe = cursor.fetchone()[0]
        cursor.close()
        return str(existe)
    
    # Obtener el siguiente valor para COD_PROVEEDOR
    cursor.execute('SELECT MAX(COD_PROVEEDOR) FROM PROVEEDOR')
    resultado = cursor.fetchone()
    cod_proveedor = (int(resultado[0]) + 1) if resultado[0] is not None else 100
    Cod_proveedor = str(cod_proveedor)
    
    try:
        cursor.execute("INSERT INTO PROVEEDOR (COD_PROVEEDOR, EMPRESA, MONEDA, FECHA) VALUES (?, ?, ?, ?)",
                       (Cod_proveedor, proveedor, moneda, timestamp))
        conexiondb.commit()
        
        cursor.execute('SELECT ID FROM PROVEEDOR WHERE EMPRESA = ?', (proveedor,))
        id = cursor.fetchone()[0]
        print('Proveedor insertado')
        return str(id)
    except Exception as e:
        registrar_error(f"Error en la función InsertProveedor: {e}")
        conexiondb.rollback()
    finally:
        cursor.close()

def InsertDepartamento(departamento):
    # Convertir el parámetro a mayúsculas
    departamento = departamento.upper()

    cursor = conexiondb.cursor()

    # Reemplazar 'nan' con 'DEFAULT DEPARTAMENTO'
    if departamento == "NAN":
        departamento = 'DEFAULT DEPARTAMENTO'

    # Verificar si el departamento ya existe
    cursor.execute('SELECT COUNT(*) FROM DEPARTAMENTO WHERE NOMBRE_DEPARTAMENTO = ?', (departamento,))
    existe = cursor.fetchone()[0]
    
    if existe:
        print(f'El departamento "{departamento}" ya existe.')
        cursor.execute('SELECT ID FROM DEPARTAMENTO WHERE NOMBRE_DEPARTAMENTO = ?', (departamento,))
        exi = cursor.fetchone()[0]
        cursor.close()
        return str(exi)
    
    # Obtener el siguiente valor para NO_DEPART
    cursor.execute('SELECT MAX(NO_DEPART) FROM DEPARTAMENTO')
    resultado = cursor.fetchone()
    no_depart = (int(resultado[0]) + 1) if resultado[0] is not None else 100
    no_depart = str(no_depart)
    color = 'clWindow'
    
    try:
        # Insertar el nuevo departamento
        cursor.execute(
            'INSERT INTO DEPARTAMENTO (NO_DEPART, NOMBRE_DEPARTAMENTO, MOSTRAR_EN_COLOR) VALUES (?, ?, ?)',
            (no_depart, departamento, color)
        )
        conexiondb.commit()

        # Obtener el ID del departamento recién insertado
        cursor.execute('SELECT ID FROM DEPARTAMENTO WHERE NOMBRE_DEPARTAMENTO = ?', (departamento,))
        id_departamento = cursor.fetchone()[0]
        print('Departamento insertado')
        return str(id_departamento)
    except Exception as e:
        registrar_error(f"Error en función InsertDepartamento: {e}")
        conexiondb.rollback()
    finally:
        cursor.close()

def DepartamentoIdVacio():
    cursor = None
    
    try:
        cursor = conexiondb.cursor()
        # Verificar si existe el departamento con nombre 'Default_Departamento'
        cursor.execute("SELECT COUNT(*) FROM DEPARTAMENTO WHERE NOMBRE_DEPARTAMENTO = 'DEFAULT DEPARTAMENTO'")
        count = cursor.fetchone()[0]
        # Verificar si existe el departamento con ID=100
        cursor.execute("SELECT COUNT(*) FROM DEPARTAMENTO WHERE ID = 100")
        id_100_exists = cursor.fetchone()[0]
        if count == 0:
            # Crear el departamento con ID=100 si no existe
            if id_100_exists == 0:
                id_departamento = 100
                colorbd='clWindow'
                cursor.execute(f"INSERT INTO DEPARTAMENTO (ID, NO_DEPART, NOMBRE_DEPARTAMENTO,MOSTRAR_EN_COLOR) VALUES ({id_departamento}, '{id_departamento}', 'DEFAULT DEPARTAMENTO','{colorbd}')")
                print(f'Departamento DEFAULT DEPARTAMENTO creado con ID={id_departamento}')
            else:
                # Si ID=100 ya existe, buscar el siguiente ID disponible
                cursor.execute("SELECT MAX(ID) FROM DEPARTAMENTO")
                max_id = cursor.fetchone()[0]
                id_departamento = max_id + 1
                cursor.execute(f"INSERT INTO DEPARTAMENTO (ID, NO_DEPART, NOMBRE_DEPARTAMENTO) VALUES ({id_departamento}, '{id_departamento}', 'DEFAULT DEPARTAMENTO')")
                print(f'Departamento DEFAULT DEPARTAMENTO creado con ID={id_departamento}')
        else:
            # Obtener el ID del departamento existente
            cursor.execute("SELECT ID FROM DEPARTAMENTO WHERE NOMBRE_DEPARTAMENTO = 'DEFAULT DEPARTAMENTO'")
            id_departamento = cursor.fetchone()[0]
            print(f'El departamento DEFAULT DEPARTAMENTO ya existe con ID={id_departamento}')

        # Actualizar DEPARTAMENTO_ID a id_departamento donde está NULL
        cursor.execute(f"UPDATE INVENTARIO SET DEPARTAMENTO_ID = {id_departamento} WHERE DEPARTAMENTO_ID IS NULL")
        proveedor="PROVEEDOR DEFAULT"
        cursor.execute('SELECT ID FROM PROVEEDOR WHERE EMPRESA = ?', (proveedor,))
        resultado = cursor.fetchone()

        if resultado is None:
            print(f'No se encontró el proveedor "{proveedor}"')

        else:
            existe = resultado[0]
            print(existe)
            cursor.execute(f"UPDATE INVENTARIO SET PROVEEDOR_ID = {str(existe)} WHERE PROVEEDOR_ID IS NULL")
            cursor.execute(f"UPDATE PROVEEDOR SET EMPRESA = UPPER(EMPRESA)")
        # SE PONDRAN LOS NOMBRE DE LA BD EN MAYUSCULA DEPARTAMENTO Y PROVEEDOR
        cursor.execute(f"UPDATE DEPARTAMENTO SET NOMBRE_DEPARTAMENTO = UPPER(NOMBRE_DEPARTAMENTO)")
        conexiondb.commit()  # Guarda los cambios en la base de datos
        print('Actualización exitosa')
    except Exception as e:
        print(f'Error: {e}')
        registrar_error(f"esta es la funcion Departamento Vacio: {e}")
        conexiondb.rollback()  # Revierte los cambios en caso de error
    finally:
        if cursor is not None:
            cursor.close()
        crear_proveedor_default()
        departamentoupdateinv()
        proveedorupdateinv()
def departamentoupdateinv():
    try:
        cursor = conexiondb.cursor()
        cursor.execute("SELECT DEPARTAMENTO_ID, ID FROM inventario")
        inventario_datos = cursor.fetchall()
        # Validar y capturar datos tipo string
        departamentos_strings = [row for row in inventario_datos if isinstance(row[0], str)]
        for row in departamentos_strings:
            departamento_id = row[0]
            inventario_id = row[1]
            # Hacer una consulta para obtener el nuevo ID del departamento basado en las primeras 4 letras del campo NOMBRE_DEPARTAMENTO
            cursor.execute("SELECT ID FROM departamento WHERE SUBSTRING(NOMBRE_DEPARTAMENTO FROM 1 FOR 4) = ?", (departamento_id,))
            departamento_datos = cursor.fetchone()
            if departamento_datos:
                departamento_id_nuevo = departamento_datos[0]
                # Actualizar el departamento_id en la tabla inventario solo donde coincida el ID
                cursor.execute("UPDATE inventario SET DEPARTAMENTO_ID = ? WHERE ID = ?", (str(departamento_id_nuevo), inventario_id))
                print("actualización completada")
        # Confirmar los cambios
        conexiondb.commit()
    except Exception as e:
        registrar_error(f"error en la funcion departamentoupdateinv {e}")
    finally:
        # Cerrar la conexión
        cursor.close()
def proveedorupdateinv():
    try:
        cursor = conexiondb.cursor()
        cursor.execute("SELECT PROVEEDOR_ID, ID FROM inventario")
        inventario_datos = cursor.fetchall()
        # Validar y capturar datos tipo string
        proveedores_strings = [row for row in inventario_datos if isinstance(row[0], str)]
        for row in proveedores_strings:
            proveedor_id = row[0]
            inventario_id = row[1]
            # Hacer una consulta para obtener el nuevo ID del proveedor basado en las primeras 4 letras del campo EMPRESA
            cursor.execute("SELECT ID FROM proveedor WHERE SUBSTRING(EMPRESA FROM 1 FOR 4) = ?", (proveedor_id,))
            proveedor_datos = cursor.fetchone()
            if proveedor_datos:
                proveedor_id_nuevo = proveedor_datos[0]
                # Actualizar el proveedor_id en la tabla inventario solo donde coincida el ID
                cursor.execute("UPDATE inventario SET PROVEEDOR_ID = ? WHERE ID = ?", (str(proveedor_id_nuevo), inventario_id))
                print("actualización completada")
        # Confirmar los cambios
        conexiondb.commit()
    except Exception as e:
        registrar_error(f"error en la funcion proveedorupdateinv {e}")
    finally:
        # Cerrar la conexión
        cursor.close()

def inv_Unidad_medidad():
    try:
        cursor = conexiondb.cursor()
        
        # Realizar una inserción de registros si no existen previamente
        cursor.execute("""
            INSERT INTO INVENTARIO_UND_MEDIDAS
            (
                IDINVENTARIOUND,
                ITEM_ID,
                UND_VENTA,
                COMP_VENTA,
                UND_COMPRA,
                COMP_COMPRA,
                UND_ALMACENAMIENTO,
                COMP_ALMACENAMIENTO
            )
            SELECT 
                GEN_ID(IDINVENTARIOUND_GEN, 1),
                CODIGO_ITEM, 
                'UND', 
                1, 
                'UND', 
                1, 
                'UND', 
                1
            FROM INVENTARIO
            WHERE CODIGO_ITEM > 1 
              AND NOT EXISTS (
                SELECT 1
                FROM INVENTARIO_UND_MEDIDAS
                WHERE ITEM_ID = INVENTARIO.CODIGO_ITEM
            )
        """)
        
        conexiondb.commit()
        print("Datos actualizados Unidad Medidas")
    except db.IntegrityError as e:
        print(f"Error de integridad al insertar datos: {e}")
        registrar_error(e)
    except Exception as e:
        print(f"Error general al insertar datos: {e}")
        registrar_error(f"error en la funcion unidad medidad {e}")
    finally:
        if cursor:
            cursor.close()

def Existencia():
    try:
        cursor = conexiondb.cursor()
        
        # Obtener el valor máximo de CODIGO_ITEM de la tabla INVENTARIO
        cursor.execute("SELECT MAX(CODIGO_ITEM) FROM INVENTARIO")
        itemaximo = cursor.fetchone()
        invItemc = itemaximo[0] if itemaximo and itemaximo[0] is not None else 999
        
        # Obtener el valor máximo de ITEM_ID de la tabla EXISTENCIAS
        cursor.execute("SELECT MAX(ITEM_ID) FROM EXISTENCIAS")
        existenciamax = cursor.fetchone()
        eresultado = existenciamax[0] if existenciamax and existenciamax[0] is not None else 999

        # Iniciar el bucle desde el valor actual de eresultado + 1 hasta invItemc
        while eresultado < invItemc:
            eresultado += 1
            cursor.execute("""
                INSERT INTO EXISTENCIAS (
                    item_ID,
                    LOCALIDAD,
                    CANTIDAD,
                    CANTIDAD_CONTADA,
                    CANTIDAD_ANTERIOR,
                    CANTIDAD_DIFERENCIA
                )
                SELECT 
                    CODIGO_ITEM, 
                    'ALM1', 
                    0, 
                    0, 
                    0, 
                    0 
                FROM INVENTARIO 
                WHERE CODIGO_ITEM = ?
            """, (eresultado,))
            conexiondb.commit()
            
    except Exception as e:
        print(f"Error: {e}")

    try:
        cursor = conexiondb.cursor()

        # Obtener los registros de la tabla INVENTARIO donde CODIGO_ITEM > 1
        cursor.execute("""
            SELECT ID, PROVEEDOR_ID 
            FROM INVENTARIO 
            WHERE CODIGO_ITEM > 1
        """)
        inventario_rows = cursor.fetchall()

        for row in inventario_rows:
            inventario_id = row[0]
            proveedor_id = row[1]

            if proveedor_id:
                # Construir el valor a comparar fuera de la consulta
                proveedor_id_prefix = proveedor_id[:4]  # Los primeros 4 caracteres de PROVEEDOR_ID
                
                # Obtener el ID del proveedor que coincide
                cursor.execute("""
                    SELECT FIRST 1 p.ID
                    FROM PROVEEDOR p
                    WHERE LEFT(p.EMPRESA, 4) = ?
                """, (proveedor_id_prefix,))
                result = cursor.fetchone()
                
                if result:
                    nuevo_proveedor_id = result[0]

                    # Actualizar el registro en la tabla INVENTARIO
                    cursor.execute("""
                        UPDATE INVENTARIO
                        SET PROVEEDOR_ID = ?
                        WHERE ID = ?
                    """, (nuevo_proveedor_id, inventario_id))

        conexiondb.commit()
        cursor.close()
    except Exception as e:
        print(f"Error al procesar proveedores: {e}")
        registrar_error(f"error en la funcion existencia {e}")
    finally:
        print("Proceso de actualización de proveedores completado")

def secuenciaInventarioUpdate():
    # Aquí deberías establecer la conexión a la base de datos
    try:
        Cursor = conexiondb.cursor()
        sequence_name = 'SEC_INVENTARIO'
        
        # Obtener el máximo ID de la tabla INVENTARIO
        Cursor.execute("SELECT MAX(ID) FROM INVENTARIO")
        idmaximo = Cursor.fetchone()[0]
        # Incrementar el valor máximo en 1 para la nueva secuencia
        newvalor = idmaximo + 1
        # Actualizar el valor del generador de secuencia
        Cursor.execute(f"SET GENERATOR {sequence_name} TO {newvalor}")
        # Confirmar la transacción si es necesario
        conexiondb.commit()
        print("Secuencia actualizada")
    except Exception as e:
        registrar_error(f"error en la funcion Secuencia Inventario {e}")
        if conexiondb:
            conexiondb.rollback()  # Revertir cambios en caso de error
        print(f"Error al actualizar la secuencia: {e}")

def secuenciasUpdate():
    # Aquí deberías establecer la conexión a la base de datos
    try:
        Cursor = conexiondb.cursor()
        sequence_name = 'SECUENCIAS'
        
        # Obtener el máximo ID de la tabla INVENTARIO
        Cursor.execute("SELECT MAX(ID) FROM INVENTARIO")
        idmaximo = Cursor.fetchone()[0]
        # Incrementar el valor máximo en 1 para la nueva secuencia
        newvalor = idmaximo + 1
        # Actualizar el valor del generador de secuencia
        Cursor.execute(f"UPDATE {sequence_name} set INV={newvalor}")
        # Confirmar la transacción si es necesario
        conexiondb.commit()
        print("Secuencia actualizada")
    except Exception as e:
        registrar_error(e)
        if conexiondb:
            conexiondb.rollback()  # Revertir cambios en caso de error
        print(f"Error al actualizar la secuencia: {e}")
        registrar_error(f"error en la funcion secuenciaUpdate {e}")

def ondetail():
    try:
        sequence_name = 'INVENTARIO_ONDETAIL_GEN'
        cursor=conexiondb.cursor()
        cursor.execute(f"SELECT GEN_ID({sequence_name}, 0) FROM RDB$DATABASE")
        ondetail=cursor.fetchone()[0]
        newvalor = ondetail + 1
        cursor.execute(f"SET GENERATOR {sequence_name} TO {newvalor}")
        conexiondb.commit()
        return newvalor
        
    except Exception as e:
        registrar_error(e)
        if conexiondb:
            conexiondb.rollback()  # Revertir cambios en caso de error
        print(f"Error al actualizar la secuencia: {e}") 

def TRAKINGNCF():
    try:
        cursor = conexiondb.cursor()
        cursor.execute("SELECT DISTINCT(NOMBRE) FROM TRAKINGNCF")
        resultado = cursor.fetchall()
        cursor.close()
        return [uni[0] for uni in resultado]
    except Exception as e:
        registrar_error(e)

def Condiciones():
    try:
        Cursor = conexiondb.cursor()
        Cursor.execute(f"SELECT DISTINCT codiciones FROM CONDICIONES_A_PAGAR")
        resultado=Cursor.fetchall()
        Cursor.close()
        return [uni[0] for uni in resultado]
    except Exception as e:
        registrar_error(e)

def Moneda():
    try:
        cursor = conexiondb.cursor()
        cursor.execute("SELECT DISTINCT(NOMBRE_MONEDA) FROM MONEDA")
        resultado = cursor.fetchall()
        cursor.close()
        return [uni[0] for uni in resultado]
    except Exception as e:
        registrar_error(e)

def Proveedormayus():
    try:
        cursor = conexiondb.cursor()
        cursor.execute("update proveedor set EMPRESA=UPPER(EMPRESA)")
        conexiondb.commit()
    except Exception as e:
        registrar_error(e)

def cantidadcliente():
    try:
        cursor=conexiondb.cursor()
        cursor.execute("SELECT COUNT(*) FROM CLIENTES")
        resultado=cursor.fetchall()
        cursor.close()
        return str(resultado[0][0])
    except Exception as e:
        registrar_error(e)
        



