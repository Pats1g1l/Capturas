import sys
import os
import logging
import mysql.connector
import time
from datetime import datetime
import pygetwindow as gw
import pyautogui





vDirApp=os.path.dirname(os.path.abspath(__file__))
vEstaApp=os.path.basename(__file__)
vArchLog = os.path.splitext(vEstaApp)[0] + ".log"
vArchConfig = os.path.splitext(vEstaApp)[0] + ".ini"

# Configuración de la conexión a la base de datos
configuracion = {
    'user': 'vigiwebc_prueba',
    'password': 'vigi.1972',
    'host': '186.64.119.90',
    'database': 'vigiwebc_sistemas',
    'port': 3306
}

#Armar archivo de log
logging.basicConfig(filename=vArchLog, encoding='utf-8', level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
#logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logging.info('')
logging.info('')

logging.info('')
logging.info('')
logging.info('***-----------------INICIANDO BOT DSS---------------------------***')
print('***-----------------INICIANDO BOT DSS---------------------------***')

#Asignando valores por defecto antes de leer la config.
logging.info('Asignando valores por defecto antes de leer la config.')
print('Asignando valores por defecto antes de leer la config.')

#Leyendo la config desde el archivo config.ini
logging.info('Leyendo la config desde el archivo: ' + vArchConfig)
print('Leyendo la config desde el archivo: ' + vArchConfig)
from configparser import ConfigParser
laConfig = ConfigParser()
laConfig.read(vArchConfig)
vCarpetaDestino = laConfig.get('Config','vCarpetaDestino')

vAbrirApp = laConfig.get('Config','vAbrirApp')
vEsperaApp = int(laConfig.get('Config','vEsperaApp'))
vEsperaTecla = int(laConfig.get('Config','vEsperaTecla'))
vIntervaloMiliSeg = int(laConfig.get('Config','vIntervaloMiliSeg'))
vConsultaSQL = laConfig.get('Config','vConsultaSQL')

vAreaIniX = int(laConfig.get('Config','vAreaIniX'))
vAreaIniY = int(laConfig.get('Config','vAreaIniY'))
vAreaFinX = int(laConfig.get('Config','vAreaFinX'))
vAreaFinY = int(laConfig.get('Config','vAreaFinY'))


vEsperaCargaDSS = int(laConfig.get('Config','vEsperaCargaDSS'))
vEsperaCargaVista = int(laConfig.get('Config','vEsperaCargaVista'))
vEsperaBusquedaTexto = int(laConfig.get('Config','vEsperaBusquedaTexto'))

vCampoClear_X = int(laConfig.get('Config','vCampoClear_X'))
vCampoClear_Y = int(laConfig.get('Config','vCampoClear_Y'))
vFocoEmpresa_X = int(laConfig.get('Config','vFocoEmpresa_X'))
vFocoEmpresa_Y = int(laConfig.get('Config','vFocoEmpresa_Y'))
vCierraVista_X = int(laConfig.get('Config','vCierraVista_X'))
vCierraVista_Y = int(laConfig.get('Config','vCierraVista_Y'))
vCierraPestania1_X = int(laConfig.get('Config','vCierraPestania1_X'))
vCierraPestania1_Y = int(laConfig.get('Config','vCierraPestania1_Y'))

vLogNivel = laConfig.get('Config','Lognivel')
vLogArchivo = laConfig.get('Config','LogArchivo')

print("Creando carpeta diaria.")
logging.info("Creando carpeta diaria.")
fecha_hora_actual = datetime.now().strftime("%Y%m%d%H%M%S")
fecha_actual = datetime.now().strftime("%Y%m%d")

vCarpetaDestino = vCarpetaDestino + fecha_actual
# Crear la subcarpeta con el nombre de empresa si no existe y si es que está definido en el config
if not os.path.exists(vCarpetaDestino):
    logging.info("Creando carpeta: " + vCarpetaDestino)
    os.makedirs(vCarpetaDestino)


# Encuentra la ventana del DSS por su título
try:
    window = gw.getWindowsWithTitle('DSS Client')[0]
    time.sleep(vEsperaCargaDSS)
    window.activate()
except:
    print("Error al intentar visualizar DSS.")
    logging.info("Error al intentar visualizar DSS.")
    print("Cerrando forzadamente por error al visualizar DSS.")
    logging.info("Cerrando forzadamente por error al visualizar DSS.")
    sys.exit()




try:
    #Activa la ventana del DSS Dahua
    print("Activando ventana de APP DSS Client.")
    logging.info("Activando ventana de APP DSS Client.")
    window.activate()
    time.sleep(vEsperaCargaDSS)
    print("Conectando a la BD.")
    logging.info("Conectando a la BD.")
    conexion = mysql.connector.connect(**configuracion)

    if conexion.is_connected():
        print("Conectando a la BD, OK.")
        logging.info("Conectando a la BD, OK.")
        # Crear un cursor
        cursor = conexion.cursor(buffered=True)

        # Ejecutar la consulta
        consulta_SQL = vConsultaSQL
        try:
            print("Ejecutando consulta a la BD: " + consulta_SQL)
            logging.info("Ejecutando consulta a la BD: " + consulta_SQL)
            cursor.execute(consulta_SQL)
        except:
            print("Error al ejecutar consulta en la BD.")
            logging.info("Error al ejecutar consulta en la BD.")
            print("Cerrando forzadamente debido al error anterior.")
            logging.info("Cerrando forzadamente debido al error anterior.")
            sys.exit()

        num_resultados = cursor.rowcount
        print("ACEMED SPATotal empresas: " + str(num_resultados))
        logging.info("Total empresas: " + str(num_resultados))

        print("-Listado de empresas a capturar")
        logging.info("inicio del listado de empresas a capturar")
        #mostrar los resultados y guardarlos al archivo de log
        for (id_empresa, nombre_empresa, EstadoCamNombre, UltEstadoCamDia, rut_empresa) in cursor:
            print(f"--ID Empresa: {id_empresa}, Nombre: {nombre_empresa}, EstadoCamNombre: {EstadoCamNombre}, RUT: {rut_empresa}")
            logging.info(f"--ID Empresa: {id_empresa}, Nombre: {nombre_empresa}, EstadoCamNombre: {EstadoCamNombre}, RUT: {rut_empresa}")
        
        print("Total empresas: " + str(num_resultados))
        logging.info("Total empresas: " + str(num_resultados))

        cursor.execute(consulta_SQL)
        # Recorrer los resultados y generar la captura de pantalla
        for (id_empresa, nombre_empresa, EstadoCamNombre, UltEstadoCamDia, rut_empresa) in cursor:
            window.activate()
            time.sleep(vEsperaCargaDSS)


            print("Preparando pantallazo [ID Empresa: {id_empresa}, Nombre Empresa: {EstadoCamNombre} ] ultima fecha procesada: " + str(UltEstadoCamDia))
            logging.info("Preparando pantallazo [ID Empresa: {id_empresa}, Nombre Empresa: {EstadoCamNombre} ] ultima fecha procesada: " + str(UltEstadoCamDia))

            
            #LIMPIAR TEXTO DE BUSQUEDA
            print("Limpiar campo de busqueda: Se hara clic en " + str(vCampoClear_X) + "," + str(vCampoClear_Y))
            logging.info("Limpiar campo de busqueda: Se hara clic en " + str(vCampoClear_X) + "," + str(vCampoClear_Y))
            pyautogui.click(x=vCampoClear_X, y=vCampoClear_Y)
            time.sleep(vEsperaBusquedaTexto)

            #HACER CLIC EN EL CAMPO DE BUSQUEDA Y ESCRIBIR LA EMPRESA
            print("Escribir Nombre Empresa [" + EstadoCamNombre + "]")
            logging.info("Escribir Nombre Empresa [" + EstadoCamNombre + "]")
            pyautogui.click(x=vCampoClear_X, y=vCampoClear_Y)
            time.sleep(vEsperaBusquedaTexto)
            pyautogui.write(EstadoCamNombre)
            pyautogui.press('enter')
            time.sleep(vEsperaBusquedaTexto)

            #HACER DOBLE CLIC EN LA EMPRESA ENCONTRADA
            print("Doble clic en la empresa encontrada, en las coordenadas del primer resultado: " + str(vFocoEmpresa_X) + "," + str(vFocoEmpresa_Y))
            logging.info("Doble clic en la empresa encontrada, en las coordenadas del primer resultado: " + str(vFocoEmpresa_X) + "," + str(vFocoEmpresa_Y))
            pyautogui.doubleClick(x=vFocoEmpresa_X, y=vFocoEmpresa_Y)
            time.sleep(vEsperaCargaVista)
            #PRINT DE PANTALLA
            nombre_archivo = f"{EstadoCamNombre}_{fecha_actual}.jpg"
            print("El nombre de archivo sera: " + nombre_archivo)
            logging.info("El nombre de archivo sera: " + nombre_archivo)            
            
            print("Haciendo el print de pantalla en las areas: (" + str(vAreaIniY) + "," + str(vAreaIniY) + " y " + str(vAreaFinX) + "," + str(vAreaFinY)+ ")")         
            logging.info("Haciendo el print de pantalla en las areas: (" + str(vAreaIniY) + "," + str(vAreaIniY) + " y " + str(vAreaFinX) + "," + str(vAreaFinY)+ ")")         
            captura = pyautogui.screenshot(region=(vAreaIniX, vAreaIniY, vAreaFinX, vAreaFinY))

            vArchivoFinal= vCarpetaDestino + "\\" + nombre_archivo

            print("Guardando archivo: " + vArchivoFinal)
            logging.info("Guardando archivo: " + vArchivoFinal)            

            captura.save(vArchivoFinal)
            print("Guardado OK.")
            logging.info("Guardado OK.")

            print("Cerrando vista cliente. Se hara clic en: " + str(vCierraVista_X) + "," + str(vCierraVista_Y))
            logging.info("Cerrando vista cliente. Se hara clic en: " + str(vCierraVista_X) + "," + str(vCierraVista_Y))   
            pyautogui.click(x=vCierraVista_X, y=vCierraVista_Y)     

            #Actualizar la empresa.UltEstadoCamDia poniendo la fecha de hoy, para que se sepa que se proceso ya hoy.
            cursor2 = conexion.cursor()
            update_query = """UPDATE empresa SET UltEstadoCamDia = NOW(), estadocamdia = 2 WHERE id_empresa = %s"""
            cursor2.execute(update_query, (id_empresa,))
            conexion.commit()
            print("Fecha actualizada correctamente")
            cursor2.close()
        # Volver al inicio de los resultados
        cursor.close()
        cursor = conexion.cursor()

except mysql.connector.Error as error:
    print("Error al conectar a la base de datos:", error)
    logging.info("Error al conectar a la base de datos:", error)

finally:
    # Cerrar la conexión
    if 'conexion' in locals() and conexion.is_connected():
        conexion.close()
        print("Conexion BD cerrada.")
        logging.info("Conexion BD cerrada.")



print("Cerrando la ultima vista cargada. Se hara clic en: " + str(vCierraVista_X) + "," + str(vCierraVista_Y))
logging.info("Cerrando la ultima vista cargada. Se hara clic en: " + str(vCierraVista_X) + "," + str(vCierraVista_Y))   
pyautogui.click(x=vCierraVista_X, y=vCierraVista_Y)

#LIMPIAR TEXTO DE BUSQUEDA
print("Limpiar campo de busqueda: Se hara clic en " + str(vCampoClear_X) + "," + str(vCampoClear_Y))
logging.info("Limpiar campo de busqueda: Se hara clic en " + str(vCampoClear_X) + "," + str(vCampoClear_Y))
pyautogui.click(x=vCampoClear_X, y=vCampoClear_Y)

#CERRAR VENTANA DE MONITOREO
print("Cerrar ventana monitoreo: Se hara clic en " + str(vCierraPestania1_X) + "," + str(vCierraPestania1_Y))
logging.info("Cerrar ventana monitoreo: Se hara clic en " + str(vCierraPestania1_X) + "," + str(vCierraPestania1_Y))
pyautogui.click(x=vCierraPestania1_X, y=vCierraPestania1_Y)

logging.info('***-----------------FINALIZANDO BOT DSS---------------------------***')
print('***-----------------FINALIZANDO BOT DSS---------------------------***')


