import sys
import os
import logging
import mysql.connector
import win32com.client as win32
import time
from datetime import datetime
import shutil
import re

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

#Leyendo la config desde el archivo config.ini
logging.info('Leyendo la config desde el archivo: ' + vArchConfig)
print('Leyendo la config desde el archivo: ' + vArchConfig)
from configparser import ConfigParser
laConfig = ConfigParser()
laConfig.read(vArchConfig)
vCarpetaDestino = laConfig.get('Config','vCarpetaDestino')
vConsultaSQL = laConfig.get('Config','vConsultaSQL')
vCorreosInternos = laConfig.get('Config','vCorreosInternos')

vLogNivel = laConfig.get('Config','Lognivel')
vLogArchivo = laConfig.get('Config','LogArchivo')


outlook = win32.Dispatch('outlook.application')

try:
    # Conectar a la base de datos
    conexion = mysql.connector.connect(**configuracion)

    if conexion.is_connected():
        print("Conectando a la BD, OK.")
        logging.info("Conectando a la BD, OK.")
        # Crear un cursor
        cursor = conexion.cursor(buffered=True)

        # Consulta SQL para obtener los campos solicitados de la tabla empresa
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
        print("Total empresas: " + str(num_resultados))
        logging.info("Total empresas: " + str(num_resultados))

       
        # Recorrer los resultados y mostrarlos
        for (id_empresa, nombre_empresa, EstadoCamNombre, email_empresa, seguridad_correo, tecnico_correo, emergencia1_email, emergencia2_email, emergencia3_email,emergencia4_email,emergencia5_email,emergencia6_email, empresa_guardia, EstadoCamDia, EstadoCamNoc, UltEstadoCamDia, UltEstadoCamnoc) in cursor:

            vAhora = datetime.now()
            vFechaHoy = str(vAhora.strftime("%d/%m/%Y"))
            vFechaHora = str(vAhora.strftime("%d/%m/%Y %H:%M:%S"))
            vFechaHora2 = str(vAhora.strftime("%Y%m%d%H%M%S"))

            nombre_imagen = EstadoCamNombre + "_"+ str(vAhora.strftime("%Y%m%d")) + ".jpg"
            ruta_imagen = vCarpetaDestino + str(vAhora.strftime("%Y%m%d")) + "\\"+ nombre_imagen

            print(f"Preparando estado de cámaras: {EstadoCamNombre}")
            print(f"Mombre de archivo imagen: {nombre_imagen}")
            print(f"Ruta de archivo imagen: {ruta_imagen}")
            logging.info(f"---Preparando estado de cámaras: {EstadoCamNombre}")
            logging.info(f"Mombre de archivo imagen: {nombre_imagen}")
            logging.info(f"Ruta de archivo imagen: {ruta_imagen}")

            #crear lista de correos leidos desde la BD
            print(f"Leyendo lista de destinatarios, eliminando repetidos y limpiando casillas mal escritas.")
            logging.info(f"Leyendo lista de destinatarios, eliminando repetidos y limpiando casillas mal escritas.")
            lista_correos = []
            lista_correos.append(email_empresa)
            lista_correos.append(seguridad_correo)
            lista_correos.append(tecnico_correo)
            lista_correos.append(emergencia1_email)
            lista_correos.append(emergencia2_email)
            lista_correos.append(emergencia3_email)
            lista_correos.append(emergencia4_email)
            lista_correos.append(emergencia5_email)
            lista_correos.append(emergencia6_email)
            # Mostrar la lista antes de eliminar duplicados
            print("Lista antes de eliminar duplicados:", lista_correos)
            # Eliminar duplicados usando un conjunto y volver a convertirlo en lista
            lista_sin_duplicados = list(set(lista_correos))
            # Mostrar la lista después de eliminar duplicados
            print("Lista después de eliminar duplicados:", lista_sin_duplicados)
            # Definir una expresión regular para un correo electrónico
            regex_email = re.compile(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")

            # Filtrar la lista para eliminar elementos que no tengan un formato de correo electrónico válido
            lista_validos = [email for email in lista_sin_duplicados if regex_email.match(email)]
            lista_destinatarios = "; ".join(lista_validos)


             # Crear el mensaje
            correo = outlook.CreateItem(0)  # 0 indica un nuevo correo (olMailItem)
            correo.Subject = f"Estado de cámaras: {EstadoCamNombre} ({vFechaHoy})"
            correo.To = lista_destinatarios
            correo.CC = vCorreosInternos
            #cuerpo_correo = f"Estimados {nombre_empresa},\n\nEl siguiente informe detalla el estado de Cámaras y Sistema de Audio."

            print(f"Lista de destinatarios: {lista_destinatarios}")
            logging.info(f"Lista de destinatarios: {lista_destinatarios}")

            # Definir el HTML del correo, incluyendo un marcador para la imagen
            html_body = """
            <!DOCTYPE html>
            <html>
            <head>
                <title>Estado de camaras</title>
            </head>
            <body>
                <h1>Estado de camaras de seguridad</h1>
                <p>Estimados {}, Vigilancia Web informa el estado de Camaras y Sistema de Audio.</p>
                <p></p>
                <p>Operativos sin novedad.</p>
                <p>Fecha y Hora: {}</p>

                <!-- LugarImagen -->  <!-- Este es el marcador donde se insertará la imagen -->
                <p>Este mensaje es para fines informativos, se han omitido automaticamente los acentos.</p>
                <p>Atentamente:</p>
                <p>Vigilancia Web</p>
            </body>
            </html>
            """




            # Incrustar la imagen en el cuerpo del correo
            if os.path.exists(ruta_imagen):
                print(f"Incrustando imagen en el correo.")
                logging.info(f"Incrustando imagen en el correo.") 
                # Añadir la imagen como un adjunto embebido
                attachment = correo.Attachments.Add(ruta_imagen, 1, 0)
                cid = "imagen001"
                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", cid)
                
                # Insertar la etiqueta <img> en el lugar deseado dentro del HTML
                etiqueta_img = f'<img src="cid:{cid}">'
                html_body = html_body.replace("<!-- LugarImagen -->", etiqueta_img)
            else:
                print(f"No se encontro imagen a incrustar.")
                logging.info(f"No se encontro imagen a incrustar.")

            correo.HTMLBody = html_body.format(EstadoCamNombre, vFechaHora)
            
            # Guardar el correo como borrador
            correo.Save()
            print(f"Guardando el borrador.")
            logging.info(f"Guardando el borrador.")

            #Mover la imagen porque ya fue enviada
            print(f"Moviendo imagen ya procesada, a la carpeta de enviados.")
            logging.info(f"Moviendo imagen ya procesada, a la carpeta de enviados.")
            vcarpeta_Enviados = vCarpetaDestino + str(vAhora.strftime("%Y%m%d")) + "\\Enviadas"
            # Crear la subcarpeta si no existe, para los archivos de imagenes enviadas
            if not os.path.exists(vcarpeta_Enviados):
                os.makedirs(vcarpeta_Enviados)
            
            vArchivoDestino = vcarpeta_Enviados + "\\"+nombre_imagen
           
            # Mover el archivo a la subcarpeta
            try:
                if os.path.exists(ruta_imagen) and os.path.exists(vArchivoDestino):
                    os.remove(vArchivoDestino)
                    print(f"Archivo anterior existente eliminado: {vArchivoDestino}")
                    logging.info(f"Archivo anterior existente eliminado: {vArchivoDestino}")
                shutil.move(ruta_imagen, vArchivoDestino)
                print(f"Archivo movido a: {vArchivoDestino}")
                logging.info(f"Archivo movido a: {vArchivoDestino}")
            except:
                print(f"No se ha podido mover el archivo: {ruta_imagen}.")
                logging.info(f"No se ha podido mover el archivo: {ruta_imagen}.")


        # Volver al inicio de los resultados
        cursor.close()
        cursor = conexion.cursor()

except mysql.connector.Error as error:
    print("Error al conectar a la base de datos:", error)

finally:
    # Cerrar la conexión
    if 'conexion' in locals() and conexion.is_connected():
        conexion.close()
        print("Conexión cerrada.")


logging.info('***-----------------FINALIZANDO BOT DSS---------------------------***')
print('***-----------------FINALIZANDO BOT DSS---------------------------***')
