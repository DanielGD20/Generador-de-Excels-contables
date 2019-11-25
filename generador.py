
# Aqui se importan las librerias para el manejo de Excels
import xlsxwriter
# Aqui se importan las librerias para el manejo de emails
import smtplib
from email.message import EmailMessage
# Aqui se importan las librerias para el manejo de bases de datos en csv
import pandas
import numpy as np

cuerpo_mail = """
"""


def busquedaEnCsv(fraseABuscar, dondeBuscar):
    # Importante
    # Row[6] = Numero de factura
    # Row[8] = Codigo de cliente
    # Row[9] = Nombre del cliente
    # Row[10] = Fecha de emision
    # Row[12] = Saldo de factura
    # Row[16] = Neto de factura

    guardadoDeDatosPorFactura = np.array([])
    for row in dondeBuscar.itertuples():
        guardadorFacturita = np.array([])
        if fraseABuscar == row[9]:
            guardadorFacturita = np.append(guardadorFacturita, row[6])
            guardadorFacturita = np.append(guardadorFacturita, row[8])
            guardadorFacturita = np.append(guardadorFacturita, row[9])
            guardadorFacturita = np.append(guardadorFacturita, row[10])
            guardadorFacturita = np.append(guardadorFacturita, row[12])
            guardadorFacturita = np.append(guardadorFacturita, row[16])
            print()
            print("---------------------------------------------------------")
            print(guardadorFacturita)
            guardadoDeDatosPorFactura = np.append(
                guardadoDeDatosPorFactura, guardadorFacturita, 0)
    print("--------------------------------Arreglo Final----------------------")
    print(guardadoDeDatosPorFactura)
    return guardadoDeDatosPorFactura


def main():
    # ------------------------------Se obtiene el documento csv y se lo lee---------------------------------

    print("leyendo base de datos de clientes...")

    base_de_datos_clientes = pandas.read_csv(
        './Excels/clientes.csv', sep=';', engine='python')

    print("leyendo base de datos de cartera...")

    base_de_datos_cartera = pandas.read_csv(
        './Excels/cartera.csv', sep=';', engine='python')

    for row in base_de_datos_clientes.head(10).itertuples():
        print()
        print("---------------------------------------------------------")
        print(row[1])
        print(row[3])
        print(row[17])
        # Importante
        # Row[1] = Codigo de cliente
        # Row[3] = Nombres del cliente
        # Row[17] = Email del cliente
        resultadoDeBusqueda = busquedaEnCsv(row[3], base_de_datos_cartera)

        # for row in base_de_datos.head(1).itertuples():
        #     # Importante
        #     # Row[2] = correos
        #     # Row[3] = nombres
        #     # Row[4] = apellidos
        #     # Row[5] = titulo del ensayo
        #     # Row[6] = nombre del tutor
        #     # Row[7] = Documento enviado
        #     # Row[8] = posee tutor?

        #     print()
        #     print("---------------------------------------------------------")
        #     print(row)
        #     print("Generando diploma numero: " + str(contadorDeDiplomas))
        #     # ------------------------------------------------------------------------------------------------------
        #     # ------------------------------Se obtiene la imagen para procesarla con el nuevo texto-----------------

        #     if(isinstance(row[3], str) or isinstance(row[4], str)):

        #          # ------------------------------Pregunta si tiene tutor o no -------------------------------------------
        #         if(row[8] == "si"):

        #             cambioDeNombres = 0

        #             imgProfe = Image.open('diploma.jpg')
        #             # Selecciona el nombre del tutor
        #             Nombre_Tutor = ""+str(row[6])
        #             destinoTutor = "C:/Users/User/Desktop/Generador de pdfs/Diplomas/Tutores/diploma" + \
        #                 str(contadorDeDiplomas)+".jpg"

        #             font = ImageFont.truetype("avenir.ttf", 160)
        #             w, h = font.getsize(Nombre_Tutor)

        #             draw = ImageDraw.Draw(imgProfe)
        #             x = (5100-w)/2
        #             y = (3000-h)/2
        #             draw.text((x, y), Nombre_Tutor, (84, 84, 84), font=font)

        #             imgProfe.copy()
        #             imgProfe.save(destinoTutor, 'JPEG', quality=80, optimize=True)
        #             print("Diploma de profesor generado!")
        #             print("Generando diploma de profesor...")

        #             # ------------------------------------------------------------------------------------------------------
        #             imgAlumno = Image.open('diploma.jpg')
        #             Nombre_Persona = ""+str(row[3])+" "+str(row[4])
        #             destinoAlumno = "C:/Users/User/Desktop/Generador de pdfs/Diplomas/Alumnos/diploma" + \
        #                 str(contadorDeDiplomas)+".jpg"

        #             font = ImageFont.truetype("avenir.ttf", 160)
        #             w, h = font.getsize(Nombre_Persona)

        #             draw = ImageDraw.Draw(imgAlumno)
        #             x = (5100-w)/2
        #             y = (3000-h)/2
        #             draw.text((x, y), Nombre_Persona, (84, 84, 84), font=font)

        #             imgAlumno.copy()
        #             imgAlumno.save(destinoAlumno, 'JPEG', quality=80, optimize=True)
        #             print("Diploma de alumno generado!")

        #             # ------------------------------Se Crea el nuevo email con el que sera guardada la fotografia-----------

        #             print("Generando el correo electrónico...")
        #             EMAIL_ADDRES = "advillegas@uees.edu.ec"
        #             EMAIL_PASSWORD = "094504722a"

        #             contact = row[2]

        #             msg = EmailMessage()
        #             msg['subject'] = 'Test'
        #             msg['from'] = EMAIL_ADDRES
        #             msg['to'] = contact

        #             msg.set_content('Certificado de participación')
        #             msg.add_alternative(cuerpo_mail, subtype='html')

        #             diplomas = [destinoAlumno, destinoTutor]

        #             for diploma in diplomas:
        #                 with open(diploma, 'rb') as f:
        #                     file_data = f.read()
        #                     file_type = imghdr.what(f.name)
        #                     if(cambioDeNombres == 0):
        #                         file_name = "Diploma de Participación Alumno"
        #                     else:
        #                         file_name = "Diploma de Participación Tutor"
        #                     cambioDeNombres += 1

        #                 msg.add_attachment(file_data, maintype='image',
        #                                    subtype=file_type, filename=file_name)

        #             with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        #                 smtp.login(EMAIL_ADDRES, EMAIL_PASSWORD)
        #                 smtp.send_message(msg)

        #             # ------------------------------------------------------------------------------------------------------
        #             contadorDeDiplomas += 1
        #             print("Correo enviado correctamente!")
        #             print("---------------------------------------------------------")
        #             print("Siguiente...")

        #         else:

        #             imgAlumno = Image.open('diploma.jpg')
        #             Nombre_Persona = ""+str(row[3])+" "+str(row[4])
        #             destinoAlumno = "C:/Users/User/Desktop/Generador de pdfs/Diplomas/Alumnos/diploma" + \
        #                 str(contadorDeDiplomas)+".jpg"

        #             font = ImageFont.truetype("avenir.ttf", 160)
        #             w, h = font.getsize(Nombre_Persona)

        #             draw = ImageDraw.Draw(imgAlumno)
        #             x = (5100-w)/2
        #             y = (3000-h)/2
        #             draw.text((x, y), Nombre_Persona, (84, 84, 84), font=font)

        #             imgAlumno.copy()
        #             imgAlumno.save(destinoAlumno, 'JPEG', quality=80, optimize=True)
        #             print("Diploma de alumno generado!")

        #             # ------------------------------------------------------------------------------------------------------

        #             # ------------------------------Se Crea el nuevo email con el que sera guardada la fotografia-----------

        #             print("Generando el correo electrónico...")
        #             EMAIL_ADDRES = "advillegas@uees.edu.ec"
        #             EMAIL_PASSWORD = "094504722a"

        #             contact = row[2]

        #             msg = EmailMessage()
        #             msg['subject'] = 'Diploma de participación - Beca completa UEES'
        #             msg['from'] = EMAIL_ADDRES
        #             msg['to'] = contact

        #             msg.set_content('Certificado de participación')
        #             msg.add_alternative(cuerpo_mail, subtype='html')

        #             with open(destinoAlumno, 'rb') as f:
        #                 file_data = f.read()
        #                 file_type = imghdr.what(f.name)
        #                 file_name = "Diploma de Participación"

        #             msg.add_attachment(file_data, maintype='image',
        #                                subtype=file_type, filename=file_name)

        #             with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        #                 smtp.login(EMAIL_ADDRES, EMAIL_PASSWORD)
        #                 smtp.send_message(msg)

        #             # ------------------------------------------------------------------------------------------------------
        #             contadorDeDiplomas += 1
        #             print("Correo enviado correctamente!")
        #             print("---------------------------------------------------------")
        #             print("Siguiente...")
        #     else:
        #         print("El diploma numero: " + str(contadorDeDiplomas) +
        #               " No cuenta con sus nombres escritos correctamente")
        #         contadorDeDiplomas += 1

        # Funciones
