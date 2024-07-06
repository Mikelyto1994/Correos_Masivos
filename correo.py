import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import asyncio
import aiosmtplib

async def send_email():
    try:
        # Leer la lista de correos electrónicos desde el archivo Excel
        excel = pd.read_excel("Libro1.xlsx")
        emails = excel["EMAILS"]
        
        # Limitar el número de correos electrónicos a 500
        max_emails = 500
        emails = emails[:max_emails]

        # Leer el contenido del mensaje HTML
        with open("mensaje.html", "r", encoding='utf-8') as file:
            mensaje_html = file.read()

        # Enviar los correos electrónicos
        for email in emails:
            msg = MIMEMultipart()
            msg['From'] = 'ejemplo@gmail.com'
            msg['To'] = email
            msg['Subject'] = 'CORREOS MASIVOS PUBLICIDAD PRUEBA'

            # Personalizar el mensaje HTML
            mensaje_personalizado = mensaje_html.replace("{email}", email)
            msg.attach(MIMEText(mensaje_personalizado, 'html'))

            try:
                await aiosmtplib.send(
                    message=msg,
                    hostname="smtp.gmail.com",
                    port=587,
                    start_tls=True,
                    username="ejemplo@gmail.com",
                    password="tu_contraseña_aquí"
                )
                print(f"Correo enviado a {email}")
            except Exception as e:
                print(f"Error al enviar correo a {email}: {e}")

            # Añadir un retardo de 2 segundos entre envíos
            await asyncio.sleep(2)

        print("Todos los correos han sido enviados correctamente")
    except Exception as e:
        print(f"Error al enviar los correos: {e}")

# Ejecutar la función asincrónicamente
asyncio.run(send_email())
