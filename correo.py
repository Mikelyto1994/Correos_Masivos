import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email():
    try:
        
        excel = pd.read_excel("Libro1.xlsx")
        conexion = smtplib.SMTP('smtp.gmail.com', 587)
        conexion.starttls()  
        conexion.login(user='ejemplo@gmail.com', password='ivat myyv zjmj hrml')
        emails = excel["EMAILS"]
        
        with open("mensaje.html", "r", encoding='utf-8') as file:
            mensaje_html = file.read()
        
        for email in emails:
            msg = MIMEMultipart()
            msg['From'] = 'ejemplo@gmail.com' 
            msg['To'] = email 
            msg['Subject'] = 'CORREOS MASIVOS PUBLICIDAD PRUEBA'
            
            msg.attach(MIMEText(mensaje_html, 'html'))
            
            conexion.sendmail(from_addr='ejemplo@gmail.com', to_addrs=email, msg=msg.as_string())

        conexion.quit()
        print("Correos enviados correctamente")
    except Exception as e:
        print(f"Error al enviar los correos: {e}")
send_email()

