from email import encoders
from email.mime.base import MIMEBase
from dotenv import dotenv_values
import smtplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

config = {
    **dotenv_values(".env"),  # prod .env  o develop .env
    **dotenv_values(".env.secret"),  #
}


def envio_email(asunto, cuerpo):
    try:
        remitente = 'felipe.alvarez@ext.enex.cl'
        destinatarios = json.loads('felipe.alvarez@ext.enex.cl','gonzalo.arbulo@enex.cl','dionisio.cataldo@enex.cl')
        mensaje = MIMEMultipart()
        mensaje['From'] = remitente
        mensaje['To'] = ", ".join(destinatarios)
        mensaje['Subject'] = asunto
        mensaje.attach(MIMEText(cuerpo, 'plain'))
        sesion_smtp = smtplib.SMTP(host=config['EMAIL_HOST'], port=config['EMAIL_PORT'])
        sesion_smtp.starttls()
        sesion_smtp.login(config['EMAIL_USUARIO'], config['EMAIL_SECRET'])
        texto = mensaje.as_string()
        sesion_smtp.sendmail(remitente, destinatarios, texto)
        sesion_smtp.quit()
    except:
        print('No se pudo enviar el correo')


def send_error_csv(cliente, codigo_jde, rut):
    message = "Error en el correo con el Cliente: " + cliente + ", Codigo JDE: " + codigo_jde + ", RUT: " + rut + "\n"
    try:
        remitente = config['EMAIL_USUARIO']
        destinatarios = json.loads(config['EMAIL_DESTINATARIOS_LIST'])
        asunto = '[CLUB] Error in mail'
        cuerpo = message
        mensaje = MIMEMultipart()
        mensaje['From'] = remitente
        mensaje['To'] = ", ".join(destinatarios)
        mensaje['Subject'] = asunto
        mensaje.attach(MIMEText(cuerpo, 'plain'))

        with open(config['ERROR_EXCEL_FILE'], "rb") as adjunto:
            parte = MIMEBase('application', 'octet-stream')
            parte.set_payload(adjunto.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f"attachment; filename= {config['ERROR_EXCEL_FILE']}")
        mensaje.attach(parte)

        sesion_smtp = smtplib.SMTP(host=config['EMAIL_HOST'], port=config['EMAIL_PORT'])
        sesion_smtp.starttls()
        sesion_smtp.login(config['EMAIL_USUARIO'], config['EMAIL_SECRET'])
        texto = mensaje.as_string()
        sesion_smtp.sendmail(remitente, destinatarios, texto)
        sesion_smtp.quit()
    except:
        print('No se pudo enviar el correo')



