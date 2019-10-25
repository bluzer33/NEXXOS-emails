import smtplib
import email.message

smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
smtpObj.starttls()

# Verificación y logeo
print(smtpObj.ehlo())  #250 = el hello anda
print(smtpObj.login('abbatelucas', 'msyqgosadudbkbsf'))  #235 = logeo bien

# Definición del Asunto y del Correo
asunto = 'Probando gmail desde python.'
