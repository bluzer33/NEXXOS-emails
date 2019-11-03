import smtplib
import pandas as pd
import xlrd
import os
import re
import PySimpleGUI as sg
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate

# Logeo
# Obj = smtplib.SMTP('smtp.gmail.com', 587)
# Obj.starttls()
# print(Obj.ehlo())  #250 = el hello anda
# print(Obj.login('abbatelucas', 'msyqgosadudbkbsf'))

def send_mail(send_from, send_to, subject, text, files=None, html=False): #server="127.0.0.1"):
	#assert isinstance(send_to, list)

	msg = MIMEMultipart()
	msg['From'] = send_from
	msg['To'] = send_to #COMMASPACE.join(send_to) Este commaspace sirve si queres mandar a todos el mismo mail (x ejemplo para el newsletter)
	msg['Date'] = formatdate(localtime=True)
	msg['Subject'] = subject
	msg.add_header('Content-Type','text/html')

	if html == False:
		msg.attach(MIMEText(text, 'plain'))
	elif html == True:
		msg.attach(MIMEText(text, 'html'))
	else:
		raise TypeError('HTML tiene que ser True o False')

	for f in files or []:
		with open(f, "rb") as fil:
			part = MIMEApplication(
				fil.read(),
				Name=basename(f)
				)
        # After the file is closed
		part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
		msg.attach(part)

	Obj.sendmail(send_from, send_to, msg.as_string())
	

def get_contacts(xlsx='Archivos Adjuntos.xlsx'):
	contacts = []
	wb = xlrd.open_workbook(xlsx)
	sheet = wb.sheet_by_index(0)
	for line in range(1, sheet.nrows):
		contact = sheet.cell_value(line, 2)
		contacts.append(str(contact))
	return contacts

def get_files(xlsx='Archivos Adjuntos.xlsx'):
	files = []
	wb = xlrd.open_workbook(xlsx)
	sheet = wb.sheet_by_index(0)
	for line in range(1, sheet.nrows):
		vector = []
		for col in range (3, sheet.ncols):
				f = str(sheet.cell_value(line, col)) + '.png'
				if '-d' in f:
					f = 'Diplomas/' + f
				elif '-f' in f:
					f = 'Facturas/' + f
				vector.append(f)
		files.append(vector)
	return files

def get_names(xlsx='Archivos Adjuntos.xlsx'):
	names = []
	wb = xlrd.open_workbook(xlsx)
	sheet = wb.sheet_by_index(0)
	for line in range(1, sheet.nrows):
		surname = sheet.cell_value(line, 0)
		name = sheet.cell_value(line, 1)
		names.append(str(name + ' ' + surname))
	return names


def Enviar_mails(asunto, body,
html_or_plain='plain', xlsx='Archivos Adjuntos.xlsx'):#diff_ mails para desarrollar dssp (con el commaspace)

	if html_or_plain == 'plain':
		html = False
	elif html_or_plain.lower() == 'html':
		html = True
	else:
		html = False

	for adress, files, name in zip(get_contacts(xlsx), get_files(xlsx), get_names(xlsx)):
		send_mail('abbatelucas@gmail.com', adress, asunto, body.format(name = name), files, html)
		print('Mensaje enviado a: '+ adress)
	

login_layout = [
			[sg.T('Mail (sin el "@gmail.com")')],
			[sg.In(key='USERNAME', size=(30, 1))],
			[sg.T('Contraseña de acceso remoto')],
			[sg.In(key='PASS', size=(30, 1))],
			[sg.Button('Log In')],
			[sg.Output(size=(30, 2))]
]





login_window = sg.Window('Log In', login_layout)

while True:      
	(event, value) = login_window.Read()      
	if event is None:      
		break # exit button clicked      
	if event == 'Log In':      
		Obj = smtplib.SMTP('smtp.gmail.com', 587)
		Obj.starttls()
		Obj.ehlo()  #250 = el hello anda
		log_status = Obj.login(value['USERNAME'], value['PASS'])
		print(log_status)
		if log_status == (235, b'2.7.0 Accepted'):
			sg.Popup('Logeo correctamente, tocá el botón para continuar')
			#login_window['-LOGIN-'].Update('Continuar')
			#if event is 'Continuar':
			login_window.Close()
			layout = [
				[sg.T('Parámetros', size=(20, 1))],
				[sg.In('Nombre', key='NAME')],
				[sg.Button('Correos'), sg.Button('Nombres')], 
				[sg.Text('Script output....', size=(20, 1))],      
				[sg.Output(size=(20, 20))]
				#[sg.Text('Manual command', size=(15, 1)), sg.InputText(focus=True), sg.Button('Run', bind_return_key=True)]
				]
			window = sg.Window('Mail Sender', layout)
			while True:      
				(event, value) = window.Read()      
				if event is None:      
					break # exit button clicked      
				if event == 'Correos':      
					print(get_contacts()) 
				elif event == 'Nombres':      
					print(get_names()) 



# Show the Window to the user


# Event loop. Read buttons, make callbacks

