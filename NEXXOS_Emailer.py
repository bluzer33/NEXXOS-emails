import smtplib
import xlrd
import PySimpleGUI as sg
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate


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
		try:
			with open(f, "rb") as fil:
				part = MIMEApplication(
					fil.read(),
					Name=basename(f)
					)
			# After the file is closed
			part['Content-Disposition'] = 'attachment; filename="%s"' % basename(f)
			msg.attach(part)
		except:
			pass
	try:
		Obj.sendmail(send_from, send_to, msg.as_string())
	except:
		pass

def get_contacts(xlsx='Archivos Adjuntos.xlsx'):
	contacts = []
	wb = xlrd.open_workbook(xlsx)
	sheet = wb.sheet_by_index(0)
	for line in range(1, sheet.nrows):
		try:
			contact = sheet.cell_value(line, 2)
			contacts.append(str(contact))
		except:
			contacts.append('')
	return contacts

def get_files(xlsx='Archivos Adjuntos.xlsx'):
	files = []
	wb = xlrd.open_workbook(xlsx)
	sheet = wb.sheet_by_index(0)
	for line in range(1, sheet.nrows):
		vector = []
		for col in range (3, sheet.ncols):
			try:
				g = str(sheet.cell_value(line, col))
				f = str(sheet.cell_value(0, col))
				h = r'{}\{}'.format(f, g)
				vector.append(h)
			except:
				vector.append('')
		files.append(vector)
	return files

def get_names(xlsx='Archivos Adjuntos.xlsx'):
	names = []
	wb = xlrd.open_workbook(xlsx)
	sheet = wb.sheet_by_index(0)
	for line in range(1, sheet.nrows):
		try:
			surname = sheet.cell_value(line, 0)
			name = sheet.cell_value(line, 1)
			names.append(str(name + ' ' + surname))
		except:
			names.append('')
	return names


def Enviar_mails(asunto, body_route, html_or_plain='plain', xlsx='Archivos Adjuntos.xlsx'):#diff_ mails para desarrollar dssp (con el commaspace)

	if html_or_plain == 'plain':
		html = False
	elif html_or_plain == 'html':
		html = True
	else:
		html = False
	f = open(r'{}'.format(body_route), 'r')
	body = f.read()
	for adress, files, name in zip(get_contacts(xlsx), get_files(xlsx), get_names(xlsx)):
		try:
			send_mail(value['USERNAME'], adress, asunto, body.format(name = name), files, html)
			print('Mensaje enviado a: '+ adress)
		except:
			print('Hubo un error, y no se envió el mail a: '+ adress + '\nFijate que el Excel esté correcto y que el mail sea valido')

login_layout = [
			[sg.T('Mail (sin el "@gmail.com")')],
			[sg.In(key='USERNAME', size=(60, 1))],
			[sg.T('Contraseña de aplicación')],
			[sg.In(key='PASS', size=(60, 1))],
			[sg.Button('Log In'), sg.Button('¿Qué contraseña?')],
			[sg.Output(size=(60, 3), key='OUTPUT')]
]


login_window = sg.Window('Log In', login_layout)

while True:      
	(event, value) = login_window.Read()
	if event is None:      
		break # exit button clicked      
	if event == '¿Qué contraseña?':
		print('Entrá a: \nhttps://support.google.com/accounts/answer/185833?hl=es\nCuando pregunte para que aplicación, pone "Otra"')
	if event == 'Log In':      
		Obj = smtplib.SMTP('smtp.gmail.com', 587)
		Obj.starttls()
		Obj.ehlo()  #250 = el hello anda
		log_status = Obj.login(value['USERNAME'], value['PASS'])
		print(log_status)
		if log_status == (235, b'2.7.0 Accepted'):
			sg.Popup('Logeo correctamente, tocá el botón para continuar')
			login_window.Close()
			layout = [
				[sg.T('Asunto:', size=(60, 1))],
				[sg.In('', key='ASUNTO')],
				[sg.Rad('Texto Plano', 'HTML-PLAIN', key='PLAIN', default=True), sg.Rad('HTML', 'HTML-PLAIN', key='HTML')],
				[sg.T('Ruta del texto/HTML (incluir el nombre del archivo y la extensión)')],
				[sg.In('', key='BODY')], 
				[sg.T('Ruta del Excel (incluir el nombre del archivo y la extensión):', size=(60, 1))],
				[sg.In('', key='XLSX')],
				[sg.B('Formato del Excel')],
				[sg.B('Enviar')],    
				[sg.Output(size=(60, 5))]
				]
			window = sg.Window('NEXXOS Bulk Mail', layout)
			while True:      
				(event_2, value_2) = window.Read()      
				if event_2 is None:      
					break # exit button clicked      
				if event_2 == 'Formato del Excel':
					print('Apellido|Nombre|Mail\nLas columnas restantes son para archivos adjuntos, tienen que tener en la primera fila la ruta de la carpeta, y en las celdas los nombres de los archivos con las extensiones')
				if event_2 == 'Enviar':
					if value_2['PLAIN'] == True:
						html = False
					elif value_2['HTML'] == True:
						html = True
					if value_2['XLSX'] == '':
						try:
							Enviar_mails(value_2['ASUNTO'], value_2['BODY'], html)
						except:
							print('Hubo un error. Los mails se mandaron hasta el último que aparece. Sino, en el "log.txt" están los detalles')
					else:
						try:
							Enviar_mails(value_2['ASUNTO'], value_2['BODY'], html, value_2['XLSX'])
						except:
							print('Hubo un error. Los mails se mandaron hasta el último que aparece. Sino, en el "log.txt" están los detalles')
		else:
			print('Error en el logeo, probá de vuelta')

				 


#Falta generar log para controlar hasta donde se mandaron los mails, y adjuntar un xlsx basico. Readme. Reemplazo de nombres y apellidos.
#Error handling, selección de files opcional, 
