import PySimpleGUI as sg
import subprocess

def script1(name):
    print('Script 1 '+name)

def script2(name):
    print('Script 2 '+name)

def ExecuteCommandSubprocess(command, *args):      
	try:      
		sp = subprocess.Popen([command, *args], shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)      
		out, err = sp.communicate()      
		if out:      
			print(out.decode("utf-8"))      
		if err:      
			print(err.decode("utf-8"))      
	except:      
		pass      



layout = [
		[sg.T('Parámetros', size=(20, 1))],
		[sg.In('Nombre', key='NAME')],
		[sg.Button('script1'), sg.Button('script2')], 
        [sg.Text('Script output....', size=(20, 1))],      
        [sg.Output(size=(20, 20))]
        #[sg.Text('Manual command', size=(15, 1)), sg.InputText(focus=True), sg.Button('Run', bind_return_key=True)]
		]

# Show the Window to the user
window = sg.Window('Mail', layout)

# Event loop. Read buttons, make callbacks
while True:      
	(event, value) = window.Read()      
	if event is None:      
		break # exit button clicked      
	if event == 'script1':      
		script1(value['NAME'])    
	elif event == 'script2':      
		script2(value['NAME'])     

window.Close()
