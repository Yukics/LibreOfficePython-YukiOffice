import os, time

'''print(os.path.dirname(os.path.abspath(__file__)))
os.system('dir c:\\')
soffice.exe -invisible --calc --accept="socket,host=localhost,port=2002;urp;"
#Abrir mi plantilla de facturas habilitando la introducción mediante script
scalc --accept="socket,host=localhost,port=2002;urp;" "C:\Users\Yuki\Documents\Ofimatica\Final de curso\PlantillaFactura.ots"
#Ejecutar desde la ruta del Python de office mmi script
C:\Program Files\LibreOffice\program>python "C:\Users\Yuki\Documents\Ofimatica\Final de curso\hola.py"
#Ejecutar el script con el python de Office
"C:\Program Files\LibreOffice\program\python.exe" hola.py
#Sacar solo la ruta de donde esta el archivo
ruta_actual=os.path.dirname(os.path.abspath(__file__))
#Saca la ruta y el nombre del archivo desde el que se ejecuta
ruta_mas_nombrearchivo=os.path.abspath(__file__)
#Ejecutar macro desde consola
soffice "macro://Sin título 1/Standard.AplicarIVA.IVA"
soffice "macro:///Standard.AplicarIVA.IVA" "C:\Users\Yuki\Documents\Ofimatica\Final de curso\PlantillaFactura.ots"
'''
def abrir_agenda():
	os.system('cd C:\\Users\\Yuki\\Documents\\Ofimatica\\Final de curso & Agenda.odb')
def abrir_plantilla():
	os.system('cd C:\\Users\\Yuki\\Documents\\Ofimatica\\Final de curso & PlantillaEmpresa.ott')
def abrir_albaran():
	os.system('cd C:\\Users\\Yuki\\Documents\\Ofimatica\\Final de curso & PlantillaFactura.ott')

time.sleep(5.5) 

#########################################################################################################
ABRIR FICHERO DESDE COMANDO							 			#													#
													#
#########################################################################################################


"C:\Program Files (x86)\OpenOffice 4\program\soffice.exe" --accept="socket,host=localhost,port=2002;urp;" NombreArchivo.extension

#########################################################################################################
CODIGO PARA SACAR LA RUTA DE DONDE SE ENCUENTRE LA RUTA Y EL NUMERO DE LA FACTURA			#													#
													#
#########################################################################################################
	

def sacar_cfg():
    config=[]
    with open ('config.txt', 'rt') as archivo_config:  # Abre config.txt en modo lectura
        for line in archivo_config: # Guarda cada linea en una poosicion de la lista
            config.append(line)


    ruta=config[1]
    ruta=(ruta[config[1].find('=')+1:len(config[1])]).strip()	#Saca la ruta de donde se supone que esta el soffice
    contador=config[2]
    contador=(contador[config[2].find('=')+1:len(config[2])]).strip()  #Saca el numero contador para los albaranes y facturas
    return ruta, contador

def añadir_uno_contador ():
    with open('config.txt', 'rt') as archivo_config:  # Abre config.txt en modo lectura
        filedata = archivo_config.read() #Vuelca el archivo en filedata
    contador_2 = str(int(contador)+1) #Añade uno al contador
    filedata = filedata.replace(contador, contador_2 ) #Remplaza el nuevo numero contador
    with open('config.txt', 'w') as archivo_config:
        archivo_config.write(filedata) #Sobreescribe el archivo de cfg					





#########################################################################################################
CODIGO PARA INSERTAR TEXTO EN UNA CELDA									#									
													#
#########################################################################################################
import socket  # only needed on win32-OOo3.0.0
import uno

# get the uno component context from the PyUNO runtime
localContext = uno.getComponentContext()

# create the UnoUrlResolver
resolver = localContext.ServiceManager.createInstanceWithContext(
				"com.sun.star.bridge.UnoUrlResolver", localContext )

# connect to the running office
ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
smgr = ctx.ServiceManager

# get the central desktop object
desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)

# access the current writer document
model = desktop.getCurrentComponent()
# access the active sheet
active_sheet = model.CurrentController.ActiveSheet

# access cell C4
cell1 = active_sheet.getCellRangeByName("D6")

# set text inside
cell1.String = "1"					



