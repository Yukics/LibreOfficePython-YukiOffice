import socket  
import uno

#Extrae el numero de factura
def sacar_cfg():
    config=[]
    with open ('config.txt', 'rt') as archivo_config:  # Abre config.txt en modo lectura
        for line in archivo_config: # Guarda cada linea en una poosicion de la lista
            config.append(line)
    contador=config[0]
    contador=(contador[config[0].find('=')+1:len(config[0])]).strip()  #Saca el numero contador para los albaranes y facturas
    return contador

#Añade 1 al numero de factura
def anadir_uno_contador (contador):
    with open('config.txt', 'rt') as archivo_config:  # Abre config.txt en modo lectura
        filedata = archivo_config.read() #Vuelca el archivo en filedata
    contador_2 = str(int(contador)+1) #Anade uno al contador
    filedata = filedata.replace(contador, contador_2 ) #Remplaza el nuevo numero contador
    with open('config.txt', 'w') as archivo_config:
        archivo_config.write(filedata) #Sobreescribe el archivo de cfg

localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext(
				"com.sun.star.bridge.UnoUrlResolver", localContext )
ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
model = desktop.getCurrentComponent()
active_sheet = model.CurrentController.ActiveSheet

cell1 = active_sheet.getCellRangeByName("D6")
cell2 = active_sheet.getCellRangeByName("C6")

contador = sacar_cfg()
anadir_uno_contador(contador)

cell1.String = contador
cell2.String = "Nº Factura: " 