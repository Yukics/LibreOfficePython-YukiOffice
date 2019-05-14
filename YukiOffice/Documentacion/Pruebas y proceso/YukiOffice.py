import os, time
from threading import Thread

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

def abrir_agenda():
    os.system('cd C:\\Users\\Yuki\\Documents\\Ofimatica\\Final de curso & Agenda.odb')

def abrir_plantilla():
	os.system('cd C:\\Users\\Yuki\\Documents\\Ofimatica\\Final de curso & PlantillaEmpresa.ott')

def abrir_albaran():
    os.system('cd C:\\Program Files\\LibreOffice\\program & soffice --accept="socket,host=localhost,port=2002;urp;" "C:\\Users\\Yuki\\Documents\\Ofimatica\\Final de curso\\PlantillaFactura.ots"')

def abrir_factura():
    os.system('cd C:\\Program Files\\LibreOffice\\program & soffice --accept="socket,host=localhost,port=2002;urp;" "macro://Sin título 1/Standard.AplicarIVA.IVA" "C:\\Users\\Yuki\\Documents\\Ofimatica\\Final de curso\\PlantillaFactura.ots"')

def anadir_numero():
    time.sleep(10)
    os.system('cd C:\\Users\\Yuki\\Documents\\Ofimatica\\Final de curso\\ & "C:\\Program Files\\LibreOffice\\program\\python" contador.py')

def abrir_factura_y_numero():
    Thread(target=abrir_factura).start()
    Thread(target=anadir_numero).start()

def abrir_albaran_y_numero():
    Thread(target=abrir_albaran).start()
    Thread(target=anadir_numero).start()

if __name__ == "__main__":
    ruta, contador = sacar_cfg()
    ruta = ruta.replace('\\','\\\\')

    print(ruta, contador)
    print("1. Abrir Agenda")
    print("2. Abrir Plantilla texto")
    print("3. Abrir Albarán")
    print("4. Abrir Factura")
    eleccion = int(input("Introduce que quieres hacer: "))
    if eleccion == 1:
        abrir_agenda()
    elif eleccion == 2:
        abrir_plantilla()
    elif eleccion == 3:
        abrir_albaran_y_numero()
    elif eleccion == 4:
        abrir_factura_y_numero()


