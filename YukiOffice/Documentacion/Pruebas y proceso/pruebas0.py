import os, time
#Sacar solo la ruta de donde esta el archivo
ruta_actual=os.path.dirname(os.path.abspath(__file__))
#Saca la ruta y el nombre del archivo desde el que se ejecuta
ruta_mas_nombrearchivo=os.path.abspath(__file__)
#(soffice "-accept=socket,host=localhost,port=2002;urp;") para habilitar el modulo pyuno


config=[]
with open ('config.txt', 'rt') as archivo_config:  # Open file lorem.txt for reading of text data.
	for line in archivo_config: # Store each line in a string variable "line"
		config.append(line)

ruta=config[1]
ruta=(ruta[config[1].find('=')+1:len(config[1])]).strip()
contador=config[2]
contador=(contador[config[2].find('=')+1:len(config[2])]).strip()


print(ruta)
print(contador)
time.sleep(5.5) 
