'''Primero de todo, este programa lo he hecho para que funcione de la forma mas eficiente posible, 
por ello no he utilizado ni objetos ni nada ademas asi es mas facil a la hora de leer y entender 
sobretodo si no estas acostumbrado a leer codigo.


GNU GPL 3.0 '''


#Libreria de interfaz gráfica
import tkinter as tk
from tkinter.filedialog import askopenfilename

#Librerias de sistema
import os, time

#Libreria que permite paralelizar procesos
from threading import Thread

#Ejecuta en paralelo la creación de la factura y la adición al contador
def abrir_factura_y_numero():
    Thread(target=abrir_factura).start()
    Thread(target=anadir_numero).start()
#Ejecuta en paralelo la creación del documento de texto a partir de la plantilla
def abrir_plantilla_proceso():
    Thread(target=abrir_plantilla).start()
#Ejecuta en paralelo la base de datos de la agenda
def abrir_agenda_proceso():
    Thread(target=abrir_agenda).start()
#Ejecuta en paralelo la creación del albarán
def abrir_albaran_proceso():
    Thread(target=abrir_albaran).start()
#Ejecuta en paralelo el albarán selecionado añadiendole las características de factura
def factura_desde_albaran_proceso():
    Thread(target=factura_desde_albaran).start()

#Lineas de comando para ejecutar las plantillas
def abrir_agenda():
    os.system('Plantilla\\Agenda.odb')
def abrir_plantilla():
	os.system('Plantilla\\PlantillaEmpresa.ott')
def abrir_albaran():
    os.system('start soffice --accept="socket,host=localhost,port=2002;urp;" "Plantilla\\PlantillaFactura.ots"')
def abrir_factura():
    os.system('start soffice --accept="socket,host=localhost,port=2002;urp;" "macro://Sin título 1/Standard.AplicarIVA.IVA" "Plantilla\\PlantillaFactura.ots"')

#Se jecuta contador.py
def anadir_numero():
    time.sleep(10)
    ruta=[]
    # Abre config.txt en modo lectura
    with open ('config.txt', 'rt') as configtxt:
        for line in configtxt:
            ruta.append(line)
    contador = ruta[1]
    contador=(contador[ruta[1].find('=')+1:len(ruta[1])]).strip()
    contador = '"'+contador + "\\\\python" + '"' + " contador.py"
    #Ruta donde se encuentra el python del LibreOffice con el modulo uno incorporado
    os.system(contador)



#Abre la ventana específica de creación de facturas, tres cuartos de lo mismo que la principal
def ventana_factura():
    window = tk.Toplevel(root)
    window.title("YukiOffice/Factura")
    window.geometry("500x175")
    frame_f = tk.Frame(window)
    frame_f.pack(side=tk.TOP, fill=tk.X)
    window.resizable(0, 0)

    photo5 = tk.PhotoImage(file="fa.png")
    button5 = tk.Button(frame_f, width=250, height=175, image=photo5, command=factura_desde_albaran_proceso)
    button5.pack(side=tk.LEFT, padx=2, pady=2)
    button5.configure(width=250, height=175)
    button5.image = photo5

    photo6 = tk.PhotoImage(file="nf.png")
    button6 = tk.Button(frame_f, width=250, height=175, image=photo6, command=abrir_factura_y_numero)
    button6.pack(side=tk.LEFT, padx=2, pady=2)
    button6.configure(width=250, height=175)
    button6.image = photo6

#Extrae el nombre del archivo de una ruta completa, esto existe por culpa de la ejecución de la macro de la hoja de calculo que aplica el IVA
def extraer_nombre(ruta):
    longitud = len(ruta)-4
    contador = longitud
    while ruta[contador] != "/":
        contador -= 1
    nombre = ruta[contador+1:longitud]
    return nombre

#Permite seleccionar el albarán del cual hacer factura
def factura_desde_albaran():

    name = askopenfilename(initialdir="C:/Users", filetypes=(("Hoja de cálculo de LibreOffice", "*.ods"), ("All Files", "*.*")), title="Elige un albarán.")                               
   	
   	#Comprueba la extensión del archivo
    ext = name[len(name)-3:len(name)]
    if ext != "ods" and ext != "ots" and ext != "xls" and ext != "xlt" and name != "":
    	os.system('start msgbox.vbs')


    #El comando con la ruta del archivo que se quiere convertir a factura
    ejecutar_albaran_a_factura = 'start soffice --accept="socket,host=localhost,port=2002;urp;" "%s" ' % name
    
    #Saca el nombre del archivo de la ruta entera para usarlo en la macro 
    archivo = extraer_nombre(name)
    macro = 'macro://%s/Standard.AplicarIVA.IVA' % archivo
 	
 	#La ejecucion de la macro como tal
    ejecutar_macro = 'start soffice "%s"' % macro
    
    #Intenta ejecutar el comando y en caso de error muestra una ventana de error (solo se muestra en caso de que no se pueda ejecutar por que se es muy inutil)
    try:
        os.system(ejecutar_albaran_a_factura)
        Thread(target=anadir_numero).start()
        time.sleep(6)
        Thread(target=os.system(ejecutar_macro)).start()
    except:
    	print("Ha habido un error")     

#Programa principal, genera la interfaz gráfica
if __name__ == "__main__":
    root = tk.Tk()
   	
   	#El icono que aparece arriba a la izquierda de la ventana
    root.wm_iconbitmap(default='icono.ico')
    
    #Se crea un frame, como un lugar donde poner cosas la agenda y la plantilla de texto
    frame1 = tk.Frame(root)
    frame1.pack(side=tk.TOP, fill=tk.X)
    
    #Se crea un frame, para el boton de la plantilla del albaran y factruas
    frame2 = tk.Frame(root)
    frame2.pack(side=tk.BOTTOM, fill=tk.X)
    
    #Titulo de la ventana y el programa
    root.title("YukiOffice")
    
    #Tamaño de la ventana en pixeles
    root.geometry("500x350")
    root.configure(background="black")
    
    #Que no se pueda modificar el tamaño de la ventana, hacerlo de la otra forma no merecia la pena
    root.resizable(0, 0)

    #Cada uno de los botones de la interfaz principal
    photo1 = tk.PhotoImage(file="image10.png")
    button1 = tk.Button(frame1, width=250, height=175, image=photo1, command=abrir_plantilla_proceso)
    button1.pack(side=tk.LEFT, padx=2, pady=2)
    button1.configure(width=250, height=175)
    button1.image = photo1

    photo2 = tk.PhotoImage(file="image12.png")
    button2 = tk.Button(frame1, width=250, height=175, image=photo2, command=abrir_agenda_proceso)
    button2.pack(side=tk.LEFT, padx=2, pady=2)
    button2.configure(width=250, height=175)
    button2.image = photo2

    photo3 = tk.PhotoImage(file="image13.png")
    button3 = tk.Button(frame2, width=250, height=175, image=photo3, command=abrir_albaran)
    button3.pack(side=tk.LEFT, padx=2, pady=2)
    button3.configure(width=250, height=175)
    button3.image = photo3

    photo4 = tk.PhotoImage(file="image14.png")
    button4 = tk.Button(frame2, width=250, height=175, image=photo4, command=ventana_factura)
    button4.pack(side=tk.LEFT, padx=2, pady=2)
    button4.configure(width=250, height=175)
    button4.image = photo4


    root.mainloop()