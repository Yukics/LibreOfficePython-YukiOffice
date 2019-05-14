import tkinter as tk
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
    os.system('Plantilla\\Agenda.odb')

def abrir_plantilla():
	os.system('Plantilla\\PlantillaEmpresa.ott')

def abrir_albaran():
    os.system('start soffice --accept="socket,host=localhost,port=2002;urp;" "Plantilla\\PlantillaFactura.ots"')

def abrir_factura():
    os.system('start soffice --accept="socket,host=localhost,port=2002;urp;" "macro://Sin título 1/Standard.AplicarIVA.IVA" "C:\\Users\\Yuki\\Documents\\Ofimatica\\Final de curso\\PlantillaFactura.ots"')

def anadir_numero():
    time.sleep(10)
    os.system('"C:\\Program Files\\LibreOffice\\program\\python" contador.py')

def abrir_factura_y_numero():
    Thread(target=abrir_factura).start()
    Thread(target=anadir_numero).start()

def abrir_albaran_y_numero():
    Thread(target=abrir_albaran).start()
    Thread(target=anadir_numero).start()

def abrir_plantilla_proceso():
    Thread(target=abrir_plantilla).start()

def abrir_agenda_proceso():
    Thread(target=abrir_agenda).start()


if __name__ == "__main__":
    root = tk.Tk()
    root.wm_iconbitmap(default='icono.ico')
    frame1 = tk.Frame(root)
    frame1.pack(side=tk.TOP, fill=tk.X)
    frame2 = tk.Frame(root)
    frame2.pack(side=tk.BOTTOM, fill=tk.X)
    root.title("YukiOffice")
    root.geometry("500x350")
    root.configure(background="black")
    root.resizable(0, 0)

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
    button3 = tk.Button(frame2, width=250, height=175, image=photo3, command=abrir_albaran_y_numero)
    button3.pack(side=tk.LEFT, padx=2, pady=2)
    button3.configure(width=250, height=175)
    button3.image = photo3

    photo4 = tk.PhotoImage(file="image14.png")
    button4 = tk.Button(frame2, width=250, height=175, image=photo4, command=abrir_factura_y_numero)
    button4.pack(side=tk.LEFT, padx=2, pady=2)
    button4.configure(width=250, height=175)
    button4.image = photo4


    root.mainloop()