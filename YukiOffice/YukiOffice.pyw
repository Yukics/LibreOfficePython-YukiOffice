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

#Suma 1 al contador cada vez que se ejecuta del documento config.txt
def anadir_numero():
    time.sleep(10)
    os.system('"C:\\Program Files\\LibreOffice\\program\\python" contador.py')

#Abre la ventana específica de creación de facturas
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

#Extrae el nombre del archivo de una ruta completa
def extraer_nombre(ruta):
    longitud = len(ruta)-4
    contador = longitud
    while ruta[contador] != "/":
        contador -= 1
    nombre = ruta[contador+1:longitud]
    print(nombre)
    return nombre

#Permite seleccionar el albarán del cual hacer factura
def factura_desde_albaran():
    def OpenFile():
        name = askopenfilename(initialdir="C:/Users",
                               filetypes=(("Office Document Sheet", "*.ods"), ("All Files", "*.*")),
                               title="Elige un albarán."
                               )
        ejecutar_albaran_a_factura = 'start soffice --accept="socket,host=localhost,port=2002;urp;" "%s" ' % name
        # Using try in case user types in unknown file or closes without choosing a file.
        archivo = extraer_nombre(name)
        macro = 'macro://%s/Standard.AplicarIVA.IVA' % archivo
        print(macro)
        ejecutar_macro = 'start soffice "%s"' % macro
        try:
            os.system(ejecutar_albaran_a_factura)
            Thread(target=anadir_numero).start()
            time.sleep(6)
            Thread(target=os.system(ejecutar_macro)).start()
        except:
            print("No file exists")

    OpenFile()

#Programa principal, genera la interfaz gráfica
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