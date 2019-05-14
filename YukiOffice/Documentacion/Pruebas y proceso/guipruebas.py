from tkinter import *

tk=Tk()


C = Canvas(tk, bg="blue",  width=500, height=300)
filename = PhotoImage(file = "C:\\Users\\Yuki\\Documents\\Ofimatica\\Final de curso\\Documentacion\\bg.png")
background_label = Label(tk, image=filename)
background_label.place(x=0, y=0, relwidth=1, relheight=1)
tk.resizable(0, 0)
C.pack()

b = Button(tk, text="OK")
b.pack()

tk.mainloop()