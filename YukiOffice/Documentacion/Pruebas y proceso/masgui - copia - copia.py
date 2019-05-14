import tkinter as tk

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
button1 = tk.Button(frame1, width=250, height=175, image=photo1, text="optional text")
button1.pack(side=tk.LEFT, padx=2, pady=2)
button1.configure(width=250, height=175)
button1.image = photo1

photo2 = tk.PhotoImage(file="image12.png")
button2 = tk.Button(frame1, width=250, height=175, image=photo2, text="optional text")
button2.pack(side=tk.LEFT, padx=2, pady=2)
button2.configure(width=250, height=175)
button2.image = photo2

photo3 = tk.PhotoImage(file="image13.png")
button3 = tk.Button(frame2, width=250, height=175, image=photo3, text="optional text")
button3.pack(side=tk.LEFT, padx=2, pady=2)
button3.configure(width=250, height=175)
button3.image = photo3

photo4 = tk.PhotoImage(file="image14.png")
button4 = tk.Button(frame2, width=250, height=175, image=photo4, text="optional text")
button4.pack(side=tk.LEFT, padx=2, pady=2)
button4.configure(width=250, height=175)
button4.image = photo4




root.mainloop()