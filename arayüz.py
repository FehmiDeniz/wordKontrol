from tkinter import filedialog
from tkinter import *
import tkinter as tk
import os
import docx
from Data import cift, paragraf, girisbolum, dosyaac

root = tk.Tk()
root.title('Word Kontrol')
root.geometry("200x100")
root.eval('tk::PlaceWindow . center')
root.resizable(False, False)
root.iconphoto(False, tk.PhotoImage(file='icon.png'))



dosyaac.Clear_Console()
print("            Lütfen bir word dosyası seçin.")
print("                      _________________")
print("                     | | ___________ |o|")
print("                     | | ___________ | |")
print("                     | | ___________ | |")
print("                     | | ___________ | |")
print("                     | |_____________| |")
print("                     |     _______     |")
print("                     |    |       |   ||")
print("                     | KD |       |   V|")
print("                     |____|_______|____|")



def browsefunc():


    root.filename =  filedialog.askopenfilename(initialdir = "/",title = "Word Dosyasını Seçin",filetypes = (("Word dosyaları",".docx"),("Tüm Dosyalar",".*")))
    label1.config(text='{}'.format(os.path.basename(root.filename)))
    if os.path.exists(root.filename):
        b1.config(text="Yeniden Seç")
        dosyaac.Clear_Console()
        #print (root.filename)
        dosyayol=root.filename
        #print(dosyayol)
        cift.cifttirnakfonk(dosyayol)
        paragraf.Paragrafmidegilmi(dosyayol)
        girisbolum.Iceriyomu(dosyayol)
        #os.startfile('WordRapor.docx')
        filename='WordRapor.docx'
        dosyaac.Open_file(filename)



b1=tk.Button(root,text="Dosya Seç",font=40,command=browsefunc)
spaceLabel = tk.Label(root, text= "                     ")
label1 = tk.Label(root, text= "Lütfen bir word dosyası seçin.")
spaceLabel.pack()
label1.pack()
b1.pack()



root.mainloop()
