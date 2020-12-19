from tkinter import filedialog
from tkinter import *
import tkinter as tk
import os
import docx
from Data import cift 
from Data import paragraf 
from Data import girisbolum 
from Data import dosyaac
root = tk.Tk()
root.title('Word Kontrol')



dosyaac.Clear_Console()
print("            Lütfen kayıtlı bir word dosyası seçin.")
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
    
    
    root.filename =  filedialog.askopenfilename(initialdir = "/",title = "Tez Dosyasını Seçin",filetypes = (("Word dosyaları",".docx"),("Tüm Dosyalar",".*")))
    ent1.insert(tk.END, root.filename)
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
ent1=tk.Entry(root,font=40)
ent1.grid(row=2,column=2)
b1=tk.Button(root,text="Dosya Seçin",font=40,command=browsefunc)
b1.grid(row=2,column=4)


root.mainloop()

