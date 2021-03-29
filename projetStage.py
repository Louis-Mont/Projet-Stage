# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

from tkinter import *
from tkinter.filedialog import askopenfilename

fenetre =Tk()
fenetre.title('Application d\'extraction')

fenetre.geometry('600x600')
fenetre.configure(bg='#36A679')

CheminFile = StringVar()

def set_file():
    CheminFile.set(askopenfilename(initialdir ="/", filetypes = [ ("all files", "*.*")]))
    
barre = Entry(fenetre, textvariable=CheminFile, width=50)
barre.pack()
barre.place(relx=0.5, rely=0.5, anchor=CENTER)

btn_recherche = Button(fenetre, text='Parcourir', bg='black', fg='white', command=lambda : set_file())
btn_recherche.pack()
btn_recherche.place(relx=0.5, rely=0.55, anchor=CENTER)

btn_next = Button(fenetre, text='Suivant', bg='black', fg='white')
btn_next.pack()