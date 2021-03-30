from tkinter import *
from tkinter.filedialog import askopenfilename
import pypyodbc
import datetime
import csv

fenetre =Tk() # initialisation de la fenetre
fenetre.title('Connexion a GDR') # nom de la fenetre

fenetre.geometry('600x600') # taille de la fenetre
fenetre.configure(bg='#36A679') # couleur du fond

fenetre.minsize(480,360) # taille minimum de la fenetre

CheminFile = StringVar() # déclaration de la variable string

CheminFile.set("connexion en cours")
conn = pypyodbc.connect('DSN=GDR;')  # initialisation de la connexion au serveur
cur = conn.cursor()
CheminFile.set("connexion ok") 

def inseeCode():
    fileInsee = 'C:/Users/david/Desktop/ProjetStage/communes2020.csv'
    Lecture = csv.reader(open(fileInsee, "r", encoding='utf-8'), delimiter=',')
    for row in Lecture:
        CodeInsee = row[1], row[7]
        print(CodeInsee)
        
    


# fonction pour afficher la deuxième fenetre
def NewWindow():
    newWindow = Tk()
    newWindow.title('Appli')
    newWindow.geometry('600x600')
    newWindow.configure(bg='#36A679')
    newWindow.minsize(480,360)
    fenetre.destroy() # détruit la première fenetre

    # bouton
    btn = Button(newWindow, text='test', bg='black', fg='white', command= lambda: inseeCode())
    btn.pack()
    btn.place(relx=0.5, rely=0.5, anchor=CENTER)

# texte qui affiche à l'utilisateur si il est connecté ou non
barre = Label(fenetre, textvariable=CheminFile)
barre.pack()
barre.place(relx=0.5, rely=0.5, anchor=CENTER)

# bouton pour passer à la deuxième fenetre
btn_next = Button(fenetre, text='Suivant', bg='black', fg='white', command= lambda: NewWindow())
btn_next.pack()
btn_next.place(relx=0.8, rely=0.9)

fenetre.mainloop()
conn.close()
