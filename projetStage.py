from tkinter import *
from tkinter.filedialog import askopenfilename
import pypyodbc
import datetime
import csv


NameBase = 'GDR' # nom de votre base ODBC

fenetre =Tk() # initialisation de la fenetre
fenetre.title('Connexion a %s' % NameBase) # nom de la fenetre

fenetre.geometry('600x600') # taille de la fenetre
fenetre.configure(bg='#36A679') # couleur du fond

fenetre.minsize(480,360) # taille minimum de la fenetre

CheminFile = StringVar() # déclaration de la variable string

CheminFile.set("connexion en cours")
conn = pypyodbc.connect(DSN=NameBase)  # initialisation de la connexion au serveur
cur = conn.cursor()
CheminFile.set("connexion ok") 

def insee():
    CodeInsee = [] 
    CommInsee = []

    # récupération des code communes de l'insee
    fileInsee = 'C:/Users/david/Desktop/ProjetStage/communes2020.csv' # chemin du .csv (à changer par rapport à la date)
    Lecture = csv.reader(open(fileInsee, "r", encoding='utf-8'), delimiter=',')
    for row in Lecture:
        Code = row[1] # on récupère les code de l'insee
        Comm = row[7] # on récupère les communes de l'insee
        CodeInsee.append(Code) # insertion des code dans le tableau CodeInsee
        CommInsee.append(Comm) # insertion des communes dans le tableau CommInsee
    


# fonction pour afficher la deuxième fenetre
def Agregation():
    newWindow = Tk()
    newWindow.title('Agrégation')
    newWindow.geometry('600x600')
    newWindow.configure(bg='#36A679')
    newWindow.minsize(480,360)
    fenetre.destroy() # détruit la première fenetre
    
    insee()

    """
    TextValidationInsee = Text(newWindow)
    TextValidationInsee.insert(1.0, )
    TextValidationInsee.pack()
    TextValidationInsee.place(relx=0.5, rely=0.5, anchor=CENTER)
    """

    # bouton
    btn = Button(newWindow, text='Suivant', bg='black', fg='white')
    btn.pack()
    btn.place(relx=0.8, rely=0.9, anchor=CENTER)

    
# texte qui affiche à l'utilisateur si il est connecté ou non
barre = Label(fenetre, textvariable=CheminFile)
barre.pack()
barre.place(relx=0.5, rely=0.5, anchor=CENTER)

# bouton pour passer à la deuxième fenetre
btn_next = Button(fenetre, text='Suivant', bg='black', fg='white', command= lambda: Agregation())
btn_next.pack()
btn_next.place(relx=0.8, rely=0.9)

fenetre.mainloop()
conn.close()
