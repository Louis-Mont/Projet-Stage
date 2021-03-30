from tkinter import *
from tkinter.filedialog import askopenfilename
import pypyodbc

fenetre =Tk() # initialisation de la fenetre
fenetre.title('Connexion a GDR') # nom de la fenetre

fenetre.geometry('600x600') # taille de la fenetre
fenetre.configure(bg='#36A679') # couleur du fond

fenetre.minsize(480,360) # taille minimum de la fenetre

CheminFile = StringVar() # déclaration de la variable string
CheminFile.set("Veuillez vous connectez a GDR")

# fonction qui permet de se connecter à la base GDR
def dsn():
    CheminFile.set("connexion en cours")
    conn = pypyodbc.connect('DSN=GDR;')  # initialisation de la connexion au serveur
    cur = conn.cursor()
    CheminFile.set("connexion ok") 

# fonction qui permet d'activer ou désactiver le bouton btn_next
def changeState():
    if (btn_next['state'] == NORMAL):
        btn_next['state'] = DISABLED
    else:
        btn_next['state'] = NORMAL

# fonction pour afficher la deuxième fenetre
def NewWindow():
    newWindow = Tk()
    newWindow.title('Appli')
    newWindow.geometry('600x600')
    newWindow.configure(bg='#36A679')
    newWindow.minsize(480,360)
    fenetre.destroy() # détruit la première fenetre

    # bouton
    btn_agreger = Button(newWindow, text='Agréger', bg='black', fg='white')
    btn_agreger.pack()
    btn_agreger.place(relx=0.5, rely=0.5, anchor=CENTER)

# texte qui affiche à l'utilisateur si il est connecté ou non
barre = Label(fenetre, textvariable=CheminFile)
barre.pack()
barre.place(relx=0.5, rely=0.5, anchor=CENTER)

# bouton pour se connecter 
btn_recherche = Button(fenetre, text='Connexion', bg='black', fg='white', command=lambda : [changeState(), dsn() ])
btn_recherche.pack()
btn_recherche.place(relx=0.5, rely=0.55, anchor=CENTER)

# bouton pour passer à la deuxième fenetre
btn_next = Button(fenetre, text='Suivant', bg='black', fg='white', state = DISABLED, command= lambda: NewWindow())
btn_next.pack()
btn_next.place(relx=0.8, rely=0.9)

fenetre.mainloop()
