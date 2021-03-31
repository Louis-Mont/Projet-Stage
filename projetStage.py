from tkinter import *
from tkinter.filedialog import askopenfilename
import pypyodbc
import datetime
import csv
import xlrd
import xlwt

#-----------------------------------------------------------------------------------------#

# fonction pour récupérer les données utiles de l'insee
def insee():
    CodeInsee = [] 
    CommInsee = []

    # récupération des code communes de l'insee
    fileInsee = 'C:/Users/david/Desktop/ProjetStage/communes2020.csv' # chemin du document de l'insee (.csv le format) 
    Lecture = csv.reader(open(fileInsee, "r", encoding='utf-8'), delimiter=',')
    for row in Lecture:
        Code = row[1] # on récupère les code de l'insee
        Comm = row[7] # on récupère les communes de l'insee
        CodeInsee.append(Code) # insertion des code dans le tableau CodeInsee
        CommInsee.append(Comm) # insertion des communes dans le tableau CommInsee

def VerificationCritere(nom):
    File = xlrd.open_workbook('C:/Users/david/Desktop/ProjetStage/%s.xls' % nom)

    test = File.sheet_by_name('Collecte')

    # Col? =(test.col(?))

def Excel(nom):
    global FeuilleCollect
    global FeuilleVente

    FileExcel = xlwt.Workbook()
    FeuilleCollect = FileExcel.add_sheet('Collecte')

    FeuilleCollect.write(1, 0, "Identifiant unique de la vente")
    FeuilleCollect.write(1, 1, "Commune d'origine du client")
    FeuilleCollect.write(1, 2, "Code postal de la commune d'origine")
    FeuilleCollect.write(1, 3, "code insee")
    FeuilleCollect.write(1, 4, "Flux")
    FeuilleCollect.write(1, 5, "Catégorie")
    FeuilleCollect.write(1, 6, "Sous-Catégorie")
    FeuilleCollect.write(1, 7, "Nombre de produits")
    FeuilleCollect.write(1, 8, "Poids des produits")
    FeuilleCollect.write(1, 9, "Montant HT")
    FeuilleCollect.write(1, 10, "Montant TTC")
    FeuilleCollect.write(1, 11, "IDProduit")

    FeuilleVente = FileExcel.add_sheet('Vente')

    FeuilleVente.write(1, 0, "id produit")
    FeuilleVente.write(1, 1, "id arrivage")
    FeuilleVente.write(1, 2, "Date arrivage")
    FeuilleVente.write(1, 3, "Origine arrivage")
    FeuilleVente.write(1, 4, "id arrivage")
    FeuilleVente.write(1, 5, "Code insee")
    FeuilleVente.write(1, 6, "Flux")
    FeuilleVente.write(1, 7, "Catégorie")
    FeuilleVente.write(1, 8, "Sous-Catégorie")
    FeuilleVente.write(1, 9, "Quantité")
    FeuilleVente.write(1, 10, "Poids")
    FeuilleVente.write(1, 11, "Affectation")

    FileExcel.save('C:/Users/david/Desktop/ProjetStage/%s.xls' % nom) # chemin pour sauvegarder le fichier

# fonction pour afficher la deuxième fenetre
def Window2():

    def Window3():
        newWindow2 = Tk()
        newWindow2.title('Vérification des critères')
        newWindow2.geometry('600x600')
        newWindow2.configure(bg='#36A679')
        newWindow2.minsize(480,360)
        newWindow.destroy()

        btn_verif = Button(newWindow, text='Suivant', bg='black', fg='white', command = lambda: VerificationCritere(var))
        btn_verif.pack()
        btn_verif.place(relx=0.5, rely=0.5, anchor=CENTER)

    newWindow = Tk()
    newWindow.title('Création du fichier excel')
    newWindow.geometry('600x600')
    newWindow.configure(bg='#36A679')
    newWindow.minsize(480,360)
    fenetre.destroy() # détruit la première fenetre
    
    insee() 
    
    TextNameFile = Label(newWindow,text='Nom du fichier :')
    TextNameFile.place(relx=0.5, rely=0.15, anchor= E)
    
    EntryNameFile = Entry(newWindow, bd=3)
    EntryNameFile.place(relx=0.5, rely=0.15, anchor = W)

    def ok():
        global var
        var = EntryNameFile.get()
        if(len(var) > 0):
            btn_save['state'] = NORMAL
        else:
            btn_save['state'] = DISABLED
        
    btn_ok = Button(newWindow, text='Ok', bg='black', fg='white', command = lambda: ok())
    btn_ok.pack()
    btn_ok.place(relx=0.5, rely=0.2, anchor=CENTER)

    # bouton
    btn_save = Button(newWindow, text='Créer le fichier', bg='black', fg='white', command = lambda: Excel(var), state= DISABLED)
    btn_save.pack()
    btn_save.place(relx=0.5, rely=0.4, anchor=CENTER)

    btn_next = Button(newWindow, text='Suivant', bg='black', fg='white', command = lambda: Window3(), state= DISABLED)
    btn_next.pack()
    btn_next.place(relx=0.8, rely=0.9)

    

    newWindow.mainloop()
    
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

# texte qui affiche à l'utilisateur si il est connecté ou non
barre = Label(fenetre, textvariable=CheminFile)
barre.pack()
barre.place(relx=0.5, rely=0.5, anchor=CENTER)

# bouton pour passer à la deuxième fenetre
btn_next = Button(fenetre, text='Suivant', bg='black', fg='white', command= lambda: Window2())
btn_next.pack()
btn_next.place(relx=0.8, rely=0.9)

fenetre.mainloop()
conn.close()
