from io import StringIO
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import sqlite3
import datetime
from xlwt import Workbook
import xlrd
import csv
import os

# connexion à la base de données sinon stop le programme et affiche l'erreur
try:
    connect = sqlite3.connect("finale.db")
    cur = connect.cursor()
except IOError:
    print(IOError)

def Is_checkStruct():
    if chkValueStruct.get() == True:
        ComboStruct.config(state='normal')
        BtnAdd.config(state='normal')
        BtnDel.config(state='normal')
    else:
        ComboStruct.config(state='disabled')
        BtnAdd.config(state='disabled')
        BtnDel.config(state='disabled')

def Is_check1():
    if chkValue1.get() == True:
        EntryInsee.config(state='normal')
        CheckMany.config(state='disabled')
    else:
        EntryInsee.config(state='disabled')
        CheckMany.config(state='normal')

def Is_check2():
    if chkValue2.get() == True:
        BtnInsee.config(state='normal')
        CheckOne.config(state='disabled')
    else:
        BtnInsee.config(state='disabled')
        CheckOne.config(state='normal')
        listInsee.delete(0,END)
        LabelFile.config(text='')

def FileInsee(): 
    filename = filedialog.askopenfilename(initialdir = "/",
                                        title="sélectionner votre fichier",
                                        filetypes=(("Fichier Texte","*.txt"),
                                                    ("Fichier pdf","*.pdf")))

    LabelFile.config(text = filename)
    fichier = open(str(filename), "r")
    lines = fichier.readlines()
    fichier.close()
    listInsee.delete(0,END)
    i=1
    for ligne in lines:
        listInsee.insert(i,ligne.strip())
        i+=1

def RequeteStruct():

    if chkValue1.get() == True and chkValueStruct.get() == False:
        val = EntryInsee.get()
        cur.execute("SELECT Recyclerie FROM Organisation, Insee, Commune WHERE Organisation.Id_Recyclerie = Commune.Id_Recyclerie AND Commune.Id_insee = Insee.Id_Insee AND Code = '%s' "%\
                            (val))
        StructList = cur.fetchall()
        for struct in StructList:
            NomStruct = struct[0]
    elif chkValue2.get() == True and chkValueStruct.get() == False:
        val = listInsee.get(0, END)
        for code in val:
            print(code)
            cur.execute("SELECT Recyclerie FROM Organisation, Insee, Commune WHERE Organisation.Id_Recyclerie = Commune.Id_Recyclerie AND Commune.Id_insee = Insee.Id_Insee AND Code = '%s'"%\
                            (code))
            print(cur.fetchall())
    elif chkValue1.get() == True and chkValueStruct.get() == True:
        print("rt")
    elif chkValue2.get() == True and chkValueStruct.get() == True:
        print('kl')

def new_window(listStruct):
    RequeteStruct()

    StructList = listStruct.get(0,END)

    FirstFen.destroy() # Détruit la première fenêtre 
    SecondFen = Tk() # initialisation de la première fenêtre
    SecondFen.title("Options d'importation")

    #Partie catégorie :
    LabelCat = Label(SecondFen, text='Catégorie :', font = 60)
    LabelCat.grid(row=0,column=0, ipady = 30)

    Id = 1

    cur.execute("SELECT Catégorie FROM Catégorie, Produit WHERE Catégorie.Id_catégorie = Produit.Id_catégorie and Id_recyclerie = '%s' GROUP BY Produit.Id_catégorie HAVING count(Produit.Id_catégorie) != 0"%\
                            (Id))
    Cat = cur.fetchall()
    List = []
    for row in Cat:
        List.append(row[0])
    
    Combo = ttk.Combobox(SecondFen, values = List, width=29)
    Combo.set("Choississez la/les catégorie(s)")
    Combo.grid(row=0,column = 1)

    Combo.bind("<<ComboboxSelected>>", lambda e: Combo.get())

    BtnAdd2 = Button(SecondFen, text='Ajouter', command=lambda:addCat(listCat, Combo))
    BtnAdd2.grid(row=0,column=2)

    BtnDel2 = Button(SecondFen, text='Supprimer', command=lambda:delCat(listCat))
    BtnDel2.grid(row=0,column=3,padx=5)

    listCat = Listbox(SecondFen, width=50)
    listCat.grid(row=0, column=5)

    #Partie temps :
    Labeltemps = Label(SecondFen, text='Temps :', font = 60)
    Labeltemps.grid(row=1,column=0, ipady = 10)

    cur.execute("SELECT Min(date) FROM Arrivage WHERE Id_recyclerie = '%s'"%\
                        (Id))
    date_bdd = cur.fetchone()[0]
    an_bdd = date_bdd[:4]
    an=datetime.date.today().year 
    an=int(an)                                                    
    jours,mois,ans=[str(i).zfill(2) for i in range(1,32)],[str(i).zfill(2) for i in range(1,13)],[i for i in range(int(an_bdd),an+1)]
    mess_debut=Label(SecondFen,text="choisissez la date de début :")
    mess_debut.grid(row=2,column=1, padx = 10)
    choixjour1=ttk.Combobox(SecondFen,values=jours)
    choixjour1.grid(row=2,column=2)
    choixjour1.set('01')
    choixmois1=ttk.Combobox(SecondFen,values=mois)
    choixmois1.grid(row=2,column=3)
    choixmois1.set('01')
    choixan1=ttk.Combobox(SecondFen,values=ans)
    choixan1.grid(row=2,column=4)
    choixan1.set(int(an_bdd))
    mess_fin=Label(SecondFen,text="choisissez la date de fin :")
    mess_fin.grid(row=3,column=1, pady = 10)
    choixjour2=ttk.Combobox(SecondFen,values=jours)
    choixjour2.grid(row=3,column=2)
    choixjour2.set('31')
    choixmois2=ttk.Combobox(SecondFen,values=mois)
    choixmois2.grid(row=3,column=3)
    choixmois2.set('12')
    choixan2=ttk.Combobox(SecondFen,values=ans)
    choixan2.grid(row=3,column=4)
    choixan2.set(an)

    #Partie quantitative :
    LabelQte = Label(SecondFen, text='Quantitative :', font = 60)
    LabelQte.grid(row=4,column=0, ipady = 40,padx= 5)

    chkValue1 = BooleanVar()
    chkValue1.set(False)
    chkValue2 = BooleanVar()
    chkValue2.set(False)
    chkValue3 = BooleanVar()
    chkValue3.set(False)
    chkValue4 = BooleanVar()
    chkValue4.set(False)
    chkValue5 = BooleanVar()
    chkValue5.set(False)
    Chkbox1 = Checkbutton(SecondFen, text = 'Poids collecté (en Kg)', var = chkValue1)
    Chkbox1.grid(row=5,column=1)
    Chkbox2 = Checkbutton(SecondFen, text = 'Poids vendu (en Kg)', var = chkValue2)
    Chkbox2.grid(row=5,column=2)
    Chkbox3 = Checkbutton(SecondFen, text = 'Chiffre d\'affaire (en €)', var = chkValue3)
    Chkbox3.grid(row=5,column=3)
    Chkbox4 = Checkbutton(SecondFen, text = 'Nombre d\'objet collecté', var = chkValue4)
    Chkbox4.grid(row=6,column=1)
    Chkbox5 = Checkbutton(SecondFen, text = 'Nombre d\'objet vendu', var = chkValue5)
    Chkbox5.grid(row=6,column=2)

    BtnExport = Button(SecondFen, text='Exporter', command=lambda:export(SecondFen, listCat, choixjour1, choixmois1, choixan1, choixjour2, choixmois2, choixan2, StructList, chkValue1, chkValue2, chkValue3, chkValue4, chkValue5))
    BtnExport.grid(row=10,column=5, padx = 40, pady = 20)

    SecondFen.mainloop()

def addCat(listCat, Combo):
    listFiles = listCat.get(0,END)
    value = Combo.get()
    if value not in listFiles and value != 'Choississez la/les catégorie(s)':
        listCat.insert(END, value)

def delCat(listCat):
    itemSelected = listCat.curselection()
    listCat.delete(itemSelected[0])

def addStruct():
    listFiles = listStruct.get(0,END)
    value = ComboStruct.get()
    if ComboStruct.get() == 'tout':
        for struct in List:
            listStruct.insert(END, struct)
        listStruct.delete(0)
    if value not in listFiles and value != 'tout' and value != 'Choississez la/les structure(s)':
        listStruct.insert(END, value)

def delStruct():
    itemSelected = listStruct.curselection()
    listStruct.delete(itemSelected[0])

def export(SecondFen, listCat, jour1, mois1, an1, jour2, mois2, an2, StructList, chk1, chk2, chk3, chk4, chk5):
    filename = filedialog.asksaveasfilename(defaultextension = '.xls',
                                            filetypes = [("xls files", '*.xls')])
    
    size = listCat.size()

    ListCategorie = [c for c in listCat.get(0,END)] # création d'une liste des catégories sélectionnées 

    jourfirst = jour1.get()
    moisfirst = mois1.get()
    anfirst = an1.get()
    joursec = jour2.get()
    moissec = mois2.get()
    ansec = an2.get()
    dateFirst =  anfirst + "/" + moisfirst + "/" + jourfirst
    dateEnd = ansec + "/" + moissec + "/" + joursec
    classeurexport=Workbook()
    pageexport=classeurexport.add_sheet("EXPORT")
    pageexport.write(0,0,"Données du : " + dateFirst + " au " + dateEnd)
    pageexport.write(2,0,"Les recycleries : ")
    y = 1
    for struct in StructList:
        pageexport.write(2,y,struct)
        y+=1
    pageexport.write(4,0,"Catégories")
    if size == 0: # si la liste est vide insère toutes les catégories ayant au minimum un arrivage dans le fichier
        i=5
        cur.execute("SELECT Catégorie FROM Catégorie GROUP BY Id_Catégorie")
        Cat = cur.fetchall()
        ListCategorie = [val[0] for val in Cat]
        for val in Cat:
            pageexport.write(i,0,val[0])
            i+=1
        pageexport.write(i+1,0,"total :")
    else:
        i=5
        for c in ListCategorie:
            pageexport.write(i,0,c)
            i+=1
        pageexport.write(i+1,0,"total :")
    TailleCat = len(ListCategorie)
    TailleStruct = len(StructList)
    y = 1
    if chk1.get() == True:
        totalPoidsC = 0 # variable pour le poids collecté sur l'ensemble des catégories 
        indice = 0 # variable pour l'indexation selon la taille de la liste des catégories
        pageexport.write(4,y,"Poids collecté (en kg)")
        x = 5 # indice pour la ligne dans le fichier xls
        while (indice != TailleCat):
            indice2 = 0 # variable pour l'indexation selon la taille de la liste des structures
            val = 0 # variable qui va recevoir les valeurs de la requete
            while (indice2 != TailleStruct):
                cur.execute("SELECT sum(Poids)*nombre FROM Produit, Catégorie, Arrivage, Organisation WHERE Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_catégorie = Catégorie.Id_catégorie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                (StructList[indice2], str(dateFirst),str(dateEnd),ListCategorie[indice]))
                PoidsCollect = cur.fetchall()
                if len(PoidsCollect) == 0:
                    val+=0
                else:
                    for poids in PoidsCollect:
                        val+=poids[0]
                        totalPoidsC+= poids[0]       
                indice2+=1
            if indice2 == TailleStruct:
                    if val != 0:
                        pageexport.write(x, y, val)
                        x +=1
                    else:
                        pageexport.write(x, y, 'NR')
                        x +=1
            indice+=1
        pageexport.write(x+1,y,totalPoidsC)
        y+=1 # on passe a la colonne suivante pour les données suivantes
    if chk2.get() == True:
        totalPoidsV = 0
        indice = 0
        pageexport.write(4,y,"Poids vendu (en kg)")
        x = 5 # indice pour la ligne dans le fichier xls
        while (indice != TailleCat):
            indice2 = 0 # variable pour l'indexation selon la taille de la liste des structures
            val = 0 # variable qui va recevoir les valeurs de la requete
            while (indice2 != TailleStruct):
                cur.execute("SELECT sum(Lignes_vente.Poids) FROM Lignes_vente, Vente, Catégorie, Organisation WHERE Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_vente = Vente.Id_Vente AND Lignes_vente.Id_catégorie=Catégorie.Id_catégorie AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                        (StructList[indice2], str(dateFirst),str(dateEnd),ListCategorie[indice]))
                PoidsVendu = cur.fetchall()
                if len(PoidsVendu) == 0:
                    val+=0
                else:
                    for poids in PoidsVendu:
                        val+= poids[0]
                        totalPoidsV += poids[0]
                indice2+=1
            if indice2 == TailleStruct:
                    if val != 0:
                        pageexport.write(x, y, val)
                        x +=1
                    else:
                        pageexport.write(x, y, 'NR')
                        x +=1
            indice+=1
        pageexport.write(x+1,y,totalPoidsV)
        y +=1
    if chk3.get() == True:
        totalChiffre = 0
        indice = 0
        pageexport.write(4,y,"Chiffre d'affaire (en €)")
        x = 5 # indice pour la ligne dans le fichier xls
        while (indice != TailleCat):
            indice2 = 0 # variable pour l'indexation selon la taille de la liste des structures
            val = 0 # variable qui va recevoir les valeurs de la requete
            while (indice2 != TailleStruct):
                cur.execute("SELECT sum(Montant), Catégorie FROM Vente, Catégorie, Lignes_vente, Organisation WHERE Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_vente=Vente.Id_Vente AND Lignes_vente.Id_catégorie=Catégorie.Id_catégorie AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                        (StructList[indice2], str(dateFirst),str(dateEnd),ListCategorie[indice]))
                ChiffreAffaire = cur.fetchall()
                if len(ChiffreAffaire) == 0:
                    val+=0
                else:
                    for chiffre in ChiffreAffaire:
                        val+=chiffre[0]
                        totalChiffre+= chiffre[0]
                indice2+=1
            if indice2 == TailleStruct:
                    if val != 0:
                        pageexport.write(x, y, val)
                        x +=1
                    else:
                        pageexport.write(x, y, 'NR')
                        x +=1
            indice+=1
        pageexport.write(x+1,y,totalChiffre)
        y +=1
    if chk4.get() == True:
        totalNbrC = 0
        indice = 0
        pageexport.write(4,y,"Nombre d'objet collecté")
        x = 5 # indice pour la ligne dans le fichier xls
        while (indice != TailleCat):
            indice2 = 0 # variable pour l'indexation selon la taille de la liste des structures
            val = 0 # variable qui va recevoir les valeurs de la requete
            while (indice2 != TailleStruct):
                cur.execute("SELECT count(Id_Produit)*nombre, Catégorie FROM Produit, Arrivage, Catégorie, Organisation WHERE Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_catégorie=Catégorie.Id_catégorie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                    (StructList[indice2], str(dateFirst),str(dateEnd),ListCategorie[indice]))
                NombreCollect = cur.fetchall()
                if len(NombreCollect) == 0:
                    val+=0
                else:
                    for nbrC in NombreCollect:
                        val+=nbrC[0]
                        totalNbrC+=nbrC[0]
                indice2+=1
            if indice2 == TailleStruct:
                    if val != 0:
                        pageexport.write(x, y, val)
                        x +=1
                    else:
                        pageexport.write(x, y, 'NR')
                        x +=1
            indice+=1
        pageexport.write(x+1,y,totalNbrC)
        y +=1
    if chk5.get() == True:
        totalNbrV = 0
        indice = 0
        pageexport.write(4,y,"Nombre d'objet vendu")
        x = 5 # indice pour la ligne dans le fichier xls
        while (indice != TailleCat):
            indice2 = 0 # variable pour l'indexation selon la taille de la liste des structures
            val = 0 # variable qui va recevoir les valeurs de la requete
            while (indice2 != TailleStruct):
                cur.execute("SELECT count(Id_ligne_vente), Catégorie FROM Lignes_vente, Vente, Catégorie, Organisation WHERE Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_catégorie=Catégorie.Id_catégorie AND Vente.Id_Vente=Lignes_vente.Id_vente AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                    (StructList[indice2], str(dateFirst),str(dateEnd),ListCategorie[indice]))
                NombreVendu = cur.fetchall()
                if len(NombreVendu) == 0:
                    val+=0
                else:
                    for nbrV in NombreVendu:
                        val+=nbrV[0]
                        totalNbrV+=nbrV[0]
                indice2+=1
            if indice2 == TailleStruct:
                    if val != 0:
                        pageexport.write(x, y, val)
                        x +=1
                    else:
                        pageexport.write(x, y, 'NR')
                        x +=1
            indice+=1
        pageexport.write(x+1,y,totalNbrV)
        y +=1
    classeurexport.save(filename)
    TailleFile = len(filename)
    FileCSV = filename[:TailleFile-4]
    with xlrd.open_workbook(filename) as wb:
        sh = wb.sheet_by_index(0)  
        with open(FileCSV + '.csv', 'w') as f:   
            c = csv.writer(f)
            for r in range(sh.nrows):
                c.writerow(sh.row_values(r))
    os.remove(filename)
    try:
        messagebox.showinfo('Exportation réussie','Votre fichier a bien été chargé')
    except:
        messagebox.showinfo('Exportation échouée','Votre fichier n\'a pas pu se charger')
    SecondFen.destroy()

FirstFen = Tk() # initialisation de la première fenetre
FirstFen.title("Structure")

LabelStruct = Label(FirstFen, text='Structure :', font = 60)
LabelStruct.grid(row=0,column=0, ipady = 30, padx = 5)

chkValueStruct = BooleanVar()
chkValueStruct.set(False)

CheckStruct = Checkbutton(FirstFen, var=chkValueStruct, command=lambda:Is_checkStruct()) 
CheckStruct.grid(row=1, column=1, padx=4)

List = [] # initialise un tableau 
List.append("tout")

cur.execute("SELECT Recyclerie FROM Organisation") # récupère les recycleries insérées dans la base de données
OrgaList=cur.fetchall()

for row in OrgaList: # insère les recycleries dans le tableau 
    List.append(row[0])

ComboStruct = ttk.Combobox(FirstFen, values = List, width = 28, state='disabled') # combobox qui récupère les recycleries comme valeur
ComboStruct.set("Choississez la/les structure(s)") # valeur par défaut du combobox
ComboStruct.grid(row=1,column=2)

BtnAdd = Button(FirstFen, text='Ajouter',command=lambda:addStruct(), state='disabled') # bouton qui ajoute la recylerie sélectionné du combobox dans la listbox
BtnAdd.grid(row=1,column=3)

BtnDel = Button(FirstFen, text='Supprimer',command=lambda:delStruct(), state='disabled') # bouton qui supprime la recylerie sélectionné dans la listbox
BtnDel.grid(row=1,column=4,padx=5)

listStruct = Listbox(FirstFen, width=50) # listbox qui va contenir les recyleries à étudier
listStruct.grid(row=1, column=5)

LabelSecteur = Label(FirstFen, text='Secteur :', font = 60)
LabelSecteur.grid(row=2,column=0, ipady = 70, padx = 5)

LabelInsee = Label(FirstFen, text='Insee :')
LabelInsee.grid(row=3,column=1)

chkValue1 = BooleanVar() 
chkValue1.set(False)

chkValue2 = BooleanVar() 
chkValue2.set(False)

CheckOne = Checkbutton(FirstFen, text='1 code :', var=chkValue1, command=lambda:Is_check1()) # si on coche, active la ligne pour l'insertion d'1 code sinon reste désactivé 
CheckOne.grid(row=3, column=2, padx=4)

EntryInsee = Entry(FirstFen, state='disabled') #désactive tant que la checkbox n'est pas coché
EntryInsee.grid(row=3,column=3, padx = 10)

CheckMany = Checkbutton(FirstFen, text='plusieurs codes :', var=chkValue2, command=lambda:Is_check2()) # si on coche, active la ligne pour l'insertion de plusieurs codes sinon reste désactivé
CheckMany.grid(row=4, column=2, padx=4)

BtnInsee = Button(FirstFen, text='Importer votre fichier', command=lambda:FileInsee(), state='disabled') #désactive tant que la checkbox n'est pas coché
BtnInsee.grid(row=4,column=3)

LabelFile = Label(FirstFen, text= '')
LabelFile.grid(row=4, column=4)

listInsee = Listbox(FirstFen, height=10)
listInsee.grid(row=4, column=5, padx = 40)

BtnNext = Button(FirstFen, text='Suivant', command=lambda:new_window(listStruct)) # passe à la prochaine fenetre et prend en compte les données inscrites de la première fenetre 
BtnNext.grid(row=6,column=5, padx = 40, pady = 20)

FirstFen.mainloop()