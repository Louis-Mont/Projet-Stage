from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import sqlite3
import datetime
from xlwt import Workbook
import os
import pandas as pd

# connexion à la base de données sinon stop le programme et affiche l'erreur
try:
    #connect = sqlite3.connect("X:/finale.db") 
    connect = sqlite3.connect("finale.db")
    cur = connect.cursor()
except IOError:
    print(IOError)

def Is_checkStruct():
    '''
    fonction qui change l'etat des widgets en lien avec les structures
    '''
    if chkValueStruct.get() == True:
        ComboStruct.config(state='normal')
        BtnAdd.config(state='normal')
        BtnDel.config(state='normal')
    else:
        ComboStruct.config(state='disabled')
        BtnAdd.config(state='disabled')
        BtnDel.config(state='disabled')

def Is_check1():
    '''
    fonction qui change l'état des widgets en lien avec 1 code insee
    '''
    if chkInseeOne.get() == True:
        EntryInsee.config(state='normal')
        CheckMany.config(state='disabled')
    else:
        EntryInsee.config(state='disabled')
        CheckMany.config(state='normal')

def Is_check2():
    '''
    fonction qui change l'état des widgets en lien avec plusieurs codes insee
    '''
    if chkInseeMany.get() == True:
        BtnInsee.config(state='normal')
        CheckOne.config(state='disabled')
    else:
        BtnInsee.config(state='disabled')
        CheckOne.config(state='normal')
        listInsee.delete(0,END)
        LabelFile.config(text='')

def FileInsee():
    '''
    fonction pour chercher le fichier .txt contenant les plusieurs codes insee et l'insère dans la listbox
    '''
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
    '''
    fonction pour les requetes sql selon les choix de la premiere fenetre et retourne la liste des structures demandés
    '''
    ListStruct = []
    if chkInseeOne.get() == True and chkValueStruct.get() == False: # si l'utilisateur coche seulement la checkbox pour 1 code Insee
        val = EntryInsee.get()
        cur.execute("SELECT DISTINCT Recyclerie FROM Organisation, Insee, Commune WHERE Organisation.Id_Recyclerie = Commune.Id_Recyclerie AND Commune.Id_insee = Insee.Id_Insee AND Code = '%s' "%\
                            (val))
        StructList = cur.fetchall()
        ListStruct = [struct[0] for struct in StructList]
        
    elif chkInseeMany.get() == True and chkValueStruct.get() == False: # si l'utilisateur coche seulement la checkbox pour plusieurs codes Insee
        val = listInsee.get(0, END)
        for code in val:
            cur.execute("SELECT Recyclerie FROM Organisation, Insee, Commune WHERE Organisation.Id_Recyclerie = Commune.Id_Recyclerie AND Commune.Id_insee = Insee.Id_Insee AND Code = '%s'"%\
                            (code))
            StructList = cur.fetchall()
            for struct in StructList:
                ListStruct.append(struct[0])

    elif chkInseeOne.get() == True and chkValueStruct.get() == True: # si l'utilisateur coche la checkbox pour 1 code Insee et la checkbox pour choisir les structures
        val = EntryInsee.get()
        cur.execute("SELECT Recyclerie FROM Organisation, Insee, Commune WHERE Organisation.Id_Recyclerie = Commune.Id_Recyclerie AND Commune.Id_insee = Insee.Id_Insee AND Code = '%s'"%\
                            (val))
        StructList = cur.fetchall()
        ListStruct = [struct[0] for struct in StructList]
        print(ListStruct)

    elif chkInseeMany.get() == True and chkValueStruct.get() == True: # si l'utilisateur coche la checkbox pour pour plusieurs codes Insee et la checkbox pour choisir les structures
        print('kl')

    elif chkValueStruct.get() == True and chkInseeOne.get() == False and chkInseeMany.get() == False: # si l'utilisateur coche seulement la checkbox pour choisir les structures
        StructList = listStruct.get(0,END)
        ListStruct = [s for s in StructList]
 
    return list(set(ListStruct))

def ModaliteCollect(chkValue1, chkValue4, ComboModalite, LabelModalite, BtnAddModal, BtnDelModal, ListBoxModal):
    '''
    fonction qui affiche les widgets en lien avec les modalités si les checkbox des collectes sont cochés sinon les cache
    '''
    if chkValue1.get() == True:
        ComboModalite.grid(row=7,column=2)
        LabelModalite.grid(row=7,column=1)
        BtnAddModal.grid(row=7,column=3)
        BtnDelModal.grid(row=7,column=4)
        ListBoxModal.grid(row=7,column=5)
    elif chkValue4.get() == True:
        ComboModalite.grid(row=7,column=2)
        LabelModalite.grid(row=7,column=1)
        BtnAddModal.grid(row=7,column=3)
        BtnDelModal.grid(row=7,column=4)
        ListBoxModal.grid(row=7,column=5)
    else:
        ComboModalite.grid_forget()
        LabelModalite.grid_forget()
        BtnAddModal.grid_forget()
        BtnDelModal.grid_forget()
        ListBoxModal.grid_forget()

def new_window():
    '''
    fonction pour la deuxieme fenetre
    '''
    ListStruct = RequeteStruct()

    FirstFen.destroy() # Détruit la première fenêtre 
    SecondFen = Tk() # initialisation de la première fenêtre
    SecondFen.title("Options d'importation")

    #Partie catégorie :
    LabelCat = Label(SecondFen, text='Catégorie :', font = 60)
    LabelCat.grid(row=0,column=0, ipady = 30)

    ListCategorieBox = ["tout"] # initialise un tableau 
    cur.execute("SELECT Catégorie FROM Catégorie")
    Cat = cur.fetchall()
    for row in Cat:
        ListCategorieBox.append(row[0])
    
    Combo = ttk.Combobox(SecondFen, values = ListCategorieBox, width=29)
    Combo.set("Choississez la/les catégorie(s)")
    Combo.grid(row=0,column = 1)
    Combo.bind("<<ComboboxSelected>>", lambda e: Combo.get())

    BtnAdd2 = Button(SecondFen, text='Ajouter', command=lambda:addCat(listCat, Combo, ListCategorieBox))
    BtnAdd2.grid(row=0,column=2)

    BtnDel2 = Button(SecondFen, text='Supprimer', command=lambda:delCat(listCat))
    BtnDel2.grid(row=0,column=3,padx=5)

    listCat = Listbox(SecondFen, width=50)
    listCat.grid(row=0, column=5)

    #Partie temps :
    Labeltemps = Label(SecondFen, text='Temps :', font = 60)
    Labeltemps.grid(row=1,column=0, ipady = 10)

    cur.execute("SELECT Min(date) FROM Arrivage")
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

    LabelModalite = Label(SecondFen, text='Modalité(s) :')
    ListModaliteCombo = ['tout']
    cur.execute('SELECT DISTINCT origine FROM Arrivage')
    Modal = cur.fetchall()
    for row in Modal:
        if row[0] != '0':
            ListModaliteCombo.append(row[0]) 
    ComboModalite = ttk.Combobox(SecondFen, values = ListModaliteCombo, width=25)
    ComboModalite.set("Choississez la/les modalité(s)")
    BtnAddModal = Button(SecondFen, text='ajouter',command=lambda:addModal(ListBoxModal, ComboModalite, ListModaliteCombo))
    BtnDelModal = Button(SecondFen, text='supprimer',command=lambda:delModal(ListBoxModal))
    ListBoxModal = Listbox(SecondFen, width=30)

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
    Chkbox1 = Checkbutton(SecondFen, text = 'Poids collecté (en Kg)', var = chkValue1, command=lambda:ModaliteCollect(chkValue1, chkValue4, ComboModalite, LabelModalite, BtnAddModal, BtnDelModal, ListBoxModal))
    Chkbox1.grid(row=5,column=1)
    Chkbox2 = Checkbutton(SecondFen, text = 'Poids vendu (en Kg)', var = chkValue2)
    Chkbox2.grid(row=5,column=2)
    Chkbox3 = Checkbutton(SecondFen, text = 'Chiffre d\'affaire (en €)', var = chkValue3)
    Chkbox3.grid(row=5,column=3)
    Chkbox4 = Checkbutton(SecondFen, text = 'Nombre d\'objet collecté', var = chkValue4, command=lambda:ModaliteCollect(chkValue1,chkValue4, ComboModalite, LabelModalite, BtnAddModal, BtnDelModal, ListBoxModal))
    Chkbox4.grid(row=6,column=1)
    Chkbox5 = Checkbutton(SecondFen, text = 'Nombre d\'objet vendu', var = chkValue5)
    Chkbox5.grid(row=6,column=2)

    BtnExport = Button(SecondFen, text='Exporter', command=lambda:export(SecondFen, listCat, choixjour1, choixmois1, choixan1, choixjour2, choixmois2, choixan2, ListStruct, chkValue1, chkValue2, chkValue3, chkValue4, chkValue5, ListBoxModal))
    BtnExport.grid(row=10,column=5, padx = 40, pady = 20)

    SecondFen.mainloop()

def addModal(ListBoxModal, ComboModalite, ListModaliteCombo):
    '''
    fonction du bouton ajouter pour des modalités
    '''
    listFiles = ListBoxModal.get(0,END)
    value = ComboModalite.get()
    if ComboModalite.get() == 'tout':
        for modal in ListModaliteCombo:
            if modal not in listFiles:
                ListBoxModal.insert(END, modal)
        idx = ListBoxModal.get(0, END).index('tout')
        ListBoxModal.delete(idx)

    if value not in listFiles and value != "Choississez la/les modalité(s)" and value != 'tout':
        ListBoxModal.insert(END, value)

def delModal(ListBoxModal):
    '''
    fonction pour le bouton supprimer des modalités

    Arguments:
        listBoxModal {listbox} -- listbox contenant les modalités sélectionnées
    '''
    itemSelected = ListBoxModal.curselection()
    ListBoxModal.delete(itemSelected[0])

def addCat(listCat, Combo, ListCatBox):
    '''
    fonction pour le bouton ajouter des catégories

    Arguments:
        listCat {Listbox} -- listbox contenant les catégories sélectionnées
        Combo {Combobox} -- combobox contenant la liste des catégories de la base sql
        ListCatBox {list[str]} -- liste contenant les catégories de la base sql
    '''
    listFiles = listCat.get(0,END)
    value = Combo.get()
    if Combo.get() == 'tout':
        for cat in ListCatBox:
            if cat not in listFiles:
                listCat.insert(END, cat)
        idx = listCat.get(0, END).index('tout')
        listCat.delete(idx)

    if value not in listFiles and value != 'Choississez la/les catégorie(s)' and value != 'tout':
        listCat.insert(END, value)

def delCat(listCat):
    '''
    fonction pour le bouton supprimer des catégories

    Arguments:
        listCat {listbox} -- listbox contenant les catégories sélectionnées
    '''
    itemSelected = listCat.curselection()
    listCat.delete(itemSelected[0])

def addStruct():
    '''
    fonction pour le bouton ajouter des structures
    '''
    listFiles = listStruct.get(0,END)
    value = ComboStruct.get()
    if ComboStruct.get() == 'tout':
        for struct in ListRecyclerieBox:
            if struct not in listFiles:
                listStruct.insert(END, struct)
        idx = listStruct.get(0, END).index('tout')
        listStruct.delete(idx)
                    
    if value not in listFiles and value != 'tout' and value != 'Choississez la/les structure(s)':
        listStruct.insert(END, value)

def delStruct():
    '''
    fonction pour le bouton supprimer des structures
    '''
    itemSelected = listStruct.curselection()
    listStruct.delete(itemSelected[0])

def export(SecondFen, listCat, jour1, mois1, an1, jour2, mois2, an2, StructList, chk1, chk2, chk3, chk4, chk5, ListBoxModal):
    '''
    fonction d'export du csv contenant les données demandées

    Arguments:
        SecondFen {Tk} -- Deuxième fenetre
        listCat {listbox} -- listbox contenant les catégories choisis
        StructList {list[str]} -- liste contenant les structures choisis
        ListBoxModal {listbox} -- listbox contenant les modalités choisis
    '''
    filename = filedialog.asksaveasfilename(defaultextension = '.xls',
                                            filetypes = [("xls files", '*.xls')])
    
    sizeCat = listCat.size()

    ListCategorie = [c for c in listCat.get(0,END)] # création d'une liste des catégories sélectionnées
    ListModal = [m for m in ListBoxModal.get(0,END)] # création d'une liste des modales sélectionnées

    jourfirst = jour1.get()
    moisfirst = mois1.get()
    anfirst = an1.get()
    joursec = jour2.get()
    moissec = mois2.get()
    ansec = an2.get()
    dateFirst =  anfirst + "/" + moisfirst + "/" + jourfirst
    dateEnd = ansec + "/" + moissec + "/" + joursec
    classeurexport=Workbook()
    pageexport=classeurexport.add_sheet("EXPORT", cell_overwrite_ok=True)
    if sizeCat == 0: # si la liste des catégories sélectionnées est vide alors affiche les données totaux
        pageexport.write(0,0,"Structure(s)")
        TailleCat = len(ListCategorie)
        TailleStruct = len(StructList)
        TailleModal = len(ListModal)
        y = 1
        if chk1.get() == True and TailleModal != 0:      
            pageexport.write(0,y,"Modalité(s)")
            pageexport.write(0,y+1,"Poids collecté (en kg)")
            indice = 0           
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                indice2 = 0
                while (indice2 != TailleModal):
                    cur.execute("SELECT sum(Poids)*nombre FROM Produit, Arrivage, Organisation WHERE Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND origine = '%s'"%\
                                        (StructList[indice], str(dateFirst),str(dateEnd), ListModal[indice2]))
                    PoidsCollect = cur.fetchone() [0]
                    pageexport.write(x,1,ListModal[indice2])
                    pageexport.write(x,0,StructList[indice])
                    if PoidsCollect == None:
                        pageexport.write(x, y+1, 'NR')
                    else:
                        pageexport.write(x, y+1, round(PoidsCollect))
                    indice2+=1
                    x+=1
                indice+=1
            y+=1 # on passe a la colonne suivante pour les données suivantes
        elif chk1.get() == True and TailleModal == 0:
            pageexport.write(0,y,"Poids collecté (en kg)")
            indice = 0           
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                cur.execute("SELECT sum(Poids)*nombre FROM Produit, Arrivage, Organisation WHERE Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s'"%\
                                    (StructList[indice], str(dateFirst),str(dateEnd)))
                PoidsCollect = cur.fetchone() [0]
                pageexport.write(x,0,StructList[indice])
                if PoidsCollect == None:
                    pageexport.write(x, y, 'NR')
                else:
                    pageexport.write(x, y, round(PoidsCollect))
                x+=1
                indice+=1
            y+=1 # on passe a la colonne suivante pour les données suivantes
        if chk2.get() == True:
            pageexport.write(0,y,"Poids vendu (en kg)")
            indice = 0
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                cur.execute("SELECT sum(Lignes_vente.Poids) FROM Lignes_vente, Vente, Organisation WHERE Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_vente = Vente.Id_Vente AND date > '%s' AND date < '%s'"%\
                                    (StructList[indice], str(dateFirst),str(dateEnd)))
                PoidsVendu = cur.fetchone() [0]
                pageexport.write(x,0,StructList[indice])
                if PoidsVendu == None:
                    pageexport.write(x, y, 'NR')
                else:
                    pageexport.write(x, y, round(PoidsVendu))
                indice+=1
                x+=1
            y+=1
        if chk3.get() == True:
            pageexport.write(0,y,"Chiffre d'affaire (en €)")
            indice = 0
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                cur.execute("SELECT sum(Montant) FROM Vente, Lignes_vente, Organisation WHERE Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_vente=Vente.Id_Vente AND date > '%s' AND date < '%s'"%\
                                    (StructList[indice], str(dateFirst),str(dateEnd)))
                Chiffre = cur.fetchone() [0]
                pageexport.write(x,0,StructList[indice])
                if Chiffre == None:
                    pageexport.write(x, y, 'NR')
                else:
                    pageexport.write(x, y, round(Chiffre))
                indice+=1
                x+=1
            y+=1
        if chk4.get() == True and TailleModal != 0:
            if chk1.get() == False:
                pageexport.write(0,y,"Modalité(s)")
            pageexport.write(0,y+1,"Nombre d'objet collecté")
            indice = 0
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                indice2 = 0
                while (indice2 != TailleModal):
                    cur.execute("SELECT count(Id_Produit)*nombre FROM Produit, Arrivage, Organisation WHERE Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND origine = '%s'"%\
                                        (StructList[indice], str(dateFirst),str(dateEnd), ListModal[indice2]))
                    NbrCollect = cur.fetchone() [0]
                    pageexport.write(x,1,ListModal[indice2])
                    pageexport.write(x,0,StructList[indice])
                    if NbrCollect == None:
                        pageexport.write(x, y+1, 'NR')
                    else:
                        pageexport.write(x, y+1, round(NbrCollect))
                    indice2+=1
                    x+=1
                indice+=1
            y+=1
        elif chk4.get() == True and TailleModal == 0:
            pageexport.write(0,y,"Nombre d'objet collecté")
            indice = 0
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                cur.execute("SELECT count(Id_Produit)*nombre FROM Produit, Arrivage, Organisation WHERE Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s'"%\
                                        (StructList[indice], str(dateFirst),str(dateEnd)))
                NbrCollect = cur.fetchone() [0]
                pageexport.write(x,0,StructList[indice])
                if NbrCollect == None:
                    pageexport.write(x, y, 'NR')
                else:
                    pageexport.write(x, y, round(NbrCollect))
                x+=1
                indice+=1
            y+=1
        if chk5.get() == True:
            pageexport.write(0,y,"Nombre d'objet vendu")
            indice = 0
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                cur.execute("SELECT count(Id_ligne_vente) FROM Lignes_vente, Vente, Organisation WHERE Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Vente.Id_Vente=Lignes_vente.Id_vente AND date > '%s' AND date < '%s'"%\
                                    (StructList[indice], str(dateFirst),str(dateEnd)))
                NbrVente = cur.fetchone() [0]
                pageexport.write(x,0,StructList[indice])
                if NbrVente == None:
                    pageexport.write(x, y, 'NR')
                else:
                    pageexport.write(x, y, round(NbrVente))
                indice+=1
                x+=1
            y+=1
    else:
        x = 1
        pageexport.write(0,0,"Structure(s)")
        pageexport.write(0,1,"Catégorie(s)")
        TailleCat = len(ListCategorie)
        TailleStruct = len(StructList)
        TailleModal = len(ListModal)
        y = 2
        '''
        if chk1.get() == True:
            indice = 0 # variable pour l'indexation selon la taille de la liste des catégories
            pageexport.write(0,y,"Poids collecté (en kg)")
            x = 1 # indice pour la ligne dans le fichier xls
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
                        
                    indice2+=1
                if indice2 == TailleStruct:
                        if val != 0:
                            pageexport.write(x, y, round(val))
                            x +=1
                        else:
                            pageexport.write(x, y, 'NR')
                            x +=1
                indice+=1
            y+=1 # on passe a la colonne suivante pour les données suivantes
            '''
        if chk1.get() == True and TailleModal != 0:
            indice = 0 # variable pour l'indexation selon la taille de la liste des catégories
            pageexport.write(0,y,"Modalité(s)")
            pageexport.write(0,y+1,"Poids collecté (en kg)")
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                indice2 = 0
                while (indice2 != TailleCat):
                    indice3 = 0 # variable pour l'indexation selon la taille de la liste des modalités
                    while (indice3 != TailleModal):
                        cur.execute("SELECT sum(Poids)*nombre FROM Produit, Catégorie, Arrivage, Organisation WHERE Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_catégorie = Catégorie.Id_catégorie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND Catégorie = '%s' AND origine = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                        (StructList[indice], str(dateFirst),str(dateEnd),ListCategorie[indice2], ListModal[indice3]))
                        PoidsCollect = cur.fetchall()
                        pageexport.write(x, 0, StructList[indice])
                        pageexport.write(x, 2, ListModal[indice3])
                        pageexport.write(x, 1, ListCategorie[indice2])
                        if len(PoidsCollect) == 0:
                            pageexport.write(x, y+1, 'NR')
                            x+=1
                        else:
                            for poids in PoidsCollect:
                                pageexport.write(x, y+1, round(poids[0]))
                                x +=1 
                        indice3+=1 
                    indice2+=1                                      
                indice+=1
            y+=1 # on passe a la colonne suivante pour les données suivantes
        elif chk1.get() == True and TailleModal == 0:
            indice = 0 # variable pour l'indexation selon la taille de la liste des catégories
            pageexport.write(0,y,"Poids collecté (en kg)")
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                indice2 = 0
                while (indice2 != TailleCat):
                    cur.execute("SELECT sum(Poids)*nombre FROM Produit, Catégorie, Arrivage, Organisation WHERE Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_catégorie = Catégorie.Id_catégorie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                        (StructList[indice], str(dateFirst),str(dateEnd),ListCategorie[indice2]))
                    PoidsCollect = cur.fetchall()
                    pageexport.write(x, 0, StructList[indice])
                    pageexport.write(x, 1, ListCategorie[indice2])
                    if len(PoidsCollect) == 0:
                        pageexport.write(x, y, 'NR')
                        x+=1
                    else:
                        for poids in PoidsCollect:
                            pageexport.write(x, y, round(poids[0]))
                            x +=1 
                    indice2+=1                                      
                indice+=1
            y+=1 # on passe a la colonne suivante pour les données suivantes
        if chk2.get() == True:
            indice = 0 # variable pour l'indexation selon la taille de la liste des catégories
            pageexport.write(0,y,"Poids vendu (en kg)")
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                indice2 = 0 # variable pour l'indexation selon la taille de la liste des structures
                while (indice2 != TailleCat):
                    cur.execute("SELECT sum(Lignes_vente.Poids) FROM Lignes_vente, Vente, Catégorie, Organisation WHERE Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_vente = Vente.Id_Vente AND Lignes_vente.Id_catégorie=Catégorie.Id_catégorie AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                    (StructList[indice], str(dateFirst),str(dateEnd),ListCategorie[indice2]))
                    PoidsVente = cur.fetchall()
                    pageexport.write(x, 0, StructList[indice])
                    pageexport.write(x, 1, ListCategorie[indice2])
                    if len(PoidsVente) == 0:
                        pageexport.write(x, y, 'NR')
                        x+=1
                    else:
                        for poids in PoidsVente:
                            pageexport.write(x, y, round(poids[0]))
                            x +=1  
                    indice2+=1                                      
                indice+=1
            y+=1 # on passe a la colonne suivante pour les données suivantes
        if chk3.get() == True:
            indice = 0 # variable pour l'indexation selon la taille de la liste des catégories
            pageexport.write(0,y,"Chiffre d'affaire (en €)")
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                indice2 = 0 # variable pour l'indexation selon la taille de la liste des structures
                while (indice2 != TailleCat):
                    cur.execute("SELECT sum(Montant), Catégorie FROM Vente, Catégorie, Lignes_vente, Organisation WHERE Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_vente=Vente.Id_Vente AND Lignes_vente.Id_catégorie=Catégorie.Id_catégorie AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                    (StructList[indice], str(dateFirst),str(dateEnd),ListCategorie[indice2]))
                    Chiffre = cur.fetchall()
                    pageexport.write(x, 0, StructList[indice])
                    pageexport.write(x, 1, ListCategorie[indice2])
                    if len(Chiffre) == 0:
                        pageexport.write(x, y, 'NR')
                        x+=1
                    else:
                        for c in Chiffre:
                            pageexport.write(x, y, round(c[0]))
                            x +=1  
                    indice2+=1                                      
                indice+=1
            y+=1 # on passe a la colonne suivante pour les données suivantes
        if chk4.get() == True and TailleModal != 0:
            indice = 0 # variable pour l'indexation selon la taille de la liste des catégories
            if chk1.get() == False:
                pageexport.write(0,y,"Modalité(s)")
            pageexport.write(0,y+1,"Nombre d'objet collecté")
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                indice2 = 0 # variable pour l'indexation selon la taille de la liste des structures
                while (indice2 != TailleCat):
                    indice3 = 0 # variable pour l'indexation selon la taille de la liste des modalités
                    while (indice3 != TailleModal):
                        cur.execute("SELECT count(Id_Produit)*nombre FROM Produit, Arrivage, Catégorie, Organisation WHERE Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_catégorie=Catégorie.Id_catégorie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND Catégorie = '%s' AND origine = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                        (StructList[indice], str(dateFirst),str(dateEnd),ListCategorie[indice2], ListModal[indice3]))
                        NbrCollect = cur.fetchall()
                        pageexport.write(x, 0, StructList[indice])
                        pageexport.write(x, 2, ListModal[indice3])
                        pageexport.write(x, 1, ListCategorie[indice2])
                        if len(NbrCollect) == 0:
                            pageexport.write(x, y+1, 'NR')
                            x+=1
                        else:
                            for nbr in NbrCollect:
                                pageexport.write(x, y+1, round(nbr[0]))
                                x +=1
                        indice3+=1
                    indice2+=1                                      
                indice+=1
            y+=1 # on passe a la colonne suivante pour les données suivantes
        elif chk4.get() == True and TailleModal == 0:
            indice = 0 # variable pour l'indexation selon la taille de la liste des catégories
            pageexport.write(0,y,"Nombre d'objet collecté")
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                indice2 = 0 # variable pour l'indexation selon la taille de la liste des structures
                while (indice2 != TailleCat):
                    cur.execute("SELECT count(Id_Produit)*nombre FROM Produit, Arrivage, Catégorie, Organisation WHERE Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_catégorie=Catégorie.Id_catégorie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                    (StructList[indice], str(dateFirst),str(dateEnd),ListCategorie[indice2]))
                    NbrCollect = cur.fetchall()
                    pageexport.write(x, 0, StructList[indice])
                    pageexport.write(x, 1, ListCategorie[indice2])
                    if len(NbrCollect) == 0:
                        pageexport.write(x, y, 'NR')
                        x+=1
                    else:
                        for nbr in NbrCollect:
                            pageexport.write(x, y, round(nbr[0]))
                            x +=1
                    indice2+=1                                      
                indice+=1
            y+=1 # on passe a la colonne suivante pour les données suivantes
        if chk5.get() == True:
            indice = 0 # variable pour l'indexation selon la taille de la liste des catégories
            pageexport.write(0,y,"Nombre d'objet vendu")
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleStruct):
                indice2 = 0 # variable pour l'indexation selon la taille de la liste des structures
                while (indice2 != TailleCat):
                    cur.execute("SELECT count(Id_ligne_vente) FROM Lignes_vente, Vente, Catégorie, Organisation WHERE Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_catégorie=Catégorie.Id_catégorie AND Vente.Id_Vente=Lignes_vente.Id_vente AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie"%\
                                    (StructList[indice], str(dateFirst),str(dateEnd),ListCategorie[indice2]))
                    NbrVente = cur.fetchall()
                    pageexport.write(x, 0, StructList[indice])
                    pageexport.write(x, 1, ListCategorie[indice2])
                    if len(NbrVente) == 0:
                        pageexport.write(x, y, 'NR')
                        x+=1
                    else:
                        for nbr in NbrVente:
                            pageexport.write(x, y, round(nbr[0]))
                            x +=1  
                    indice2+=1                                      
                indice+=1
            y+=1 # on passe a la colonne suivante pour les données suivantes
    try: # sauvegarde le fichier et le convertit en csv 
        classeurexport.save(filename)
        csv_file = os.path.splitext(filename)[0] + '.csv'
        read_file = pd.read_excel(filename)
        read_file.to_csv(csv_file, index = None, header=True)
        os.remove(filename)
        messagebox.showinfo('Exportation réussie','Votre fichier a bien été chargé')
    except IOError:
        messagebox.showerror('Exportation échouée','Votre fichier n\'a pas pu se charger')
        print(IOError)
    SecondFen.destroy() # quitte l'application

FirstFen = Tk() # initialisation de la première fenetre
FirstFen.title("Structure")

LabelStruct = Label(FirstFen, text='Structure :', font = 60)
LabelStruct.grid(row=0,column=0, ipady = 30, padx = 5)

chkValueStruct = BooleanVar()
chkValueStruct.set(False)

CheckStruct = Checkbutton(FirstFen, var=chkValueStruct, command=lambda:Is_checkStruct()) 
CheckStruct.grid(row=1, column=1, padx=4)

ListRecyclerieBox = ["tout"] # initialise la liste de la combobox des structures 

cur.execute("SELECT Recyclerie FROM Organisation") # récupère les structures insérées dans la base de données
OrgaList=cur.fetchall()

for row in OrgaList: # insère les recycleries dans la liste
    ListRecyclerieBox.append(row[0])

ComboStruct = ttk.Combobox(FirstFen, values = ListRecyclerieBox, width = 28, state='disabled') # combobox qui récupère les recycleries comme valeur
ComboStruct.set("Choississez la/les structure(s)") # valeur par défaut du combobox
ComboStruct.grid(row=1,column=2)

BtnAdd = Button(FirstFen, text='Ajouter',command=lambda:addStruct(), state='disabled') # bouton qui ajoute la recylerie sélectionné du combobox dans la listbox
BtnAdd.grid(row=1,column=3)

BtnDel = Button(FirstFen, text='Supprimer',command=lambda:delStruct(), state='disabled') # bouton qui supprime la recylerie sélectionnée dans la listbox
BtnDel.grid(row=1,column=4,padx=5)

listStruct = Listbox(FirstFen, width=50) # listbox qui va contenir les recyleries à étudier
listStruct.grid(row=1, column=5)

LabelSecteur = Label(FirstFen, text='Secteur :', font = 60)
LabelSecteur.grid(row=2,column=0, ipady = 70, padx = 5)

LabelInsee = Label(FirstFen, text='Insee :')
LabelInsee.grid(row=3,column=1)

chkInseeOne = BooleanVar()
chkInseeOne.set(False)

chkInseeMany = BooleanVar()
chkInseeMany.set(False)

CheckOne = Checkbutton(FirstFen, text='1 code :', var=chkInseeOne, command=lambda:Is_check1()) # si on coche, active la ligne pour l'insertion d'1 code sinon reste désactivé 
CheckOne.grid(row=3, column=2, padx=4)

EntryInsee = Entry(FirstFen, state='disabled') #désactivé tant que la checkbox n'est pas coché
EntryInsee.grid(row=3,column=3, padx = 10)

CheckMany = Checkbutton(FirstFen, text='plusieurs codes :', var=chkInseeMany, command=lambda:Is_check2()) # si on coche, active la ligne pour l'insertion de plusieurs codes sinon reste désactivé
CheckMany.grid(row=4, column=2, padx=4)

BtnInsee = Button(FirstFen, text='Importer votre fichier', command=lambda:FileInsee(), state='disabled') #désactivé tant que la checkbox n'est pas coché
BtnInsee.grid(row=4,column=3)

LabelFile = Label(FirstFen, text= '')
LabelFile.grid(row=4, column=4)

listInsee = Listbox(FirstFen, height=10)
listInsee.grid(row=4, column=5, padx = 40)

BtnNext = Button(FirstFen, text='Suivant', command=lambda:new_window()) # passe à la prochaine fenetre et prend en compte les données inscrites de la première fenetre 
BtnNext.grid(row=6,column=5, padx = 40, pady = 20)

FirstFen.mainloop()