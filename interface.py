# from io import StringIO
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
cur = None
try:
    connect = sqlite3.connect("finale.db")
    cur = connect.cursor()
except IOError:
    print(IOError)


def is_check_struct():
    if chkValueStruct.get():
        ComboStruct.config(state='normal')
        BtnAdd.config(state='normal')
        BtnDel.config(state='normal')
    else:
        ComboStruct.config(state='disabled')
        BtnAdd.config(state='disabled')
        BtnDel.config(state='disabled')


def is_check1():
    if chkValue1.get():
        EntryInsee.config(state='normal')
        CheckMany.config(state='disabled')
    else:
        EntryInsee.config(state='disabled')
        CheckMany.config(state='normal')


def is_check2():
    if chkValue2.get():
        BtnInsee.config(state='normal')
        CheckOne.config(state='disabled')
    else:
        BtnInsee.config(state='disabled')
        CheckOne.config(state='normal')
        listInsee.delete(0, END)
        LabelFile.config(text='')


def file_insee():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Sélectionnez votre fichier",
                                          filetypes=(("Fichier texte", "*.txt"),
                                                     ("Fichier pdf", "*.pdf")))

    LabelFile.config(text=filename)
    fichier = open(str(filename), "r")
    lines = fichier.readlines()
    fichier.close()
    listInsee.delete(0, END)
    i = 1
    for line in lines:
        listInsee.insert(i, line.strip())
        i += 1


def requete_struct():
    if chkValue1.get() and not chkValueStruct.get():
        val = EntryInsee.get()
        cur.execute(
            f"SELECT Recyclerie FROM Organisation, Insee, Commune WHERE Organisation.Id_Recyclerie = Commune.Id_Recyclerie AND Commune.Id_insee = Insee.Id_Insee AND Code = '{val}'")
        struct_list = cur.fetchall()
        for struct in struct_list:
            nom_struct = struct[0]
    elif chkValue2.get() and not chkValueStruct.get():
        val = listInsee.get(0, END)
        for code in val:
            print(code)
            cur.execute(
                f"SELECT Recyclerie FROM Organisation, Insee, Commune WHERE Organisation.Id_Recyclerie = Commune.Id_Recyclerie AND Commune.Id_insee = Insee.Id_Insee AND Code = '{code}'")
            print(cur.fetchall())
    elif chkValue1.get() and chkValueStruct.get():
        print("rt")
    elif chkValue2.get() and chkValueStruct.get():
        print('kl')


def new_window(list_struct):
    requete_struct()

    struct_list = list_struct.get(0, END)

    FirstFen.destroy()  # Détruit la première fenêtre
    second_fen = Tk()  # initialisation de la première fenêtre
    second_fen.title("Options d'importation")

    # Partie catégorie :
    label_cat = Label(second_fen, text='Catégorie :', font=60)
    label_cat.grid(row=0, column=0, ipady=30)

    id = 1
    cur.execute(
        f"SELECT Catégorie FROM Catégorie, Produit WHERE Catégorie.Id_catégorie = Produit.Id_catégorie and Id_recyclerie = '{id}' GROUP BY Produit.Id_catégorie HAVING count(Produit.Id_catégorie) != 0")
    cat = cur.fetchall()
    List = []
    for row in cat:
        List.append(row[0])

    combo = ttk.Combobox(second_fen, values=List, width=29)
    combo.set("Choississez la/les catégorie(s)")
    combo.grid(row=0, column=1)

    combo.bind("<<ComboboxSelected>>", lambda e: combo.get())

    btn_add2 = Button(second_fen, text='Ajouter', command=lambda: add_cat(list_cat, combo))
    btn_add2.grid(row=0, column=2)

    btn_del2 = Button(second_fen, text='Supprimer', command=lambda: del_cat(list_cat))
    btn_del2.grid(row=0, column=3, padx=5)

    list_cat = Listbox(second_fen, width=50)
    list_cat.grid(row=0, column=5)

    # Partie temps :
    label_temps = Label(second_fen, text='Temps :', font=60)
    label_temps.grid(row=1, column=0, ipady=10)

    cur.execute(f"SELECT Min(date) FROM Arrivage WHERE Id_recyclerie = '{id}'")
    date_bdd = cur.fetchone()[0]
    an_bdd = date_bdd[:4]
    an = datetime.date.today().year
    an = int(an)
    dates = lambda l, u: [str(i).zfill(2) for i in range(l, u)]
    jours = dates(1, 32)
    mois = dates(1, 13)
    ans = [i for i in range(int(an_bdd), an + 1)]
    """jours = [str(i).zfill(2) for i in range(1, 32)]
    mois = [str(i).zfill(2) for i in range(1, 13)]"""
    mess_debut = Label(second_fen, text="choisissez la date de début :")
    mess_debut.grid(row=2, column=1, padx=10)
    choixjour1 = ttk.Combobox(second_fen, values=jours)
    choixjour1.grid(row=2, column=2)
    choixjour1.set('01')
    choixmois1 = ttk.Combobox(second_fen, values=mois)
    choixmois1.grid(row=2, column=3)
    choixmois1.set('01')
    choixan1 = ttk.Combobox(second_fen, values=ans)
    choixan1.grid(row=2, column=4)
    choixan1.set(int(an_bdd))
    mess_fin = Label(second_fen, text="choisissez la date de fin :")
    mess_fin.grid(row=3, column=1, pady=10)
    choixjour2 = ttk.Combobox(second_fen, values=jours)
    choixjour2.grid(row=3, column=2)
    choixjour2.set('31')
    choixmois2 = ttk.Combobox(second_fen, values=mois)
    choixmois2.grid(row=3, column=3)
    choixmois2.set('12')
    choixan2 = ttk.Combobox(second_fen, values=ans)
    choixan2.grid(row=3, column=4)
    choixan2.set(an)

    # Partie quantitative :
    label_qte = Label(second_fen, text='Quantitative :', font=60)
    label_qte.grid(row=4, column=0, ipady=40, padx=5)

    chk_value1 = BooleanVar()
    chk_value1.set(False)
    chk_value2 = BooleanVar()
    chk_value2.set(False)
    chk_value3 = BooleanVar()
    chk_value3.set(False)
    chk_value4 = BooleanVar()
    chk_value4.set(False)
    chk_value5 = BooleanVar()
    chk_value5.set(False)
    chkbox1 = Checkbutton(second_fen, text='Poids collecté (en Kg)', var=chk_value1)
    chkbox1.grid(row=5, column=1)
    chkbox2 = Checkbutton(second_fen, text='Poids vendu (en Kg)', var=chk_value2)
    chkbox2.grid(row=5, column=2)
    chkbox3 = Checkbutton(second_fen, text='Chiffre d\'affaire (en €)', var=chk_value3)
    chkbox3.grid(row=5, column=3)
    chkbox4 = Checkbutton(second_fen, text='Nombre d\'objet collecté', var=chk_value4)
    chkbox4.grid(row=6, column=1)
    chkbox5 = Checkbutton(second_fen, text='Nombre d\'objet vendu', var=chk_value5)
    chkbox5.grid(row=6, column=2)

    btn_export = Button(second_fen, text='Exporter',
                        command=lambda: export(second_fen, list_cat, choixjour1, choixmois1, choixan1, choixjour2,
                                               choixmois2, choixan2, struct_list, chk_value1, chk_value2, chk_value3,
                                               chk_value4, chk_value5))
    btn_export.grid(row=10, column=5, padx=40, pady=20)

    second_fen.mainloop()


def add_cat(list_cat, combo):
    list_files = list_cat.get(0, END)
    value = combo.get()
    if value not in list_files and value != 'Choississez la/les catégorie(s)':
        list_cat.insert(END, value)


def del_cat(list_cat):
    item_selected = list_cat.curselection()
    list_cat.delete(item_selected[0])


def add_struct():
    listFiles = listStruct.get(0, END)
    value = ComboStruct.get()
    if ComboStruct.get() == 'tout':
        for struct in List:
            listStruct.insert(END, struct)
        listStruct.delete(0)
    if value not in listFiles and value != 'tout' and value != 'Choississez la/les structure(s)':
        listStruct.insert(END, value)


def del_struct():
    item_selected = listStruct.curselection()
    listStruct.delete(item_selected[0])


def export(second_fen, list_cat, jour1, mois1, an1, jour2, mois2, an2, struct_list, chk1, chk2, chk3, chk4, chk5):
    filename = filedialog.asksaveasfilename(defaultextension='.xls',
                                            filetypes=[("xls files", '*.xls')])

    size = list_cat.size()

    list_category = [c for c in list_cat.get(0, END)]  # création d'une liste des catégories sélectionnées

    jour_first = jour1.get()
    mois_first = mois1.get()
    an_first = an1.get()
    jour_sec = jour2.get()
    mois_sec = mois2.get()
    an_sec = an2.get()
    date_first = an_first + "/" + mois_first + "/" + jour_first
    date_end = an_sec + "/" + mois_sec + "/" + jour_sec
    classeur_export = Workbook()
    page_export = classeur_export.add_sheet("EXPORT")
    page_export.write(0, 0, "Données du : " + date_first + " au " + date_end)
    page_export.write(2, 0, "Les recycleries : ")
    y = 1
    for struct in struct_list:
        page_export.write(2, y, struct)
        y += 1
    page_export.write(4, 0, "Catégories")
    if size == 0:  # si la liste est vide insert toutes les catégories ayant au minimum un arrivage dans le fichier
        i = 5
        cur.execute("SELECT Catégorie FROM Catégorie GROUP BY Id_Catégorie")
        cat = cur.fetchall()
        list_category = [val[0] for val in cat]
        for val in cat:
            page_export.write(i, 0, val[0])
            i += 1
        page_export.write(i + 1, 0, "total :")
    else:
        i = 5
        for c in list_category:
            page_export.write(i, 0, c)
            i += 1
        page_export.write(i + 1, 0, "total :")
    taille_cat = len(list_category)
    taille_struct = len(struct_list)
    y = 1
    if chk1.get():
        total_poids_c = 0  # variable pour le poids collecté sur l'ensemble des catégories
        indice = 0  # variable pour l'indexation selon la taille de la liste des catégories
        page_export.write(4, y, "Poids collecté (en kg)")
        x = 5  # indice pour la ligne dans le fichier xls
        while indice != taille_cat:
            indice2 = 0  # variable pour l'indexation selon la taille de la liste des structures
            val = 0  # variable qui va recevoir les valeurs de la requête
            while indice2 != taille_struct:
                cur.execute(
                    "SELECT sum(Poids)*nombre FROM Produit, Catégorie, Arrivage, Organisation WHERE Recyclerie = '%s' "
                    "AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_catégorie = "
                    "Catégorie.Id_catégorie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < "
                    "'%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie" %
                    (struct_list[indice2], str(date_first), str(date_end), list_category[indice]))
                poids_collect = cur.fetchall()
                if len(poids_collect) == 0:
                    val += 0
                else:
                    for poids in poids_collect:
                        val += poids[0]
                        total_poids_c += poids[0]
                indice2 += 1
            if indice2 == taille_struct:
                if val != 0:
                    page_export.write(x, y, val)
                    x += 1
                else:
                    page_export.write(x, y, 'NR')
                    x += 1
            indice += 1
        page_export.write(x + 1, y, total_poids_c)
        y += 1  # on passe a la colonne suivante pour les données suivantes
    if chk2.get():
        total_poids_v = 0
        indice = 0
        page_export.write(4, y, "Poids vendu (en kg)")
        x = 5  # indice pour la ligne dans le fichier xls
        while indice != taille_cat:
            indice2 = 0  # variable pour l'indexation selon la taille de la liste des structures
            val = 0  # variable qui va recevoir les valeurs de la requete
            while indice2 != taille_struct:
                cur.execute(
                    "SELECT sum(Lignes_vente.Poids) FROM Lignes_vente, Vente, Catégorie, Organisation WHERE "
                    "Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_vente = "
                    "Vente.Id_Vente AND Lignes_vente.Id_catégorie=Catégorie.Id_catégorie AND date > '%s' AND date < "
                    "'%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie" % \
                    (struct_list[indice2], str(date_first), str(date_end), list_category[indice]))
                poids_vendu = cur.fetchall()
                if len(poids_vendu) == 0:
                    val += 0
                else:
                    for poids in poids_vendu:
                        val += poids[0]
                        total_poids_v += poids[0]
                indice2 += 1
            if indice2 == taille_struct:
                if val != 0:
                    page_export.write(x, y, val)
                    x += 1
                else:
                    page_export.write(x, y, 'NR')
                    x += 1
            indice += 1
        page_export.write(x + 1, y, total_poids_v)
        y += 1
    if chk3.get() == True:
        totalChiffre = 0
        indice = 0
        page_export.write(4, y, "Chiffre d'affaire (en €)")
        x = 5  # indice pour la ligne dans le fichier xls
        while (indice != taille_cat):
            indice2 = 0  # variable pour l'indexation selon la taille de la liste des structures
            val = 0  # variable qui va recevoir les valeurs de la requete
            while (indice2 != taille_struct):
                cur.execute(
                    "SELECT sum(Montant), Catégorie FROM Vente, Catégorie, Lignes_vente, Organisation WHERE Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_vente=Vente.Id_Vente AND Lignes_vente.Id_catégorie=Catégorie.Id_catégorie AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie" % \
                    (struct_list[indice2], str(date_first), str(date_end), list_category[indice]))
                ChiffreAffaire = cur.fetchall()
                if len(ChiffreAffaire) == 0:
                    val += 0
                else:
                    for chiffre in ChiffreAffaire:
                        val += chiffre[0]
                        totalChiffre += chiffre[0]
                indice2 += 1
            if indice2 == taille_struct:
                if val != 0:
                    page_export.write(x, y, val)
                    x += 1
                else:
                    page_export.write(x, y, 'NR')
                    x += 1
            indice += 1
        page_export.write(x + 1, y, totalChiffre)
        y += 1
    if chk4.get() == True:
        totalNbrC = 0
        indice = 0
        page_export.write(4, y, "Nombre d'objet collecté")
        x = 5  # indice pour la ligne dans le fichier xls
        while (indice != taille_cat):
            indice2 = 0  # variable pour l'indexation selon la taille de la liste des structures
            val = 0  # variable qui va recevoir les valeurs de la requete
            while (indice2 != taille_struct):
                cur.execute(
                    "SELECT count(Id_Produit)*nombre, Catégorie FROM Produit, Arrivage, Catégorie, Organisation WHERE Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_catégorie=Catégorie.Id_catégorie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie" % \
                    (struct_list[indice2], str(date_first), str(date_end), list_category[indice]))
                NombreCollect = cur.fetchall()
                if len(NombreCollect) == 0:
                    val += 0
                else:
                    for nbrC in NombreCollect:
                        val += nbrC[0]
                        totalNbrC += nbrC[0]
                indice2 += 1
            if indice2 == taille_struct:
                if val != 0:
                    page_export.write(x, y, val)
                    x += 1
                else:
                    page_export.write(x, y, 'NR')
                    x += 1
            indice += 1
        page_export.write(x + 1, y, totalNbrC)
        y += 1
    if chk5.get() == True:
        totalNbrV = 0
        indice = 0
        page_export.write(4, y, "Nombre d'objet vendu")
        x = 5  # indice pour la ligne dans le fichier xls
        while (indice != taille_cat):
            indice2 = 0  # variable pour l'indexation selon la taille de la liste des structures
            val = 0  # variable qui va recevoir les valeurs de la requete
            while (indice2 != taille_struct):
                cur.execute(
                    "SELECT count(Id_ligne_vente), Catégorie FROM Lignes_vente, Vente, Catégorie, Organisation WHERE Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_catégorie=Catégorie.Id_catégorie AND Vente.Id_Vente=Lignes_vente.Id_vente AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie" % \
                    (struct_list[indice2], str(date_first), str(date_end), list_category[indice]))
                NombreVendu = cur.fetchall()
                if len(NombreVendu) == 0:
                    val += 0
                else:
                    for nbrV in NombreVendu:
                        val += nbrV[0]
                        totalNbrV += nbrV[0]
                indice2 += 1
            if indice2 == taille_struct:
                if val != 0:
                    page_export.write(x, y, val)
                    x += 1
                else:
                    page_export.write(x, y, 'NR')
                    x += 1
            indice += 1
        page_export.write(x + 1, y, totalNbrV)
        y += 1
    classeur_export.save(filename)
    TailleFile = len(filename)
    FileCSV = filename[:TailleFile - 4]
    with xlrd.open_workbook(filename) as wb:
        sh = wb.sheet_by_index(0)
        with open(FileCSV + '.csv', 'w') as f:
            c = csv.writer(f)
            for r in range(sh.nrows):
                c.writerow(sh.row_values(r))
    os.remove(filename)
    try:
        messagebox.showinfo('Exportation réussie', 'Votre fichier a bien été chargé')
    except:
        messagebox.showinfo('Exportation échouée', 'Votre fichier n\'a pas pu se charger')
    second_fen.destroy()


FirstFen = Tk()  # initialisation de la première fenetre
FirstFen.title("Structure")

LabelStruct = Label(FirstFen, text='Structure :', font=60)
LabelStruct.grid(row=0, column=0, ipady=30, padx=5)

chkValueStruct = BooleanVar()
chkValueStruct.set(False)

CheckStruct = Checkbutton(FirstFen, var=chkValueStruct, command=lambda: is_check_struct())
CheckStruct.grid(row=1, column=1, padx=4)

List = []  # initialise un tableau
List.append("tout")

cur.execute("SELECT Recyclerie FROM Organisation")  # récupère les recycleries insérées dans la base de données
OrgaList = cur.fetchall()

for row in OrgaList:  # insère les recycleries dans le tableau
    List.append(row[0])

ComboStruct = ttk.Combobox(FirstFen, values=List, width=28,
                           state='disabled')  # combobox qui récupère les recycleries comme valeur
ComboStruct.set("Choississez la/les structure(s)")  # valeur par défaut du combobox
ComboStruct.grid(row=1, column=2)

BtnAdd = Button(FirstFen, text='Ajouter', command=lambda: add_struct(),
                state='disabled')  # bouton qui ajoute la recylerie sélectionné du combobox dans la listbox
BtnAdd.grid(row=1, column=3)

BtnDel = Button(FirstFen, text='Supprimer', command=lambda: del_struct(),
                state='disabled')  # bouton qui supprime la recylerie sélectionné dans la listbox
BtnDel.grid(row=1, column=4, padx=5)

listStruct = Listbox(FirstFen, width=50)  # listbox qui va contenir les recyleries à étudier
listStruct.grid(row=1, column=5)

LabelSecteur = Label(FirstFen, text='Secteur :', font=60)
LabelSecteur.grid(row=2, column=0, ipady=70, padx=5)

LabelInsee = Label(FirstFen, text='Insee :')
LabelInsee.grid(row=3, column=1)

chkValue1 = BooleanVar()
chkValue1.set(False)

chkValue2 = BooleanVar()
chkValue2.set(False)

CheckOne = Checkbutton(FirstFen, text='1 code :', var=chkValue1,
                       command=lambda: is_check1())  # si on coche, active la ligne pour l'insertion d'1 code sinon reste désactivé
CheckOne.grid(row=3, column=2, padx=4)

EntryInsee = Entry(FirstFen, state='disabled')  # désactive tant que la checkbox n'est pas coché
EntryInsee.grid(row=3, column=3, padx=10)

CheckMany = Checkbutton(FirstFen, text='plusieurs codes :', var=chkValue2,
                        command=lambda: is_check2())  # si on coche, active la ligne pour l'insertion de plusieurs codes sinon reste désactivé
CheckMany.grid(row=4, column=2, padx=4)

BtnInsee = Button(FirstFen, text='Importer votre fichier', command=lambda: file_insee(),
                  state='disabled')  # désactive tant que la checkbox n'est pas coché
BtnInsee.grid(row=4, column=3)

LabelFile = Label(FirstFen, text='')
LabelFile.grid(row=4, column=4)

listInsee = Listbox(FirstFen, height=10)
listInsee.grid(row=4, column=5, padx=40)

BtnNext = Button(FirstFen, text='Suivant', command=lambda: new_window(
    listStruct))  # passe à la prochaine fenetre et prend en compte les données inscrites de la première fenetre
BtnNext.grid(row=6, column=5, padx=40, pady=20)

FirstFen.mainloop()
