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
cur = None
try:
    # connect = sqlite3.connect("X:/finale.db")
    connect = sqlite3.connect("finale.db")
    cur = connect.cursor()
except IOError:
    print(IOError)


def is_check_struct():
    """
    fonction qui change l'état des widgets en lien avec les structures
    """
    if chkValueStruct.get():
        ComboStruct.config(state='normal')
        BtnAdd.config(state='normal')
        BtnDel.config(state='normal')
    else:
        ComboStruct.config(state='disabled')
        BtnAdd.config(state='disabled')
        BtnDel.config(state='disabled')


def is_check1():
    """
    fonction qui change l'état des widgets en lien avec 1 code insee
    """
    if chkInseeOne.get():
        EntryInsee.config(state='normal')
        CheckMany.config(state='disabled')
    else:
        EntryInsee.config(state='disabled')
        CheckMany.config(state='normal')


def is_check2():
    """
    fonction qui change l'état des widgets en lien avec plusieurs codes insee
    """
    if chkInseeMany.get():
        BtnInsee.config(state='normal')
        CheckOne.config(state='disabled')
    else:
        BtnInsee.config(state='disabled')
        CheckOne.config(state='normal')
        listInsee.delete(0, END)
        LabelFile.config(text='')


def file_insee():
    """
    fonction pour chercher le fichier .txt contenant les plusieurs codes insee et l'insère dans la listbox
    """
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Sélectionnez votre fichier",
                                          filetypes=(("Fichier Texte", "*.txt"),
                                                     ("Fichier pdf", "*.pdf")))

    LabelFile.config(text=filename)
    fichier = open(str(filename), "r")
    lines = fichier.readlines()
    fichier.close()
    listInsee.delete(0, END)
    i = 1
    for ligne in lines:
        listInsee.insert(i, ligne.strip())
        i += 1


def requete_struct():
    """
    fonction pour les requêtes sql selon les choix de la premiere fenêtre et retourne la liste des structures demandées
    """
    list_struct = []
    if chkInseeOne.get() and not chkValueStruct.get():  # si l'utilisateur coche seulement la checkbox pour 1 code Insee
        val = EntryInsee.get()
        cur.execute(
            "SELECT DISTINCT Recyclerie FROM Organisation, Insee, Commune WHERE Organisation.Id_Recyclerie = "
            "Commune.Id_Recyclerie AND Commune.Id_insee = Insee.Id_Insee AND Code = '%s' " %
            val)
        struct_list = cur.fetchall()
        list_struct = [struct[0] for struct in struct_list]
    # si l'utilisateur coche seulement la checkbox pour plusieurs codes INSEE
    elif chkInseeMany.get() and not chkValueStruct.get():
        val = listInsee.get(0, END)
        for code in val:
            cur.execute(
                "SELECT Recyclerie FROM Organisation, Insee, Commune WHERE Organisation.Id_Recyclerie = "
                "Commune.Id_Recyclerie AND Commune.Id_insee = Insee.Id_Insee AND Code = '%s'" %
                code)
            struct_list = cur.fetchall()
            for struct in struct_list:
                list_struct.append(struct[0])
    # si l'utilisateur coche la checkbox pour 1 code INSEE et la checkbox pour choisir les structures
    elif chkInseeOne.get() and chkValueStruct.get():
        val = EntryInsee.get()
        cur.execute(
            "SELECT Recyclerie FROM Organisation, Insee, Commune WHERE Organisation.Id_Recyclerie = "
            "Commune.Id_Recyclerie AND Commune.Id_insee = Insee.Id_Insee AND Code = '%s'" %
            val)
        struct_list = cur.fetchall()
        list_struct = [struct[0] for struct in struct_list]
        print(list_struct)
    # si l'utilisateur coche la checkbox pour pour plusieurs codes INSEE et la checkbox pour choisir les structures
    elif chkInseeMany.get() and chkValueStruct.get():
        print('kl')
    # si l'utilisateur coche seulement la checkbox pour choisir les structures
    elif chkValueStruct.get() and not chkInseeOne.get() and not chkInseeMany.get():
        struct_list = listStruct.get(0, END)
        list_struct = [s for s in struct_list]

    return list(set(list_struct))


def modalite_collect(chk_value1, chk_value4, combo_modalite, label_modalite, btn_add_modal, btn_del_modal,
                     list_box_modal):
    """
    fonction qui affiche les widgets en lien avec les modalités
    si les checkbox des collectes sont cochées sinon les cachent
    """
    if chk_value1.get():
        combo_modalite.grid(row=7, column=2)
        label_modalite.grid(row=7, column=1)
        btn_add_modal.grid(row=7, column=3)
        btn_del_modal.grid(row=7, column=4)
        list_box_modal.grid(row=7, column=5)
    elif chk_value4.get():
        combo_modalite.grid(row=7, column=2)
        label_modalite.grid(row=7, column=1)
        btn_add_modal.grid(row=7, column=3)
        btn_del_modal.grid(row=7, column=4)
        list_box_modal.grid(row=7, column=5)
    else:
        combo_modalite.grid_forget()
        label_modalite.grid_forget()
        btn_add_modal.grid_forget()
        btn_del_modal.grid_forget()
        list_box_modal.grid_forget()


def new_window():
    """
    fonction pour la deuxième fenêtre
    """
    list_struct = requete_struct()

    FirstFen.destroy()  # Détruit la première fenêtre
    SecondFen = Tk()  # initialisation de la première fenêtre
    SecondFen.title("Options d'importation")

    # Partie catégorie :
    label_cat = Label(SecondFen, text='Catégorie :', font=60)
    label_cat.grid(row=0, column=0, ipady=30)

    list_categorie_box = ["tout"]  # initialise un tableau
    cur.execute("SELECT Catégorie FROM Catégorie")
    cat = cur.fetchall()
    for row in cat:
        list_categorie_box.append(row[0])

    combo = ttk.Combobox(SecondFen, values=list_categorie_box, width=29)
    combo.set("Choississez la/les catégorie(s)")
    combo.grid(row=0, column=1)
    combo.bind("<<ComboboxSelected>>", lambda e: combo.get())

    btn_add2 = Button(SecondFen, text='Ajouter', command=lambda: add_cat(list_cat, combo, list_categorie_box))
    btn_add2.grid(row=0, column=2)

    btn_del2 = Button(SecondFen, text='Supprimer', command=lambda: del_cat(list_cat))
    btn_del2.grid(row=0, column=3, padx=5)

    list_cat = Listbox(SecondFen, width=50)
    list_cat.grid(row=0, column=5)

    # Partie temps :
    label_temps = Label(SecondFen, text='Temps :', font=60)
    label_temps.grid(row=1, column=0, ipady=10)

    cur.execute("SELECT Min(date) FROM Arrivage")
    date_bdd = cur.fetchone()[0]
    an_bdd = date_bdd[:4]
    an = datetime.date.today().year
    an = int(an)
    dates = lambda l, u: [str(i).zfill(2) for i in range(l, u)]
    jours = dates(1, 32)
    mois = dates(1, 13)
    ans = [i for i in range(int(an_bdd), an + 1)]
    """jours = [str(i).zfill(2) for i in range(1, 32)]
    mois = [str(i).zfill(2) for i in range(1, 13)]
    ans = [i for i in range(int(an_bdd), an + 1)]"""
    mess_debut = Label(SecondFen, text="choisissez la date de début :")
    mess_debut.grid(row=2, column=1, padx=10)
    choix_jour1 = ttk.Combobox(SecondFen, values=jours)
    choix_jour1.grid(row=2, column=2)
    choix_jour1.set('01')
    choix_mois1 = ttk.Combobox(SecondFen, values=mois)
    choix_mois1.grid(row=2, column=3)
    choix_mois1.set('01')
    choix_an1 = ttk.Combobox(SecondFen, values=ans)
    choix_an1.grid(row=2, column=4)
    choix_an1.set(int(an_bdd))
    mess_fin = Label(SecondFen, text="choisissez la date de fin :")
    mess_fin.grid(row=3, column=1, pady=10)
    choix_jour2 = ttk.Combobox(SecondFen, values=jours)
    choix_jour2.grid(row=3, column=2)
    choix_jour2.set('31')
    choix_mois2 = ttk.Combobox(SecondFen, values=mois)
    choix_mois2.grid(row=3, column=3)
    choix_mois2.set('12')
    choix_an2 = ttk.Combobox(SecondFen, values=ans)
    choix_an2.grid(row=3, column=4)
    choix_an2.set(an)

    # Partie quantitative :
    label_qte = Label(SecondFen, text='Quantitative :', font=60)
    label_qte.grid(row=4, column=0, ipady=40, padx=5)

    label_modalite = Label(SecondFen, text='Modalité(s) :')
    list_modalite_combo = ['tout']
    cur.execute('SELECT DISTINCT origine FROM Arrivage')
    modal = cur.fetchall()
    for row in modal:
        if row[0] != '0':
            list_modalite_combo.append(row[0])
    combo_modalite = ttk.Combobox(SecondFen, values=list_modalite_combo, width=25)
    combo_modalite.set("Choississez la/les modalité(s)")
    btn_add_modal = Button(SecondFen, text='ajouter',
                           command=lambda: add_modal(list_box_modal, combo_modalite, list_modalite_combo))
    btn_del_modal = Button(SecondFen, text='supprimer', command=lambda: del_modal(list_box_modal))
    list_box_modal = Listbox(SecondFen, width=30)

    chk_value1 = BooleanVar(value=False)
    chk_value2 = BooleanVar(value=False)
    chk_value3 = BooleanVar(value=False)
    chk_value4 = BooleanVar(value=False)
    chk_value5 = BooleanVar(value=False)

    chkbox1 = Checkbutton(SecondFen, text='Poids collecté (en Kg)', var=chk_value1,
                          command=lambda: modalite_collect(chk_value1, chk_value4, combo_modalite, label_modalite,
                                                           btn_add_modal, btn_del_modal, list_box_modal))
    chkbox1.grid(row=5, column=1)
    chkbox2 = Checkbutton(SecondFen, text='Poids vendu (en Kg)', var=chk_value2)
    chkbox2.grid(row=5, column=2)
    chkbox3 = Checkbutton(SecondFen, text='Chiffre d\'affaire (en €)', var=chk_value3)
    chkbox3.grid(row=5, column=3)
    chkbox4 = Checkbutton(SecondFen, text='Nombre d\'objet collecté', var=chk_value4,
                          command=lambda: modalite_collect(chk_value1, chk_value4, combo_modalite, label_modalite,
                                                           btn_add_modal, btn_del_modal, list_box_modal))
    chkbox4.grid(row=6, column=1)
    chkbox5 = Checkbutton(SecondFen, text='Nombre d\'objet vendu', var=chk_value5)
    chkbox5.grid(row=6, column=2)

    btn_export = Button(SecondFen, text='Exporter',
                        command=lambda: export(SecondFen, list_cat, choix_jour1, choix_mois1, choix_an1, choix_jour2,
                                               choix_mois2, choix_an2, list_struct, chk_value1, chk_value2, chk_value3,
                                               chk_value4, chk_value5, list_box_modal))
    btn_export.grid(row=10, column=5, padx=40, pady=20)

    SecondFen.mainloop()


def add_modal(list_box_modal, combo_modalite, list_modalite_combo):
    """
    fonction du bouton ajouter pour des modalités
    """
    list_files = list_box_modal.get(0, END)
    value = combo_modalite.get()
    if combo_modalite.get() == 'tout':
        for modal in list_modalite_combo:
            if modal not in list_files:
                list_box_modal.insert(END, modal)
        idx = list_box_modal.get(0, END).index('tout')
        list_box_modal.delete(idx)

    if value not in list_files and value != "Choississez la/les modalité(s)" and value != 'tout':
        list_box_modal.insert(END, value)


def del_modal(list_box_modal):
    """
    fonction pour le bouton supprimer des modalités

    Arguments:
        listBoxModal {listbox} -- listbox contenant les modalités sélectionnées
    """
    item_selected = list_box_modal.curselection()
    list_box_modal.delete(item_selected[0])


def add_cat(list_cat, combo, list_cat_box):
    """
    fonction pour le bouton ajouter des catégories

    Arguments:
        listCat {Listbox} -- listbox contenant les catégories sélectionnées
        Combo {Combobox} -- combobox contenant la liste des catégories de la base sql
        ListCatBox {list[str]} -- liste contenant les catégories de la base sql
    """
    list_files = list_cat.get(0, END)
    value = combo.get()
    if combo.get() == 'tout':
        for cat in list_cat_box:
            if cat not in list_files:
                list_cat.insert(END, cat)
        idx = list_cat.get(0, END).index('tout')
        list_cat.delete(idx)

    if value not in list_files and value != 'Choississez la/les catégorie(s)' and value != 'tout':
        list_cat.insert(END, value)


def del_cat(list_cat):
    """
    fonction pour le bouton supprimer des catégories

    Arguments:
        listCat {listbox} -- listbox contenant les catégories sélectionnées
    """
    item_selected = list_cat.curselection()
    list_cat.delete(item_selected[0])


def add_struct():
    """
    fonction pour le bouton ajouter des structures
    """
    list_files = listStruct.get(0, END)
    value = ComboStruct.get()
    if ComboStruct.get() == 'tout':
        for struct in ListRecyclerieBox:
            if struct not in list_files:
                listStruct.insert(END, struct)
        idx = listStruct.get(0, END).index('tout')
        listStruct.delete(idx)

    if value not in list_files and value != 'tout' and value != 'Choississez la/les structure(s)':
        listStruct.insert(END, value)


def del_struct():
    """
    fonction pour le bouton supprimer des structures
    """
    item_selected = listStruct.curselection()
    listStruct.delete(item_selected[0])


def export(SecondFen, list_cat, jour1, mois1, an1, jour2, mois2, an2, struct_list, chk1, chk2, chk3, chk4, chk5,
           list_box_modal):
    """
    fonction d'export du csv contenant les données demandées

    Arguments:
        SecondFen {Tk} -- Deuxième fenêtre
        listCat {listbox} -- listbox contenant les catégories choisis
        StructList {list[str]} -- liste contenant les structures choisis
        ListBoxModal {listbox} -- listbox contenant les modalités choisis
    """
    filename = filedialog.asksaveasfilename(defaultextension='.xls',
                                            filetypes=[("xls files", '*.xls')])

    size_cat = list_cat.size()

    list_categorie = [c for c in list_cat.get(0, END)]  # création d'une liste des catégories sélectionnées
    list_modal = [m for m in list_box_modal.get(0, END)]  # création d'une liste des modales sélectionnées

    jour_first = jour1.get()
    mois_first = mois1.get()
    an_first = an1.get()
    jour_sec = jour2.get()
    mois_sec = mois2.get()
    an_sec = an2.get()
    date_first = an_first + "/" + mois_first + "/" + jour_first
    date_end = an_sec + "/" + mois_sec + "/" + jour_sec
    classeur_export = Workbook()
    page_export = classeur_export.add_sheet("EXPORT", cell_overwrite_ok=True)
    if size_cat == 0:  # si la liste des catégories sélectionnées est vide alors affiche les données totales
        page_export.write(0, 0, "Structure(s)")
        # taille_cat = len(list_categorie)
        taille_struct = len(struct_list)
        taille_modal = len(list_modal)
        y = 1
        if chk1.get() and taille_modal != 0:
            page_export.write(0, y, "Modalité(s)")
            page_export.write(0, y + 1, "Poids collecté (en kg)")
            indice = 0
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                indice2 = 0
                while indice2 != taille_modal:
                    cur.execute(
                        "SELECT sum(Poids)*nombre FROM Produit, Arrivage, Organisation WHERE Recyclerie = '%s' AND "
                        "Produit.Id_recyclerie=Organisation.Id_recyclerie AND "
                        "Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND origine = '%s'"
                        % (struct_list[indice], str(date_first), str(date_end), list_modal[indice2]))
                    poids_collect = cur.fetchone()[0]
                    page_export.write(x, 1, list_modal[indice2])
                    page_export.write(x, 0, struct_list[indice])
                    if poids_collect is None:
                        page_export.write(x, y + 1, 'NR')
                    else:
                        page_export.write(x, y + 1, round(poids_collect))
                    indice2 += 1
                    x += 1
                indice += 1
            y += 1  # on passe a la colonne suivante pour les données suivantes
        elif chk1.get() and taille_modal == 0:
            page_export.write(0, y, "Poids collecté (en kg)")
            indice = 0
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                cur.execute(
                    "SELECT sum(Poids)*nombre FROM Produit, Arrivage, Organisation WHERE Recyclerie = '%s' AND "
                    "Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_arrivage=Arrivage.Id_arrivage "
                    "AND date > '%s' AND date < '%s'" % (struct_list[indice], str(date_first), str(date_end)))
                poids_collect = cur.fetchone()[0]
                page_export.write(x, 0, struct_list[indice])
                if poids_collect is None:
                    page_export.write(x, y, 'NR')
                else:
                    page_export.write(x, y, round(poids_collect))
                x += 1
                indice += 1
            y += 1  # on passe a la colonne suivante pour les données suivantes
        if chk2.get():
            page_export.write(0, y, "Poids vendu (en kg)")
            indice = 0
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                cur.execute(
                    "SELECT sum(Lignes_vente.Poids) FROM Lignes_vente, Vente, Organisation WHERE Recyclerie = '%s' "
                    "AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_vente = Vente.Id_Vente "
                    "AND date > '%s' AND date < '%s'" % (struct_list[indice], str(date_first), str(date_end)))
                poids_vendu = cur.fetchone()[0]
                page_export.write(x, 0, struct_list[indice])
                if poids_vendu is None:
                    page_export.write(x, y, 'NR')
                else:
                    page_export.write(x, y, round(poids_vendu))
                indice += 1
                x += 1
            y += 1
        if chk3.get():
            page_export.write(0, y, "Chiffre d'affaire (en €)")
            indice = 0
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                cur.execute(
                    "SELECT sum(Montant) FROM Vente, Lignes_vente, Organisation WHERE Recyclerie = '%s' AND "
                    "Vente.Id_recyclerie=Organisation.Id_recyclerie AND Lignes_vente.Id_vente=Vente.Id_Vente AND date "
                    "> '%s' AND date < '%s'" % (struct_list[indice], str(date_first), str(date_end)))
                chiffre = cur.fetchone()[0]
                page_export.write(x, 0, struct_list[indice])
                if chiffre is None:
                    page_export.write(x, y, 'NR')
                else:
                    page_export.write(x, y, round(chiffre))
                indice += 1
                x += 1
            y += 1
        if chk4.get() and taille_modal != 0:
            if not chk1.get():
                page_export.write(0, y, "Modalité(s)")
            page_export.write(0, y + 1, "Nombre d'objet collecté")
            indice = 0
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                indice2 = 0
                while indice2 != taille_modal:
                    cur.execute(
                        "SELECT count(Id_Produit)*nombre FROM Produit, Arrivage, Organisation WHERE Recyclerie = '%s' "
                        "AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND "
                        "Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND origine = '%s'"
                        % (struct_list[indice], str(date_first), str(date_end), list_modal[indice2]))
                    nbr_collect = cur.fetchone()[0]
                    page_export.write(x, 1, list_modal[indice2])
                    page_export.write(x, 0, struct_list[indice])
                    if nbr_collect is None:
                        page_export.write(x, y + 1, 'NR')
                    else:
                        page_export.write(x, y + 1, round(nbr_collect))
                    indice2 += 1
                    x += 1
                indice += 1
            y += 1
        elif chk4.get() and taille_modal == 0:
            page_export.write(0, y, "Nombre d'objet collecté")
            indice = 0
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                cur.execute(
                    "SELECT count(Id_Produit)*nombre FROM Produit, Arrivage, Organisation WHERE Recyclerie = '%s' AND "
                    "Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_arrivage=Arrivage.Id_arrivage "
                    "AND date > '%s' AND date < '%s'" % (struct_list[indice], str(date_first), str(date_end)))
                nbr_collect = cur.fetchone()[0]
                page_export.write(x, 0, struct_list[indice])
                if nbr_collect is None:
                    page_export.write(x, y, 'NR')
                else:
                    page_export.write(x, y, round(nbr_collect))
                x += 1
                indice += 1
            y += 1
        if chk5.get():
            page_export.write(0, y, "Nombre d'objet vendu")
            indice = 0
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                cur.execute(
                    "SELECT count(Id_ligne_vente) FROM Lignes_vente, Vente, Organisation WHERE Recyclerie = '%s' AND "
                    "Vente.Id_recyclerie=Organisation.Id_recyclerie AND Vente.Id_Vente=Lignes_vente.Id_vente AND date "
                    "> '%s' AND date < '%s'" % (struct_list[indice], str(date_first), str(date_end)))
                nbr_vente = cur.fetchone()[0]
                page_export.write(x, 0, struct_list[indice])
                if nbr_vente is None:
                    page_export.write(x, y, 'NR')
                else:
                    page_export.write(x, y, round(nbr_vente))
                indice += 1
                x += 1
            y += 1
    else:
        x = 1
        page_export.write(0, 0, "Structure(s)")
        page_export.write(0, 1, "Catégorie(s)")
        taille_cat = len(list_categorie)
        taille_struct = len(struct_list)
        taille_modal = len(list_modal)
        y = 2
        '''
        if chk1.get() == True:
            indice = 0 # variable pour l'indexation selon la taille de la liste des catégories
            pageexport.write(0,y,"Poids collecté (en kg)")
            x = 1 # indice pour la ligne dans le fichier xls
            while (indice != TailleCat):
                indice2 = 0 # variable pour l'indexation selon la taille de la liste des structures
                val = 0 # variable qui va recevoir les valeurs de la requête
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
        if chk1.get() and taille_modal != 0:
            indice = 0  # variable pour l'indexation selon la taille de la liste des catégories
            page_export.write(0, y, "Modalité(s)")
            page_export.write(0, y + 1, "Poids collecté (en kg)")
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                indice2 = 0
                while indice2 != taille_cat:
                    indice3 = 0  # variable pour l'indexation selon la taille de la liste des modalités
                    while indice3 != taille_modal:
                        cur.execute(
                            "SELECT sum(Poids)*nombre FROM Produit, Catégorie, Arrivage, Organisation WHERE "
                            "Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND "
                            "Produit.Id_catégorie = Catégorie.Id_catégorie AND "
                            "Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date < '%s' AND Catégorie = "
                            "'%s' AND origine = '%s' GROUP BY Catégorie.Id_catégorie" %
                            (struct_list[indice], str(date_first), str(date_end), list_categorie[indice2],
                             list_modal[indice3]))
                        poids_collect = cur.fetchall()
                        page_export.write(x, 0, struct_list[indice])
                        page_export.write(x, 2, list_modal[indice3])
                        page_export.write(x, 1, list_categorie[indice2])
                        if len(poids_collect) == 0:
                            page_export.write(x, y + 1, 'NR')
                            x += 1
                        else:
                            for poids in poids_collect:
                                page_export.write(x, y + 1, round(poids[0]))
                                x += 1
                        indice3 += 1
                    indice2 += 1
                indice += 1
            y += 1  # on passe a la colonne suivante pour les données suivantes
        elif chk1.get() and taille_modal == 0:
            indice = 0  # variable pour l'indexation selon la taille de la liste des catégories
            page_export.write(0, y, "Poids collecté (en kg)")
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                indice2 = 0
                while indice2 != taille_cat:
                    cur.execute(
                        "SELECT sum(Poids)*nombre FROM Produit, Catégorie, Arrivage, Organisation WHERE Recyclerie = "
                        "'%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND Produit.Id_catégorie = "
                        "Catégorie.Id_catégorie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND date > '%s' AND date "
                        "< '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie" %
                        (struct_list[indice], str(date_first), str(date_end), list_categorie[indice2]))
                    poids_collect = cur.fetchall()
                    page_export.write(x, 0, struct_list[indice])
                    page_export.write(x, 1, list_categorie[indice2])
                    if len(poids_collect) == 0:
                        page_export.write(x, y, 'NR')
                        x += 1
                    else:
                        for poids in poids_collect:
                            page_export.write(x, y, round(poids[0]))
                            x += 1
                    indice2 += 1
                indice += 1
            y += 1  # on passe a la colonne suivante pour les données suivantes
        if chk2.get():
            indice = 0  # variable pour l'indexation selon la taille de la liste des catégories
            page_export.write(0, y, "Poids vendu (en kg)")
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                indice2 = 0  # variable pour l'indexation selon la taille de la liste des structures
                while indice2 != taille_cat:
                    cur.execute(
                        "SELECT sum(Lignes_vente.Poids) FROM Lignes_vente, Vente, Catégorie, Organisation WHERE "
                        "Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND "
                        "Lignes_vente.Id_vente = Vente.Id_Vente AND Lignes_vente.Id_catégorie=Catégorie.Id_catégorie "
                        "AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie" %
                        (struct_list[indice], str(date_first), str(date_end), list_categorie[indice2]))
                    poids_vente = cur.fetchall()
                    page_export.write(x, 0, struct_list[indice])
                    page_export.write(x, 1, list_categorie[indice2])
                    if len(poids_vente) == 0:
                        page_export.write(x, y, 'NR')
                        x += 1
                    else:
                        for poids in poids_vente:
                            page_export.write(x, y, round(poids[0]))
                            x += 1
                    indice2 += 1
                indice += 1
            y += 1  # on passe a la colonne suivante pour les données suivantes
        if chk3.get():
            indice = 0  # variable pour l'indexation selon la taille de la liste des catégories
            page_export.write(0, y, "Chiffre d'affaire (en €)")
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                indice2 = 0  # variable pour l'indexation selon la taille de la liste des structures
                while indice2 != taille_cat:
                    cur.execute(
                        "SELECT sum(Montant), Catégorie FROM Vente, Catégorie, Lignes_vente, Organisation WHERE "
                        "Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND "
                        "Lignes_vente.Id_vente=Vente.Id_Vente AND Lignes_vente.Id_catégorie=Catégorie.Id_catégorie "
                        "AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie" %
                        (struct_list[indice], str(date_first), str(date_end), list_categorie[indice2]))
                    chiffre = cur.fetchall()
                    page_export.write(x, 0, struct_list[indice])
                    page_export.write(x, 1, list_categorie[indice2])
                    if len(chiffre) == 0:
                        page_export.write(x, y, 'NR')
                        x += 1
                    else:
                        for c in chiffre:
                            page_export.write(x, y, round(c[0]))
                            x += 1
                    indice2 += 1
                indice += 1
            y += 1  # on passe a la colonne suivante pour les données suivantes
        if chk4.get() and taille_modal != 0:
            indice = 0  # variable pour l'indexation selon la taille de la liste des catégories
            if not chk1.get():
                page_export.write(0, y, "Modalité(s)")
            page_export.write(0, y + 1, "Nombre d'objet collecté")
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                indice2 = 0  # variable pour l'indexation selon la taille de la liste des structures
                while indice2 != taille_cat:
                    indice3 = 0  # variable pour l'indexation selon la taille de la liste des modalités
                    while indice3 != taille_modal:
                        cur.execute(
                            "SELECT count(Id_Produit)*nombre FROM Produit, Arrivage, Catégorie, Organisation WHERE "
                            "Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND "
                            "Produit.Id_catégorie=Catégorie.Id_catégorie AND Produit.Id_arrivage=Arrivage.Id_arrivage "
                            "AND date > '%s' AND date < '%s' AND Catégorie = '%s' AND origine = '%s' GROUP BY "
                            "Catégorie.Id_catégorie" %
                            (struct_list[indice], str(date_first), str(date_end), list_categorie[indice2],
                             list_modal[indice3]))
                        nbr_collect = cur.fetchall()
                        page_export.write(x, 0, struct_list[indice])
                        page_export.write(x, 2, list_modal[indice3])
                        page_export.write(x, 1, list_categorie[indice2])
                        if len(nbr_collect) == 0:
                            page_export.write(x, y + 1, 'NR')
                            x += 1
                        else:
                            for nbr in nbr_collect:
                                page_export.write(x, y + 1, round(nbr[0]))
                                x += 1
                        indice3 += 1
                    indice2 += 1
                indice += 1
            y += 1  # on passe a la colonne suivante pour les données suivantes
        elif chk4.get() and taille_modal == 0:
            indice = 0  # variable pour l'indexation selon la taille de la liste des catégories
            page_export.write(0, y, "Nombre d'objets collectés")
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                indice2 = 0  # variable pour l'indexation selon la taille de la liste des structures
                while indice2 != taille_cat:
                    cur.execute(
                        "SELECT count(Id_Produit)*nombre FROM Produit, Arrivage, Catégorie, Organisation WHERE "
                        "Recyclerie = '%s' AND Produit.Id_recyclerie=Organisation.Id_recyclerie AND "
                        "Produit.Id_catégorie=Catégorie.Id_catégorie AND Produit.Id_arrivage=Arrivage.Id_arrivage AND "
                        "date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie" %
                        (struct_list[indice], str(date_first), str(date_end), list_categorie[indice2]))
                    nbr_collect = cur.fetchall()
                    page_export.write(x, 0, struct_list[indice])
                    page_export.write(x, 1, list_categorie[indice2])
                    if len(nbr_collect) == 0:
                        page_export.write(x, y, 'NR')
                        x += 1
                    else:
                        for nbr in nbr_collect:
                            page_export.write(x, y, round(nbr[0]))
                            x += 1
                    indice2 += 1
                indice += 1
            y += 1  # on passe a la colonne suivante pour les données suivantes
        if chk5.get():
            indice = 0  # variable pour l'indexation selon la taille de la liste des catégories
            page_export.write(0, y, "Nombre d'objet vendu")
            x = 1  # indice pour la ligne dans le fichier xls
            while indice != taille_struct:
                indice2 = 0  # variable pour l'indexation selon la taille de la liste des structures
                while indice2 != taille_cat:
                    cur.execute(
                        "SELECT count(Id_ligne_vente) FROM Lignes_vente, Vente, Catégorie, Organisation WHERE "
                        "Recyclerie = '%s' AND Vente.Id_recyclerie=Organisation.Id_recyclerie AND "
                        "Lignes_vente.Id_catégorie=Catégorie.Id_catégorie AND Vente.Id_Vente=Lignes_vente.Id_vente "
                        "AND date > '%s' AND date < '%s' AND Catégorie = '%s' GROUP BY Catégorie.Id_catégorie" %
                        (struct_list[indice], str(date_first), str(date_end), list_categorie[indice2]))
                    nbr_vente = cur.fetchall()
                    page_export.write(x, 0, struct_list[indice])
                    page_export.write(x, 1, list_categorie[indice2])
                    if len(nbr_vente) == 0:
                        page_export.write(x, y, 'NR')
                        x += 1
                    else:
                        for nbr in nbr_vente:
                            page_export.write(x, y, round(nbr[0]))
                            x += 1
                    indice2 += 1
                indice += 1
            y += 1  # on passe a la colonne suivante pour les données suivantes
    try:  # sauvegarde le fichier et le convertit en csv
        classeur_export.save(filename)
        csv_file = os.path.splitext(filename)[0] + '.csv'
        read_file = pd.read_excel(filename)
        read_file.to_csv(csv_file, index=None, header=True)
        os.remove(filename)
        messagebox.showinfo('Exportation réussie', 'Votre fichier a bien été chargé')
    except IOError:
        messagebox.showerror('Exportation échouée', 'Votre fichier n\'a pas pu se charger')
        print(IOError)
    SecondFen.destroy()  # quitte l'application


FirstFen = Tk()  # initialisation de la première fenêtre
FirstFen.title("Structure")

LabelStruct = Label(FirstFen, text='Structure :', font=60)
LabelStruct.grid(row=0, column=0, ipady=30, padx=5)

chkValueStruct = BooleanVar()
chkValueStruct.set(False)

CheckStruct = Checkbutton(FirstFen, var=chkValueStruct, command=lambda: is_check_struct())
CheckStruct.grid(row=1, column=1, padx=4)

ListRecyclerieBox = ["tout"]  # initialise la liste de la combobox des structures

cur.execute("SELECT Recyclerie FROM Organisation")  # récupère les structures insérées dans la base de données
orga_list = cur.fetchall()

for row in orga_list:  # insère les recycleries dans la liste
    ListRecyclerieBox.append(row[0])

ComboStruct = ttk.Combobox(FirstFen, values=ListRecyclerieBox, width=28,
                           state='disabled')  # combobox qui récupère les recycleries comme valeur
ComboStruct.set("Choississez la/les structure(s)")  # valeur par défaut du combobox
ComboStruct.grid(row=1, column=2)

BtnAdd = Button(FirstFen, text='Ajouter', command=lambda: add_struct(),
                state='disabled')  # bouton qui ajoute la recylerie sélectionnée du combobox dans la listbox
BtnAdd.grid(row=1, column=3)

BtnDel = Button(FirstFen, text='Supprimer', command=lambda: del_struct(),
                state='disabled')  # bouton qui supprime la recylerie sélectionnée dans la listbox
BtnDel.grid(row=1, column=4, padx=5)

listStruct = Listbox(FirstFen, width=50)  # listbox qui va contenir les recyleries à étudier
listStruct.grid(row=1, column=5)

label_secteur = Label(FirstFen, text='Secteur :', font=60)
label_secteur.grid(row=2, column=0, ipady=70, padx=5)

LabelInsee = Label(FirstFen, text='Insee :')
LabelInsee.grid(row=3, column=1)

chkInseeOne = BooleanVar()
chkInseeOne.set(False)

chkInseeMany = BooleanVar()
chkInseeMany.set(False)

# si on coche, active la ligne pour l'insertion d'1 code sinon reste désactivé
CheckOne = Checkbutton(FirstFen, text='1 code :', var=chkInseeOne,
                       command=lambda: is_check1())
CheckOne.grid(row=3, column=2, padx=4)

EntryInsee = Entry(FirstFen, state='disabled')  # désactivé tant que la checkbox n'est pas cochée
EntryInsee.grid(row=3, column=3, padx=10)

# si on coche, active la ligne pour l'insertion de plusieurs codes sinon reste désactivé
CheckMany = Checkbutton(FirstFen, text='plusieurs codes :', var=chkInseeMany,
                        command=lambda: is_check2())
CheckMany.grid(row=4, column=2, padx=4)

BtnInsee = Button(FirstFen, text='Importer votre fichier', command=lambda: file_insee(),
                  state='disabled')  # désactivé tant que la checkbox n'est pas cochée
BtnInsee.grid(row=4, column=3)

LabelFile = Label(FirstFen, text='')
LabelFile.grid(row=4, column=4)

listInsee = Listbox(FirstFen, height=10)
listInsee.grid(row=4, column=5, padx=40)

# passe à la prochaine fenêtre et prends en compte les données inscrites de la première fenêtre
BtnNext = Button(FirstFen, text='Suivant',
                 command=lambda: new_window())
BtnNext.grid(row=6, column=5, padx=40, pady=20)

FirstFen.mainloop()
