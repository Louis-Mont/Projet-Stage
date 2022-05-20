import os
from tkinter import END, messagebox, filedialog
import pandas as pd
from xlwt import Workbook
import sqlite3


def export(frame, list_cat, day_start, month_start, year_start, day_end, month_end, year_end, struct_list, w_coll,
           w_sold, revenue, n_coll, n_sold,
           list_box_modal):
    """
    fonction d'export du csv contenant les données demandées

    Arguments:
        frame {Tk} -- Deuxième fenêtre
        listCat {listbox} -- listbox contenant les catégories choisis
        StructList {list[str]} -- liste contenant les structures choisis
        ListBoxModal {listbox} -- listbox contenant les modalités choisis
    """
    filename = filedialog.asksaveasfilename(defaultextension='.xls',
                                            filetypes=[("xls files", '*.xls')])

    # connexion à la base de données sinon stop le programme et affiche l'erreur
    cur = None
    try:
        # connect = sqlite3.connect("X:/finale.db")
        connect = sqlite3.connect("finale.db")
        cur = connect.cursor()
    except IOError:
        print(IOError)
    #
    size_cat = list_cat.size()

    list_categorie = [c for c in list_cat.get(0, END)]  # création d'une liste des catégories sélectionnées
    list_modal = [m for m in list_box_modal.get(0, END)]  # création d'une liste des modales sélectionnées

    jour_first = day_start.get()
    mois_first = month_start.get()
    an_first = year_start.get()
    jour_sec = day_end.get()
    mois_sec = month_end.get()
    an_sec = year_end.get()
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
        if w_coll.get() and taille_modal != 0:
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
        elif w_coll.get() and taille_modal == 0:
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
        if w_sold.get():
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
        if revenue.get():
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
        if n_coll.get() and taille_modal != 0:
            if not w_coll.get():
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
        elif n_coll.get() and taille_modal == 0:
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
        if n_sold.get():
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
        if w_coll.get() and taille_modal != 0:
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
        elif w_coll.get() and taille_modal == 0:
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
        if w_sold.get():
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
        if revenue.get():
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
        if n_coll.get() and taille_modal != 0:
            indice = 0  # variable pour l'indexation selon la taille de la liste des catégories
            if not w_coll.get():
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
        elif n_coll.get() and taille_modal == 0:
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
        if n_sold.get():
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
    frame.destroy()  # quitte l'application
