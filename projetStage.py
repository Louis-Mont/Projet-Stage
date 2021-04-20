import pypyodbc
import datetime
import csv
import sqlite3
import xlrd
import pandas as pd
import unidecode

#-----------------------------------------------------------------------------------------#
#fontions principales:

def IDStructure():
    curSQL.execute("SELECT Id_Recyclerie FROM Organisation WHERE Recyclerie = ?", (RecyclerieNomGDR))
    id_Comm = curSQL.fetchone()
    for row in id_Comm:
        ID_Comm = row

    return ID_Comm

def insertComm():
    curGDR.execute("SELECT Commune, CodePostal, Déchèterie, Apport, Domicile FROM Commune")
    CommuneList = curGDR.fetchall()

    ID_Struc = IDStructure()
    
    curSQL.execute("SELECT Commune FROM Commune WHERE Id_Recyclerie = (?)", (str(ID_Struc)))
    CommSQL = curSQL.fetchall()
    CommTab = list()
    for row in CommSQL:
        verif = row[0].upper()
        verif = verif.replace("'", "").replace("-", " ")
        CommTab.append(verif)

    for row in CommuneList:
        Commune = row[0].upper()
        Commune = unidecode.unidecode(Commune)
        Commune = Commune.replace("'", "").replace("-", " ")
        CodePostal = row[1]
        Déchet = row[2]
        Apport = row[3]
        Domicile = row[4]
        curSQL.execute("SELECT Id_Insee FROM Insee WHERE Commune = (?)", (Commune,))
        id_insee = curSQL.fetchone()
        id_insee = str(id_insee)
        id_insee = id_insee.replace("(", "").replace(",", "").replace(")", "")
        if Commune not in CommTab:
            curSQL.execute("INSERT INTO Commune (Commune, Code_postal, Id_Recyclerie, Id_Insee, Apport, Déchèterie, Domicile) VALUES (?,?,?,?,?,?,?)", (Commune, CodePostal, ID_Struc, id_insee, Apport, Déchet, Domicile))
        connect.commit()

def insertArr():
    ID_Orga = IDStructure()

    curSQL.execute("SELECT Id_Commune,Commune FROM Commune WHERE Id_Recyclerie = ?", str(ID_Orga))
    CommSQL = curSQL.fetchall()
    CommTab = {}
    for row in CommSQL:
        CommTab[row[1]] = row[0]

    curSQL.execute("SELECT Id_Tournée, Tournée FROM Tournée WHERE Id_recyclerie = ?", str(ID_Orga))
    TournéeSQL = curSQL.fetchall()
    TournéeDic = {}
    for row in TournéeSQL:
        TournéeDic[row[1]] = row[0]

    curSQL.execute("SELECT max(Id_Arrivage) FROM Arrivage")
    test=curSQL.fetchone() [0]

    curGDR.execute("SELECT to_char(Date,'DD/MM/YYYY'), Origine, Poids_total, Tournée.Intitulé, IDArrivage FROM Arrivage, Tournée WHERE IDCommune = 0 AND Tournée.IDTournée = Arrivage.IDTournée")
    ArrivList2 = curGDR.fetchall()
    for row in ArrivList2:
        Date = row[0]
        orig = row[1]
        poids = row[2]
        tournée = row[3]
        tournée = tournée.replace("'", "").replace("-", " ")
        ID_tour = TournéeDic[tournée]
        if not test:
            Id_arr = row[4]
        else:
            Id_arr = test + row[4]
        curSQL.execute("INSERT INTO Arrivage (Id_arrivage, Date, Id_commune, origine, poids_total, Id_recyclerie, Id_tournée) VALUES (?,?,?,?,?,?,?)", (Id_arr, Date, 0, orig, poids, ID_Orga, ID_tour))
        connect.commit()

    curGDR.execute("SELECT to_char(Date,'DD/MM/YYYY'), Origine, Poids_total, Commune.Commune, IDArrivage FROM Arrivage, Commune WHERE Commune.IDCommune = Arrivage.IDCommune")
    ArrivList = curGDR.fetchall()
    for row in ArrivList:
        Date = row[0]
        orig = row[1]
        poids = row[2]
        Comm = row[3]
        Comm = Comm.replace("'","").replace("-"," ")
        ID_comm = CommTab[Comm]
        if not test:
            Id_arr = row[4]
        else:
            Id_arr = test + row[4]
        curSQL.execute("INSERT INTO Arrivage (Id_arrivage, Date, Id_commune, origine, poids_total, Id_recyclerie, Id_tournée) VALUES (?,?,?,?,?,?,?)", (Id_arr, Date, ID_comm, orig, poids, ID_Orga, 0))
        connect.commit()

def InsertProduit():
    ID_Struc = IDStructure()

    curGDR.execute('SELECT Produit.Nombre, Produit.Poids, Flux.Flux, Etat_produit.Désignation, Categorie.Désignation, Produit.IDArrivage FROM Flux, Produit, Etat_produit, Categorie WHERE Produit.IDFlux = Flux.IDFlux AND Etat_produit.IDEtat_produit = Produit.IDEtat_produit AND Produit.IDCatégorie = Categorie.IDCatégorie')
    List = curGDR.fetchall()

    curSQL.execute("SELECT max(Id_Arrivage) FROM Produit")
    test=curSQL.fetchone() [0]
    for row in List:
        Nombre = row[0]
        Poids = row[1]
        Flux = row[2]
        Flux = Flux.upper().replace("'", "").replace("-", "").replace("/", "").replace(" ", "")
        IDFlux = flux(Flux)
        Orient = row[3]
        Categorie = row[4]
        Categorie = Categorie.upper().replace("'", "").replace("-", "").replace("/", "").replace(" ", "")
        IDCat = cat(Categorie)
        if not test:
            ID_arr = row[5]
        else:
            ID_arr = test + row[5]
        curSQL.execute('INSERT INTO Produit (Orientation, Id_catégorie, Id_Flux, nombre, Id_recyclerie, Poids, Id_arrivage) VALUES (?,?,?,?,?,?,?)', (Orient, IDCat, IDFlux, Nombre, ID_Struc, Poids, ID_arr))
        connect.commit()

def InsertTournee():
    ID_Struc = IDStructure()

    curGDR.execute('SELECT Intitulé FROM Tournee')
    List = curGDR.fetchall()

    for row in List:
        Tournee = row[0]
        Tournee = Tournee.replace("'", "").replace("-"," ")
        curSQL.execute('INSERT INTO Tournée (Tournée, Id_recyclerie) VALUES (?,?)', (Tournee, ID_Struc))
        connect.commit()           

def InsertVente():
    ID_Struc=IDStructure()
    
    curGDR.execute("SELECT to_char(date,'DD/MM/YYYY'),code_postal,ville,montant_total,tauxremise from vente_magasin")
    b = curGDR.fetchall()
    for venteorigine in b :
        ville=venteorigine[2]
        if ville.find("'") :
            ville=ville.replace("'"," ")
        IdInsee = Ville(ville)
        curSQL.execute("INSERT INTO Vente (Id_insee, Date, Code_Postal, Commune, Montant_total, TauxRemise, Id_recyclerie) VALUES('%s', %s,'%s','%s','%s','%s','%s') " %\
                    (IdInsee, venteorigine[0],venteorigine[1],ville,venteorigine[3],venteorigine[4], ID_Struc))
        curSQL.execute("SELECT max(Id_vente) FROM Vente")
        venteoriginemax=curSQL.fetchone() [0]
        curGDR.execute("SELECT Categorie.Désignation,lignes_vente.montant,lignes_vente.poids,lignes_vente.tauxtva,lignes_vente.montanttva, Flux.Flux FROM lignes_vente, Sous_Categorie, Flux, Categorie WHERE idvente_magasin = '%s' AND lignes_vente.IDSous_Catégorie = Sous_Categorie.IDSous_Catégorie AND Sous_Categorie.IDFlux = Flux.IDFlux AND Categorie.IDCatégorie = Lignes_Vente.IDCatégorie" %\
                    (venteoriginemax))
        c=curGDR.fetchall()
        for lignevente in c :
            Categorie = lignevente[0]
            Categorie = Categorie.upper().replace("'", "").replace("-", "").replace("/", "").replace(" ", "")
            IDCat = cat(Categorie)
            Flux = lignevente[5]
            Flux = Flux.upper().replace("'", "").replace("-", "").replace("/", "").replace(" ", "")
            IDFlux = flux(Flux)
            curSQL.execute("INSERT INTO Lignes_vente (Id_catégorie,Montant,Poids,Taux_tva,Montant_tva,Id_vente, Id_Flux) values ('%s','%s','%s','%s','%s','%s','%s')" %\
                        (IDCat,lignevente[1],lignevente[2],lignevente[3],lignevente[4],venteoriginemax, IDFlux))
            connect.commit()   

def Ville(ville):

    ville = ville.upper().replace("-", " ").replace("'", "")
    ville = unidecode.unidecode(ville)
    curSQL.execute("SELECT Id_Insee, Commune FROM Insee")
    insee = curSQL.fetchall()
    for row in insee:
        if ville == row[1]:
            IdInsee = row[0]
            break
        else:
            IdInsee = 0

    return IdInsee

def cat(cat):
    
    MotClé = {"MOB":1, "MEUBLE":1, "AMEUBLEMENT":1,
                "ELECTRO":2, "APPAREIL":2,
                "LITTERATURE":3, "CULTURE":3,
                "BIBELOT":4, "VAISSELLE":4,
                "TEXTILE":5,
                "INFO":6, "MULTIMEDIA":6,
                "JEU":7, "JOUET":7,
                "BRICO":8, "JARD":8, "OUTIL":8, "NATURE":8,
                "LOISIR":9, "SPORT":9,
                "DECO":10,
                "CYCLE":11}

    IDCat = 12
    for mot, Id in MotClé.items():
        if cat.find(mot) != -1:
            IDCat = Id
            break
        else:
            continue

    return IDCat

def flux(flux):

    MotClé = {"TOUT":1, "VENANT":1, "ENCOMBRANT":1,
                "DEA":2,
                "DEEE":3,
                "METAUX":4, "FERRAILLE":4,
                "PAPIER":5,
                "TEXTILE":6, "TLC":6,
                "GRAVAT":7,
                "BOIS":8,
                "CARTON":9,
                "VERRE":10,
                "DECHET":11,
                "POLYSTYRENE":12}

    IDFlux = 1
    for mot, Id in MotClé.items():
        if flux.find(mot) != -1:
            IDFlux = Id
            break
        else:
            continue

    return IDFlux

def remplacement(mot):
    mot = mot.replace("'", "\\")
    return (mot)

def date():
    jourj = datetime.date.today()
    annuel = jourj - datetime.timedelta(days=366)
    annuel = str(annuel)
    annuel = annuel.replace("-","")
    
#--------------------------------------------------------------------------------------------------------------
# Code principal

print("connexion en cours à la base de la recyclerie à extraire")
conn = pypyodbc.connect(DSN='Extraction')  # initialisation de la connexion au serveur
curGDR = conn.cursor()
print("connexion ok\n")

print("connexion en cours à la grosse base de données")
connect = sqlite3.connect("finale.db")
curSQL = connect.cursor()
print("connexion ok\n")

# insertion du nom de la recyclerie dans la grosse base de données
curGDR.execute("SELECT RaisonSociale FROM Organisation")
RecyclerieNomGDR = curGDR.fetchone()

curSQL.execute("INSERT OR IGNORE INTO Organisation (Recyclerie) VALUES (?) ", (RecyclerieNomGDR))
connect.commit()

#InsertTournee()
#insertComm()
#insertArr()
#InsertProduit()
#InsertVente()

print("insertion des données effectué")
 
connect.close()

conn.close()
curGDR.close()

