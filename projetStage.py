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
    curGDR.execute("SELECT to_char(Date,'DD/MM/YYYY'), Origine, Poids_total, Commune.Commune FROM Arrivage, Commune WHERE Commune.IDCommune = Arrivage.IDCommune")
    ArrivList = curGDR.fetchall()

    ID_Orga = IDStructure()

    curSQL.execute("SELECT Id_Commune,Commune FROM Commune WHERE Id_Recyclerie = ?", str(ID_Orga))
    CommSQL = curSQL.fetchall()
    CommTab = {}
    for row in CommSQL:
        CommTab[row[1]] = row[0]

    curGDR.execute("SELECT to_char(Date,'DD/MM/YYYY'), Origine, Poids_total FROM Arrivage WHERE IDCommune = 0")
    ArrivList2 = curGDR.fetchall()
    for row in ArrivList2:
        Date = row[0]
        orig = row[1]
        poids = row[2]
        curSQL.execute("INSERT INTO Arrivage (Date, Id_commune, origine, poids_total, Id_recyclerie, Id_tournée) VALUES (?,?,?,?,?,?)", (Date, 0, orig, poids, ID_Orga, 1))
        connect.commit()

    for row in ArrivList:
        Date = row[0]
        orig = row[1]
        poids = row[2]
        Comm = row[3]
        Comm = Comm.replace("'","").replace("-"," ")
        ID_comm = CommTab[Comm]
        curSQL.execute("INSERT INTO Arrivage (Date, Id_commune, origine, poids_total, Id_recyclerie, Id_tournée) VALUES (?,?,?,?,?,?)", (Date, ID_comm, orig, poids, ID_Orga, 1))
        connect.commit()

def InsertProduit():
    ID_Struc = IDStructure()

    curGDR.execute('SELECT Produit.Nombre, Produit.Poids, Flux.Flux, Etat_produit.Désignation FROM Flux, Produit, Etat_produit WHERE Produit.IDFlux = Flux.IDFlux AND Etat_produit.IDEtat_produit = Produit.IDEtat_produit')
    List = curGDR.fetchall()

    Categories = list()
    curGDR.execute('SELECT Categorie.Désignation FROM Produit, Categorie WHERE Produit.IDCatégorie = Categorie.IDCatégorie')
    ListCat = curGDR.fetchall()
    for row in ListCat:
        Cat = row[0]
        Cat = Cat.upper()
        Cat = Cat.replace("'", "").replace("-", "").replace("/", "").replace(" ", "")
        Cat = unidecode.unidecode(Cat)
        Categories.append(Cat)

    IDCat = cat()

    for row in List:
        Nombre = row[0]
        Poids = row[1]
        Flux = row[2]
        Orient = row[3]
        curSQL.execute('INSERT INTO Produit (Orientation, Id_catégorie, Flux, nombre, Id_recyclerie, Poids) VALUES (?,?,?,?,?,?)', (Orient, IDCat, Flux, Nombre, ID_Struc, Poids))
        #connect.commit()

def InsertTournee():
    ID_Struc = IDStructure()

    curGDR.execute('SELECT Intitulé FROM Tournee')
    List = curGDR.fetchall()

    for row in List:
        Tournee = row[0]
        curSQL.execute('INSERT INTO Tournée (Tournée, Id_recyclerie) VALUES (?,?)', (Tournee, ID_Struc))
        connect.commit()           

def cat():
    CatCSV=csv.reader(open('categories.csv', "r", encoding='ISO-8859-1'), delimiter=',')
    next(CatCSV, None)
    for row in CatCSV:
        Mobilier = row[0]
        Mobilier = Mobilier.upper()
        curSQL.execute("SELECT Id_Catégorie FROM Catégorie WHERE Catégorie LIKE '%s' " %\
                    (Mobilier))
        IDCat = curSQL.fetchone()
    return IDCat

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

test = cat()
print(test)

print("insertion des données effectué")
 
connect.close()

conn.close()
curGDR.close()

