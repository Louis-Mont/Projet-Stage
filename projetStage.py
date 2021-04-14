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
    curGDR.execute("SELECT to_char(Date,'DD/MM/YYYY'), Origine, Poids_total FROM Arrivage")
    ArrivList = curGDR.fetchall()

    ID_Orga = IDStructure()

    curSQL.execute("SELECT Id_Commune,Commune FROM Commune")
    CommSQL = curSQL.fetchall()
    CommTab = list()
    for row in CommSQL:
        CommTab.append(row[1])

    curGDR.execute("SELECT Arrivage.IDCommune, Commune.Commune FROM Arrivage, Commune WHERE Commune.IDCommune = Arrivage.IDCommune")
    CommList = curGDR.fetchall()
    CommTab2 = list()
    for row in CommList:
        ligne = row[1]
        ligne = ligne.replace("'", "").replace("-", " ")
        CommTab2.append(ligne)
 
    for row in CommTab2:
        Comm = row
        if Comm in CommTab:
            ID_comm = 1
        else:
            ID_comm = None

    for row in ArrivList:
        Date = row[0]
        orig = row[1]
        poids = row[2]
        curSQL.execute("INSERT INTO Arrivage (Date, Id_commune, origine, poids_total, Id_recyclerie) VALUES (?,?,?,?,?)", (Date, ID_comm, orig, poids, ID_Orga))
        connect.commit()

def InsertProduit():
    ID_Struc = IDStructure()

    curGDR.execute('SELECT Produit.Nombre, Produit.Poids, Flux.Flux FROM Flux, Produit WHERE Produit.IDFlux = Flux.IDFlux')
    List = curGDR.fetchall()

    for row in List:
        Nombre = row[0]
        Poids = row[1]
        Flux = row[2]
        curSQL.execute('INSERT INTO Produit (Id_orientation, Id_catégorie, Flux, nombre, Id_recyclerie, Poids) VALUES (?,?,?,?,?,?)', (1, 1, Flux, Nombre, ID_Struc, Poids))
        connect.commit()

def cat():  
    curGDR.execute('SELECT Désignation FROM Categorie')
    t=curGDR.fetchall()
    for val in t:
        a=str(val).lower()
        a=unidecode.unidecode(a)
    curGDR.execute('SELECT Produit.IDProduit, Catégorie.Désignation FROM Produit INNER JOIN Catégorie ON Produit.IDCatégorie = Catégorie.IDCatégorie')
    t=curGDR.fetchall()

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
#insertComm()
InsertProduit()
print("insertion des données effectué")

 
connect.close()

conn.close()
curGDR.close()

