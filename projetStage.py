import pypyodbc
import datetime
import csv
import sqlite3
import xlrd
import pandas as pd
import unidecode

#-----------------------------------------------------------------------------------------#
#fontions principales:

def insertComm():
    curGDR.execute("SELECT Commune, CodePostal, Déchèterie, Apport, Domicile FROM Commune")
    CommuneList = curGDR.fetchall()

    curSQL.execute("SELECT Id_Recyclerie FROM Organisation WHERE Recyclerie = ?", (RecyclerieNomGDR))
    id_Comm = curSQL.fetchone()
    for row in id_Comm:
        ID_Comm = row
    
    curSQL.execute("SELECT Commune FROM Commune WHERE Id_Recyclerie = (?)", (str(ID_Comm)))
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
            curSQL.execute("INSERT INTO Commune (Commune, Code_postal, Id_Recyclerie, Id_Insee, Apport, Déchèterie, Domicile) VALUES (?,?,?,?,?,?,?)", (Commune, CodePostal, ID_Comm, id_insee, Apport, Déchet, Domicile))
        connect.commit()

def insertArr():
    curGDR.execute("SELECT to_char(Date,'DD/MM/YYYY'), Origine, Poids_total FROM Arrivage")
    ArrivList = curGDR.fetchall()

    curSQL.execute("SELECT Id_Recyclerie FROM Organisation WHERE Recyclerie = ?", (RecyclerieNomGDR))
    id_Orga = curSQL.fetchone()
    for row in id_Orga:
        ID_Orga = row

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
for row in RecyclerieNomGDR:
    nom = row.upper()
    nom = unidecode.unidecode(nom)

curSQL.execute('SELECT Recyclerie FROM Organisation')
RecyclerieNomSQL = curSQL.fetchall()

curSQL.execute('SELECT count(*) FROM Organisation')
compteur = curSQL.fetchone()
for row in compteur:
    compt = row

# si la table Organisation est vide
if(compt == 0):
    curSQL.execute("INSERT OR IGNORE INTO Organisation (Recyclerie) VALUES (?) ", (RecyclerieNomGDR))
    connect.commit()

    insertComm()
    print("insertion des données effectué")
    
else:
    # on regarde si la recyclerie existe sinon on insert les données
    for row in RecyclerieNomSQL:
        nomSQL = row[0].upper()
        nomSQL = unidecode.unidecode(nomSQL)
        if nomSQL == nom:
            print("Recyclerie déjà existante, veuillez utiliser une autre recyclerie")
        else:
            curSQL.execute("INSERT OR IGNORE INTO Organisation (Recyclerie) VALUES (?) ", (RecyclerieNomGDR))
            connect.commit()

insertComm()
print("insertion des données effectué")
#insertArr()    
connect.close()

conn.close()
curGDR.close()

