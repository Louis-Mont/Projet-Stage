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
    curGDR.execute("SELECT Commune, CodePostal FROM Commune")
    CommuneList = curGDR.fetchall()

    curSQL.execute("SELECT Id_Recyclerie FROM Organisation WHERE Recyclerie = ?", (RecyclerieNomGDR))
    id_Comm = curSQL.fetchone()
    for row in id_Comm:
        ID_Comm = row

    

    for row in CommuneList:
        Commune = row[0].upper()
        Commune = unidecode.unidecode(Commune)
        CodePostal = row[1]
        curSQL.execute("SELECT Id_Insee FROM Insee WHERE Commune = (?)", (Commune,))
        id_insee = curSQL.fetchone()
        for row2 in id_insee:
            ID_insee = row2
        curSQL.execute("INSERT INTO Commune (Commune, Code_postal, Id_Recyclerie, Id_Insee) VALUES (?,?,?,?)", (Commune, CodePostal, ID_Comm, ID_insee))
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
    mot = mot.replace("'", "\\'")
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
    print("insertion des données effectué")
    insertComm()

# on regarde si la recyclerie existe sinon on insert les données
for row in RecyclerieNomSQL:
    nomSQL = row[0].upper()
    nomSQL = unidecode.unidecode(nomSQL)
    if nomSQL == nom:
        print("Recyclerie déjà existante, veuillez utiliser une autre recyclerie")
    else:
        curSQL.execute("INSERT OR IGNORE INTO Organisation (Recyclerie) VALUES (?) ", (RecyclerieNomGDR))
        connect.commit()
        print("insertion des données effectué")

        # insertion des communes selon la recyclerie

insertComm()
        
connect.close()

conn.close()
curGDR.close()

