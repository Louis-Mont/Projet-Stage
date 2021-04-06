import pypyodbc
import datetime
import csv
import sqlite3
import easygui
import pandas as pd
import openpyxl

#-----------------------------------------------------------------------------------------#
#fontions principales:

# fonction pour convertir le fichier insee xlsm en csv
def Convert():

    file = easygui.fileopenbox() # l'utilisateur va chercher le fichier .xlsm dans son pc
    workbook = pd.read_excel(file)
    print ("Veuillez saisir le nom de votre fichier csv :")
    NameCSV=input() + '.csv'

    FileInsee = workbook.to_csv(NameCSV, index=False) # conversion du fichier en .csv

    return NameCSV

# fonction pour créer les tables de la BDD
def CreateTable():

    CreateTableVente = """CREATE TABLE if not exists Vente (
                                Id_Vente INTEGER NOT NULL PRIMARY KEY,
                                Id_Lots INTEGER NOT NULL,
                                Commune_Orig TEXT,
                                Code_Postal_orig TEXT,
                                Id_Insee INTEGER NOT NULL,
                                Flux TEXT,
                                Categorie TEXT,
                                Sous_Categorie TEXT,
                                Montant_HT REAL,
                                Montant_TTC REAL,
                                Date_Vente DATE,
                                CONSTRAINT Id_lots_fk FOREIGN KEY(Id_Lots) REFERENCES Lots(Id_Lots),
                                CONSTRAINT Id_insee_fk FOREIGN KEY(Id_Insee) REFERENCES Lots(Id_Insee)
                            );"""

    CreateTableLigne_Vente = """CREATE TABLE if not exists Ligne_Vente (
                                Id_Ligne_Vente INTEGER NOT NULL PRIMARY KEY,
                                Id_Produit INTEGER NOT NULL,
                                Id_Vente INTEGER NOT NULL,
                                Date_Vente DATE,
                                Montant_HT REAL,
                                Montant_TTC REAL,
                                CONSTRAINT Id_collecte_fk FOREIGN KEY(Id_Collecte) REFERENCES Collecte(Id_Collecte),
                                CONSTRAINT Id_vente_fk FOREIGN KEY(Id_Vente) REFERENCES Vente(Id_Vente)
                            );"""

    CreateTableCollecte = """CREATE TABLE if not exists Collecte (
                                Id_Produit INTEGER NOT NULL PRIMARY KEY,
                                Date_arrivage DATE,
                                Orig_arrivage TEXT,
                                Code_insee INTEGER NOT NULL,
                                Flux TEXT,
                                Categorie TEXT,
                                Sous_categorie TEXT,
                                Qte INTEGER,
                                Poids REAL,
                                Affectation TEXT
                            );""" # affectation = orientation (rebuts, valorisé, ...)

    CreateTableInsee = """CREATE TABLE if not exists Code_Insee (
                                Id_Insee INTEGER NOT NULL PRIMARY KEY,
                                Code INTEGER,
                                Commune TEXT
                            );"""                       
                                
    CreateTableLot = """CREATE TABLE if not exists Lots (
                                Id_Lots INTEGER NOT NULL PRIMARY KEY,
                                Id_Ligne_Vente INTEGER NOT NULL,
                                Nbr_produit INTEGER,
                                Poids_produits REAL
                            );"""                                             

    #curSQL.execute(CreateTableVente)
    #curSQL.execute(CreateTableLigne_Vente)
    #curSQL.execute(CreateTableCollecte)
    #curSQL.execute(CreateTableLot)
    curSQL.execute(CreateTableInsee)

# fonction pour récupérer et insérer les codes insee dans la BDD
def InsertionInsee(FileInsee):

    print('\nLecture des codes insee...')

    Lecture = csv.reader(open(FileInsee, "r", encoding='utf-8'), delimiter=',')
    next(Lecture, None) # on retire l'en-tete
    for row in Lecture:
        Code = row[0] # on récupère les code de l'insee
        Comm = row[1] # on récupère les communes de l'insee
        curSQL.execute('INSERT INTO Code_Insee (Code, Commune) VALUES (?, ?)', (Code, Comm)) # insertion des codes insee dans la bdd
   
    print('Lecture terminé\n')

#--------------------------------------------------------------------------------------------------------------
# Code principal

print("\nconnexion en cours à GDR")
conn = pypyodbc.connect(DSN='GDR')  # initialisation de la connexion au serveur
curGDR = conn.cursor()
print("connexion ok\n")

# partie fichier INSEE
print ("Voulez-vous convertir votre fichier xlsm ? [O/N]")
reponse = input()
if(reponse == "O"):
    print ("fichier xlsm en cours de conversion...")
    FileInsee = Convert()
    print ("fichier xlsm convertit en csv avec succès !\n")
elif(reponse == "N"):
    print ("Sélectionner votre fichier csv")
    FileInsee = easygui.fileopenbox()
    print ("Fichier sélectionné")

# partie base de donnée
print ("Voulez-vous créer une base de données ? [O/N]")
reponse = input()
if(reponse == "O"):
    print ("Veuillez choisir l'emplacement de votre base de données :\n")
    CheminBDD = easygui.filesavebox() + '.db'
    print ("base de données créée avec succès !\n")
elif(reponse == "N"):
    print ("Veuillez vous connecter à la base de données :\n")
    CheminBDD = easygui.fileopenbox()
    
print ("Connexion de la base de données...\n")
try:
    connect = sqlite3.connect(CheminBDD) # connexion à la database
    curSQL = connect.cursor()
    print("Base de données connectée à SQLite")
    
    CreateTable()# création des tables 

    InsertionInsee(FileInsee)# insertion des données INSEE dans la base

    curSQL.close()

    conn.close()
    print("La connexion SQLite est fermée")
except sqlite3.Error as error:
    print("Erreur lors de la connexion à SQLite", error)

curGDR.close()