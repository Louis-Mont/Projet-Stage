
import pypyodbc
import datetime
import csv
import sqlite3

#-----------------------------------------------------------------------------------------#
#fontions principales:

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
def InsertionInsee():

    print('\nLecture des codes insee...')
    fileInsee = 'C:/Users/david/Desktop/ProjetStage/communes2020.csv' # chemin du document de l'insee (.csv le format)

    Lecture = csv.reader(open(fileInsee, "r", encoding='utf-8'), delimiter=',')
    next(Lecture, None) # on retire l'en-tete
    for row in Lecture:
        Code = row[1] # on récupère les code de l'insee
        Comm = row[7] # on récupère les communes de l'insee
        curSQL.execute('INSERT INTO Code_Insee (Code, Commune) VALUES (?, ?)', (Code, Comm)) # insertion des codes dans le dictionnaire par rapport à la commune
   
    print('Lecture terminé\n')

#--------------------------------------------------------------------------------------------------------------
# Code principal

NameBase = 'GDR' # nom de votre base ODBC

print("\nconnexion en cours à la base GDR")
conn = pypyodbc.connect(DSN=NameBase)  # initialisation de la connexion au serveur
curGDR = conn.cursor()
print("connexion ok\n") 

try:
    connect = sqlite3.connect('C:/Users/david/Desktop/ProjetStage/test.db') # connexion à la database
    curSQL = connect.cursor()
    print("Base de données connectée à SQLite")
    
    CreateTable()

    InsertionInsee()

    curSQL.close()
    conn.close()
    print("La connexion SQLite est fermée")
except sqlite3.Error as error:
    print("Erreur lors de la connexion à SQLite", error)

# print(Code['Froissy']) 

curGDR.close()