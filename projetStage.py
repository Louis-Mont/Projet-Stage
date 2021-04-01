
import pypyodbc
import datetime
import csv
import sqlite3

#-----------------------------------------------------------------------------------------#
#fontions principales:

# fonction pour récupérer les données utiles de l'insee
def insee():
    CodeInsee = {} 

    # récupération des code communes de l'insee
    print('Lecture des codes insee')
    fileInsee = 'C:/Users/david/Desktop/ProjetStage/communes2020.csv' # chemin du document de l'insee (.csv le format)

    Lecture = csv.reader(open(fileInsee, "r", encoding='utf-8'), delimiter=',')
    for row in Lecture:
        Code = row[1] # on récupère les code de l'insee
        Comm = row[7] # on récupère les communes de l'insee
        CodeInsee[Comm] = Code # insertion des codes dans le dictionnaire par rapport à la commune

    print('Lecture terminé\n')
    return CodeInsee

def CreateDatabase():

    try:
        connect = sqlite3.connect('C:/Users/david/Desktop/ProjetStage/test.db') # connexion à la database
        curSQL = connect.cursor()
        print("Base de données connectée à SQLite")
    
        CreateTableVente = """CREATE TABLE if not exists Vente (
                                Id_vente INTEGER NOT NULL PRIMARY KEY,
                                Id_Produits INTEGER NOT NULL,
                                Commune_orig TEXT,
                                Code_postal_orig TEXT,
                                Code_insee INTEGER NOT NULL,
                                Flux TEXT,
                                Categorie TEXT,
                                Sous_categorie TEXT,
                                Montant_HT REAL,
                                Montant_TTC REAL,
                                Date_vente DATE,
                                CONSTRAINT Id_produits_fk FOREIGN KEY(Id_Produits) REFERENCES Produits(Id_Produits)
                            );"""

        CreateTableLigne_Vente = """CREATE TABLE if not exists Ligne_Vente (
                                Id_collecte INTEGER NOT NULL,
                                Id_vente INTEGER NOT NULL,
                                CONSTRAINT Id_collecte_fk FOREIGN KEY(Id_collecte) REFERENCES Collecte(Id_collecte),
                                CONSTRAINT Id_vente_fk FOREIGN KEY(Id_vente) REFERENCES Vente(Id_vente)
                            );"""

        CreateTableCollecte = """CREATE TABLE if not exists Collecte (
                                Id_collecte INTEGER NOT NULL PRIMARY KEY,
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
                                
        CreateTableProduit = """CREATE TABLE if not exists Produits (
                                Id_Produits INTEGER NOT NULL PRIMARY KEY,
                                Nbr_produit INTEGER,
                                Poids_produits REAL
                            );"""                                             

        curSQL.execute(CreateTableVente)
        curSQL.execute(CreateTableLigne_Vente)
        curSQL.execute(CreateTableCollecte)
        curSQL.execute(CreateTableProduit)


        curSQL.close()
        conn.close()
        print("La connexion SQLite est fermée")
    except sqlite3.Error as error:
        print("Erreur lors de la connexion à SQLite", error)


"""
def VerificationCritere(nom):
    File = xlrd.open_workbook('C:/Users/david/Desktop/ProjetStage/%s.xls' % nom)

    test = File.sheet_by_name('Collecte')

    # Col? =(test.col(?))
"""

# a voir a la fin pour l'excel
"""
def Excel(nom):
    FileExcel = xlwt.Workbook()
    FeuilleCollect = FileExcel.add_sheet('Collecte')

    # tableau des titres de collecte
    ColTitreCollect = ["Identifiant unique de la vente","Commune d'origine du client","Code postal de la commune d'origine", "code insee","Flux", "Catégorie", "Sous-Catégorie", "Nombre de produits", "Poids des produits", "Montant HT", "Montant TTC", "IDProduit"]
    
    for i in range(len(ColTitreCollect)):
        FeuilleCollect.write(1,i,ColTitreCollect[i])

    FeuilleVente = FileExcel.add_sheet('Vente')

    # tableau des titres de vente
    ColTitreVente = ["id produit", "id arrivage", "Date arrivage", "Origine arrivage", "id arrivage", "Code insee", "Flux", "Catégorie", "Sous-Catégorie", "Quantité", "Poids", "Affectation"]

    for i in range(len(ColTitreVente)):
        FeuilleVente.write(1,i,ColTitreVente[i])

    FileExcel.save('C:/Users/david/Desktop/ProjetStage/%s.xls' % nom) # chemin pour sauvegarder le fichier
"""

NameBase = 'GDR' # nom de votre base ODBC

print("\nconnexion en cours à la base GDR")
conn = pypyodbc.connect(DSN=NameBase)  # initialisation de la connexion au serveur
curGDR = conn.cursor()
print("connexion ok\n") 

Code = insee() # récupération des codes dans une variable
CreateDatabase()

# print(Code['Froissy']) 

curGDR.close()