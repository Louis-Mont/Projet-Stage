
import pypyodbc
import datetime
import csv
import xlrd
import xlwt

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

    print('Lecture terminé')
    return CodeInsee
    
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

print("connexion en cours")
conn = pypyodbc.connect(DSN=NameBase)  # initialisation de la connexion au serveur
cur = conn.cursor()
print("connexion ok") 

Code = insee() # récupération des codes dans une variable

print(Code['Froissy']) # test