import pypyodbc
import datetime
import csv
import sqlite3
import xlrd
import pandas as pd
import unidecode
from xlwt import Workbook, easyxf

#-----------------------------------------------------------------------------------------#
#fontions principales:

def RecupInsee():
    print("Récupération des codes Insee...")
    NameFile = 'BD INSEE 2020'
    CodeInsee = {}
    ClasseurInsee = pd.read_excel(NameFile +'.xlsm')
    ClasseurInsee.to_csv(NameFile+'.csv', index=False)
    test=csv.reader(open(NameFile+'.csv', "r", encoding='utf-8'), delimiter=',')
    next(test, None)
    for row in test:
        Code = row[0] # on récupère les code de l'insee
        Comm = row[1].upper().replace("-", " ").replace("'", "") # on récupère les communes de l'insee
        Comm = unidecode.unidecode(Comm)
        CodeInsee[Comm] = Code
    print('Récupération terminée !\n')
    #print(CodeInsee)

def Collecte(classeur, style):
    Collect = {0:"id produit", 
                1:"id arrivage", 
                2:"Date arrivage",
                3:"Origine arrivage",
                4:"si déchèterie : nom de la déchèterie; si apport sur site ou rendez-vous = nom de la commune d'origine",
                5:"Code insee",
                6:"Flux",
                7:"Catégorie",
                8:"Quantité",
                9:"Poids",
                10:"Affectation"}

    CollectSheet = classeur.add_sheet("BD_Collecte")
    for col, title in Collect.items():
        CollectSheet.write(0, col, title, style)

def Vente(classeur, style):

    Vente = {0:"id vente",
                1:"Code postal",
                2:"Code insee",
                3:"Flux",
                4:"Catégorie",
                5:"Nombre de produit",
                6:"Poids des produits",
                7:"Montant HT",
                8:"Montant TTC"}

    VenteSheet = classeur.add_sheet("BD_Vente")
    for col, title in Vente.items():
        VenteSheet.write(0, col, title, style)

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

print("\nconnexion en cours à la base d'extraction")
conn = pypyodbc.connect(DSN='Extraction')  # initialisation de la connexion au serveur
curGDR = conn.cursor()
print("connexion ok\n")

style = easyxf("font: bold 1; alignment: horizontal center")

RecupInsee()

print('Création du classeur d\'extraction')
classeur = Workbook()
Collecte(classeur, style)
Vente(classeur, style)

classeur.save("extractionGDR.xlsx")
print('Classeur sauvegardé !')

conn.close()
curGDR.close()