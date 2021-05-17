import sqlite3
import pandas as pd
import unidecode
import csv

connect = sqlite3.connect("finale.db") # initialisation de la connexion à la base sql
cur = connect.cursor()

# Parcourt le fichier contenant les codes insee avec la commune correspondante et les insère dans la base de donnée
print("Insertion des codes Insee...")
NameFile = 'BD INSEE 2020'
ClasseurInsee = pd.read_excel(NameFile +'.xlsm')
ClasseurInsee.to_csv(NameFile+'.csv', index=False)
test=csv.reader(open(NameFile+'.csv', "r", encoding='utf-8'), delimiter=',')
next(test, None)
for row in test:
    Comm = row[1].upper().replace("-", " ").replace("'", "") 
    Comm = unidecode.unidecode(Comm)
    cur.execute("INSERT OR IGNORE INTO Insee (Code, Commune) VALUES (?, ?)", (row[0], Comm))

# Parcourt le fichier contenant les cooordonnées de chaque commune avec le code insee correspondant et met a jour les données de la table Insee
NameFile = 'Coordonnées GPS - France Entière'
ClasseurInsee = pd.read_excel(NameFile + '.xlsx')
ClasseurInsee.to_csv(NameFile+'.csv', index=False)
test=csv.reader(open(NameFile+'.csv', "r", encoding='utf-8'), delimiter=',')
next(test, None)
for row in test:
    cur.execute("UPDATE Insee SET Longitude = ?, Latitude = ? WHERE Code = ?", (row[2], row[1], row[0]))

connect.commit()
print('Insertion terminée !')
connect.close()