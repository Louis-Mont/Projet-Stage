import sqlite3
import pandas as pd
import unidecode
import csv

connect = sqlite3.connect("finale.db")
cur = connect.cursor()

print("Insertion des codes Insee...")
NameFile = 'BD INSEE 2020'
ClasseurInsee = pd.read_excel(NameFile +'.xlsm')
ClasseurInsee.to_csv(NameFile+'.csv', index=False)
test=csv.reader(open(NameFile+'.csv', "r", encoding='utf-8'), delimiter=',')
next(test, None)
for row in test:
    Code = row[0] # on récupère les code de l'insee
    Comm = row[1].upper().replace("-", " ").replace("'", "") # on récupère les communes de l'insee
    Comm = unidecode.unidecode(Comm)
    cur.execute("INSERT OR IGNORE INTO Insee (Code, Commune) VALUES (?, ?)", (Code, Comm))
connect.commit()
print('Insertion terminée !\n')

connect.close()