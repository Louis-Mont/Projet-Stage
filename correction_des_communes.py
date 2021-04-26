import sqlite3
from math import sin, cos, acos, pi
import csv
import unidecode

fichier = "ancienne_commune.csv"
test=csv.reader(open(fichier, "r", encoding='latin-1'), delimiter=',')
next(test, None)
listing_insee={}

def dico_insee() :
    for row in test:
        try :
            a=row[3]
            if a.find("'") :
                a=a.replace("'","\\'")
            a=unidecode.unidecode(a.strip().upper()).replace("-", " ")
            b=row[0]
            listing_insee[a]= b
        except :
            print("fin")

def lec_dico(mot):
    for k,v in listing_insee.items() :
        if k == mot :
            Code = v
            return Code

def deg2rad(dd):
    """Convertit un angle "degrés décimaux" en "radians"
    """
    return dd/180*pi

def distanceGPS(latA, longA, latB, longB):
    """Retourne la distance en mètres entre les 2 points A et B connus grâce à
       leurs coordonnées GPS (en radians).
    """
    # Rayon de la terre en mètres (sphère IAG-GRS80)
    RT = 6378137
    # angle en radians entre les 2 points
    S = acos(sin(latA)*sin(latB) + cos(latA)*cos(latB)*cos(abs(longB-longA)))
    # distance entre les 2 points, comptée sur un arc de grand cercle
    return S*RT

def correction(connect, curSQL, ID, LONG, LATIT):

    curSQL.execute("SELECT Commune, Code_Postal FROM Commune WHERE Id_Insee = 'None' AND Id_Recyclerie = ?", (ID,))
    Comm = curSQL.fetchall()

    dico_insee()
    if not Comm:
        print('Il n\'y a pas de commune à corriger')
        connect.close()
    else:
        print('Commune sans code insee trouvé !\nCorrection en cours...')
        for row in Comm:
            test = row[0] + 'S'
            curSQL.execute("SELECT Id_Insee FROM Insee WHERE Commune = (?)", (test,))
            id_insee = curSQL.fetchone()
            id_insee = str(id_insee)
            id_insee = id_insee.replace("(", "").replace(",", "").replace(")", "")
            if id_insee != 'None':
                curSQL.execute("UPDATE Commune SET Id_Insee = (?) WHERE Commune = (?) AND Id_Recyclerie = ?",(id_insee, row[0], ID,))
            else:
                Code = row[1]
                j = str(Code)
                j = j[:2] + '%'
                string = row[0][:8] + '%'
                curSQL.execute("SELECT Longitude, Latitude FROM Insee WHERE Commune LIKE '%s' AND Code LIKE '%s'" %\
                                (string, j))
                listGPS = curSQL.fetchall()
                for ligne in listGPS:
                    latA = deg2rad(ligne[0]) # Nord
                    longA = deg2rad(ligne[1]) # Est
                    # cooordonnées GPS en radians du 2ème point (ici, brive la gaillarde)
                    latB = deg2rad(LONG) # Nord
                    longB = deg2rad(LATIT) # Est
            
                    dist = distanceGPS(latA, longA, latB, longB)
                    verif = int(dist)
                    if verif < 49000:
                        curSQL.execute("SELECT Id_Insee FROM Insee WHERE Longitude = ? AND Latitude = ?", (ligne[0], ligne[1]))
                        id_insee = curSQL.fetchone()
                        curSQL.execute("UPDATE Commune SET Id_Insee = (?) WHERE Commune = (?) AND Id_Recyclerie = ?",(id_insee[0], row[0], ID,))
                    else:
                        curSQL.execute("UPDATE Commune SET Id_Insee = (?) WHERE Commune = (?) AND Id_Recyclerie = ?",(0, row[0], ID,))             

            if id_insee == 'None':
                code = lec_dico(row[0])
                if code == None:
                    id_insee = 0
                else:
                    curSQL.execute("SELECT Id_Insee FROM Insee WHERE Code = ?", (code,))
                    id_insee = curSQL.fetchone()[0]
                curSQL.execute("UPDATE Commune SET Id_Insee = (?) WHERE Commune = (?) AND Id_Recyclerie = ?",(id_insee, row[0], ID,))
                
        print("Correction effectué !")

