import pypyodbc
import datetime
import csv
import sqlite3
import xlrd
import pandas as pd
import unidecode
import correction_des_communes
from math import *

#-----------------------------------------------------------------------------------------#
#fontions principales:

def IDStructure():
    curSQL.execute("SELECT Id_Recyclerie FROM Organisation WHERE Recyclerie = ?", (RecyclerieNomGDR))
    id_Comm = curSQL.fetchone()
    for row in id_Comm:
        ID_Comm = row

    return ID_Comm

def verifAnnee(MaxDate, annee):
    date = str(annee) + '0101'
    if MaxDate > date:
        return date
    else:
        date = str(annee-1) + '0101'
        return date

def insertComm():
    curGDR.execute("SELECT Commune, CodePostal, Déchèterie, Apport, Domicile FROM Commune WHERE EnrActif = 1 AND CodePostal != ''")
    CommuneList1 = curGDR.fetchall()

    curGDR.execute("SELECT Commune, Déchèterie, Apport, Domicile FROM Commune WHERE EnrActif = 1 AND CodePostal = ''")
    CommuneList2 = curGDR.fetchall()

    ID_Struc = IDStructure()
    curSQL.execute("SELECT Commune FROM Commune WHERE Id_Recyclerie = ?", (ID_Struc,))
    CommSQL = curSQL.fetchall()
    CommTab = list()
    for row in CommSQL:
        verif = row[0].upper()
        verif = verif.replace("'", "").replace("-", " ")
        CommTab.append(verif)

    for row in CommuneList1:
        Commune = row[0].upper()
        Commune = unidecode.unidecode(Commune)
        Commune = Commune.replace("'", "").replace("-", " ")
        Commune = Commune.strip(' ').replace(" ST ", " SAINT ").replace(" STE ", " SAINTE ")
        if Commune[:3] == "ST " or Commune[:4] =="STE ":
            Commune = Commune.replace("ST ", "SAINT ").replace("STE ","SAINTE ")
        CodePostal = row[1]
        Déchet = row[2]
        Apport = row[3]
        Domicile = row[4]
        codePost = str(CodePostal)
        codePost = codePost[:2] + '%'
        curSQL.execute("SELECT Id_Insee FROM Insee WHERE Commune = ? AND Code LIKE ?", (Commune, codePost))
        id_insee = curSQL.fetchone()
        id_insee = str(id_insee)
        id_insee = id_insee.replace("(", "").replace(",", "").replace(")", "")
        if Commune not in CommTab:
            curSQL.execute("INSERT INTO Commune (Commune, Code_postal, Id_Recyclerie, Id_Insee, Apport, Déchèterie, Domicile) VALUES('%s','%s','%s','%s','%s','%s','%s')" %\
                                (Commune, CodePostal, ID_Struc, id_insee, Apport, Déchet, Domicile))

    for row in CommuneList2:
        Commune = row[0].upper()
        Commune = unidecode.unidecode(Commune)
        Commune = Commune.replace("'", "").replace("-", " ")
        Commune = Commune.strip(' ').replace(" ST ", " SAINT ").replace(" STE ", " SAINTE ")
        if Commune[:3] == "ST " or Commune[:4] =="STE ":
            Commune = Commune.replace("ST ", "SAINT ").replace("STE ","SAINTE ")
        Déchet = row[1]
        Apport = row[2]
        Domicile = row[3]
        curSQL.execute("SELECT Id_Insee FROM Insee WHERE Commune = ?", (Commune,))
        id_insee = curSQL.fetchone()
        id_insee = str(id_insee)
        id_insee = id_insee.replace("(", "").replace(",", "").replace(")", "")
        if Commune not in CommTab:
            curSQL.execute("INSERT INTO Commune (Commune, Code_postal, Id_Recyclerie, Id_Insee, Apport, Déchèterie, Domicile) VALUES('%s','%s','%s','%s','%s','%s','%s')" %\
                                (Commune, 'NR', ID_Struc, id_insee, Apport, Déchet, Domicile))

def insertArr():
    ID_Orga = IDStructure()

    annee = datetime.date.today().year
    curGDR.execute("select max(to_char(Date,'YYYY/MM/DD')) from arrivage")
    MaxDate = curGDR.fetchone()[0].replace("/","")
    
    date = verifAnnee(MaxDate, annee)
    
    curSQL.execute("SELECT Max(date) FROM Arrivage WHERE Id_recyclerie = ?", (ID_Orga,))
    verifDate = curSQL.fetchone()[0]

    if verifDate == None:
        verifDate = '00000000'
    
    verifDate = verifDate.replace("/","")

    curSQL.execute("SELECT Id_Commune,Commune FROM Commune WHERE Id_Recyclerie = ?", (ID_Orga,))
    CommSQL = curSQL.fetchall()
    CommTab = {}
    for row in CommSQL:
        CommTab[row[1]] = row[0]

    curSQL.execute("SELECT Id_Tournée, Tournée FROM Tournée WHERE Id_recyclerie = ?", (ID_Orga,))
    TournéeSQL = curSQL.fetchall()
    TournéeDic = {}
    for row in TournéeSQL:
        TournéeDic[row[1]] = row[0]

    curSQL.execute("SELECT max(Id_Arrivage) FROM Arrivage")
    test=curSQL.fetchone() [0]

    curGDR.execute("SELECT to_char(Date,'YYYY/MM/DD'), Origine, Poids_total, IDArrivage FROM Arrivage WHERE IDCommune = 0 AND IDTournée = 0 AND date > %s AND date < %s " %\
            (verifDate,date))
    ArrivList3 = curGDR.fetchall()
    for row in ArrivList3:
        Date = row[0]
        orig = row[1]
        poids = row[2]
        if not test:
            Id_arr = row[3]
        else:
            Id_arr = test + row[3]
        curSQL.execute("INSERT INTO Arrivage (Id_arrivage, Date, Id_commune, origine, poids_total, Id_recyclerie, Id_tournée) VALUES (?,?,?,?,?,?,?)", (Id_arr, Date, 0, orig, poids, ID_Orga, 0))

    curGDR.execute("SELECT to_char(Date,'YYYY/MM/DD'), Origine, Poids_total, Tournee.Intitulé, IDArrivage FROM Arrivage, Tournee WHERE IDCommune = 0 AND Tournee.IDTournée = Arrivage.IDTournée AND date > %s AND date < %s " %\
            (verifDate,date))
    ArrivList2 = curGDR.fetchall()
    for row in ArrivList2:
        Date = row[0]
        orig = row[1]
        poids = row[2]
        tournée = row[3]
        tournée = tournée.replace("'", "").replace("-", " ")
        ID_tour = TournéeDic[tournée]
        if not test:
            Id_arr = row[4]
        else:
            Id_arr = test + row[4]
        curSQL.execute("INSERT INTO Arrivage (Id_arrivage, Date, Id_commune, origine, poids_total, Id_recyclerie, Id_tournée) VALUES (?,?,?,?,?,?,?)", (Id_arr, Date, 0, orig, poids, ID_Orga, ID_tour))

    curGDR.execute("SELECT to_char(Date,'YYYY/MM/DD'), Origine, Poids_total, Commune.Commune, IDArrivage FROM Arrivage, Commune WHERE Commune.IDCommune = Arrivage.IDCommune AND date > %s AND date < %s " %\
            (verifDate,date))
    ArrivList = curGDR.fetchall()
    for row in ArrivList:
        Date = row[0]
        orig = row[1]
        poids = row[2]
        Comm = row[3]
        Comm = Comm.upper().replace("'","").replace("-"," ")
        Comm = unidecode.unidecode(Comm)
        Comm = Comm.strip(' ')
        Comm = Comm.replace(" ST ", " SAINT ").replace(" STE ", " SAINTE ")
        if Comm[:3] == "ST " or Comm[:4] =="STE ":
            Comm = Comm.replace("ST ", "SAINT ").replace("STE ","SAINTE ")
        ID_comm = CommTab[Comm]
        if not test:
            Id_arr = row[4]
        else:
            Id_arr = test + row[4]
        curSQL.execute("INSERT INTO Arrivage (Id_arrivage, Date, Id_commune, origine, poids_total, Id_recyclerie, Id_tournée) VALUES (?,?,?,?,?,?,?)", (Id_arr, Date, ID_comm, orig, poids, ID_Orga, 0))

def ListPoids():

    curGDR.execute("SELECT IDSous_Catégorie FROM Sous_Categorie")
    IDSousCat = curGDR.fetchall()
    
    PoidsDico = {}
    NombreDico = {}
    for val in IDSousCat:
        curGDR.execute("SELECT count(IDProduit) FROM Produit WHERE IDSous_Catégorie = '%s'"%\
                            (val[0]))
        n = curGDR.fetchone()[0]

        NombreDico[val[0]] = n

        ListPoids = list()
        curGDR.execute("SELECT poids FROM Produit WHERE IDSous_Catégorie = '%s' ORDER BY poids"%\
                                (val[0]))
        l = curGDR.fetchall()

        for v in l:
            ListPoids.append(v[0])  

        PoidsDico[val[0]] = ListPoids

    
    return PoidsDico,NombreDico

def poidsProd(sousCat,ListPoids,n):
    
    n = n[sousCat]

    pos = ceil((n+1)/2) #formule de la position pour trouver la médiane

    if not(ListPoids[sousCat]):
        Poids = 'NR'
        return Poids

    Poids = str(ListPoids[sousCat][int(pos)-1]) #selectionne le poids selon la position de la médiane
    
    return Poids.replace("(", "").replace(",", "").replace(")", "")

def CycleCsv():
    Csv=csv.reader(open('Cycles.csv', "r", encoding='latin-1'), delimiter=',')
    next(Csv, None)
    MotClé = []

    for row in Csv:
        MotClé.append(row[0])

    return MotClé

def JeuCsv():
    Csv=csv.reader(open('JeuEtJouet.csv', "r", encoding='latin-1'), delimiter=',')
    next(Csv, None)
    MotClé = []

    for row in Csv:
        MotClé.append(row[0])

    return MotClé

def InsertProduit():
    ID_Struc = IDStructure()

    curGDR.execute('SELECT Produit.Nombre, Produit.Poids, Flux.Flux, Etat_produit.Désignation, Categorie.Désignation, Produit.IDArrivage, Sous_Categorie.Désignation FROM Flux, Produit, Etat_produit, Categorie, Sous_Categorie WHERE Produit.IDFlux = Flux.IDFlux AND Etat_produit.IDEtat_produit = Produit.IDEtat_produit AND Produit.IDCatégorie = Categorie.IDCatégorie AND Produit.IDSous_Catégorie = Sous_Categorie.IDSous_Catégorie')
    List = curGDR.fetchall()

    curSQL.execute("SELECT max(Id_Arrivage) FROM Produit")
    test=curSQL.fetchone() [0]
    for row in List:
        Nombre = row[0]
        Poids = row[1]
        Flux = row[2]
        Flux = Flux.upper().replace("'", "").replace("-", "").replace("/", "").replace(" ", "")
        IDFlux = flux(Flux)
        Orient = row[3]
        Categorie = row[4]
        Categorie = Categorie.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
        SousCategorie = row[6]
        SousCategorie = SousCategorie.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
        IDCat = cat(Categorie)
        if IDCat == '9' :
            IDCat = souscatCycle(SousCategorie)
        if IDCat != '7' and IDCat != '5' and IDCat != '1':
            IDCat = souscatJeu(SousCategorie, IDCat)
        if not test:
            ID_arr = row[5]
        else:
            ID_arr = test + row[5]
        curSQL.execute('INSERT INTO Produit (Orientation, Id_catégorie, Id_Flux, nombre, Id_recyclerie, Poids, Id_arrivage) VALUES (?,?,?,?,?,?,?)', (Orient, IDCat, IDFlux, Nombre, ID_Struc, Poids, ID_arr))

def InsertTournee():
    ID_Struc = IDStructure()

    curSQL.execute("SELECT Tournée FROM Tournée WHERE Id_Recyclerie = ?", (ID_Struc,))
    TourSQL = curSQL.fetchall()
    TourTab = list()
    for row in TourSQL:
        verif = row[0].replace("'", "").replace("-", " ")
        TourTab.append(verif)

    curGDR.execute('SELECT Intitulé FROM Tournee')
    List = curGDR.fetchall()

    for row in List:
        Tournee = row[0]
        Tournee = Tournee.replace("'", "").replace("-"," ")
        if Tournee not in TourTab:
            curSQL.execute('INSERT INTO Tournée (Tournée, Id_recyclerie) VALUES (?,?)', (Tournee, ID_Struc))
          
def InsertVente():
    ID_Struc=IDStructure()
    
    annee = datetime.date.today().year
    curGDR.execute("SELECT MAX(to_char(Date,'YYYY/MM/DD')) FROM vente_magasin")
    MaxDate = curGDR.fetchone()[0].replace("/","")

    date = verifAnnee(MaxDate, annee)

    curSQL.execute("SELECT Max(date) FROM Vente WHERE Id_recyclerie = ?", (ID_Struc,))
    verifDate = curSQL.fetchone()[0]

    if verifDate == None:
        verifDate = '00000000'

    verifDate = verifDate.replace("/","")

    curSQL.execute("SELECT Id_Insee, Commune FROM Insee")
    insee = curSQL.fetchall()

    ListePoids,n = ListPoids()
    
    curGDR.execute("SELECT to_char(date,'YYYY/MM/DD'),code_postal,ville,montant_total,tauxremise from vente_magasin WHERE date > %s AND date < %s ORDER BY code_postal,ville" %\
            (verifDate,date))
    b = curGDR.fetchall()
    for venteorigine in b :
        
        if venteorigine[2] == '' or venteorigine[2] == '-1':
            ville = 'NR'
            IdInsee = 'NR'
        else:
            ville=venteorigine[2]
            ville = unidecode.unidecode(ville).upper().replace("-", " ")
            ville = ville.replace(" ST ", " SAINT ").replace(" STE ", " SAINTE ")
            ville = ville.strip(' ')
            if ville.find("'") :
                ville=ville.replace("'"," ")
            if ville[:3] == "ST " or ville[:4] =="STE ":
                ville = ville.replace("ST ", "SAINT ").replace("STE ","SAINTE ")
            IdInsee = Ville(ville, insee)
            
        curSQL.execute("INSERT INTO Vente (Id_insee, Date, Code_Postal, Commune, Montant_total, TauxRemise, Id_recyclerie) VALUES('%s', '%s','%s','%s','%s','%s','%s') " %\
                    (IdInsee,venteorigine[0],venteorigine[1],ville,venteorigine[3],venteorigine[4], ID_Struc))
        
    curSQL.execute("SELECT max(Id_vente) FROM Vente")
    venteoriginemax=curSQL.fetchone() [0]
    curGDR.execute("SELECT Categorie.Désignation,lignes_vente.montant,lignes_vente.poids,lignes_vente.tauxtva,lignes_vente.montanttva, Flux.Flux, lignes_vente.IDProduit, lignes_vente.IDSous_Catégorie, Sous_Categorie.Désignation FROM lignes_vente, Sous_Categorie, Flux, Categorie WHERE lignes_vente.IDSous_Catégorie = Sous_Categorie.IDSous_Catégorie AND Sous_Categorie.IDFlux = Flux.IDFlux AND Categorie.IDCatégorie = Lignes_Vente.IDCatégorie")
    c=curGDR.fetchall()
    for lignevente in c :
        Categorie = lignevente[0]
        Categorie = Categorie.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
        SousCategorie = lignevente[8]
        SousCategorie = SousCategorie.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
        IDCat = cat(Categorie)
        if IDCat == '9' :
            IDCat = souscatCycle(SousCategorie)
        if IDCat != '7' and IDCat != '5' and IDCat != '1':
            IDCat = souscatJeu(SousCategorie, IDCat)
        Flux = lignevente[5]
        Flux = Flux.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
        IDFlux = flux(Flux)
        poids = lignevente[2]
        if poids == 0:
            poids = poidsProd(lignevente[7],ListePoids,n)
        curSQL.execute("INSERT INTO Lignes_vente (Id_catégorie,Montant,Poids,Taux_tva,Montant_tva,Id_vente, Id_Flux) values ('%s','%s','%s','%s','%s','%s','%s')" %\
                        (IDCat,lignevente[1],poids,lignevente[3],lignevente[4],venteoriginemax, IDFlux))
    
    curSQL.execute("SELECT max(Id_vente) FROM Vente")
    venteoriginemax=curSQL.fetchone() [0]
    curGDR.execute("SELECT Categorie.Désignation,lignes_vente.montant,lignes_vente.poids,lignes_vente.tauxtva,lignes_vente.montanttva,lignes_vente.IDProduit, lignes_vente.IDSous_Catégorie, Sous_Categorie.Désignation FROM lignes_vente, Sous_Categorie, Categorie WHERE Sous_Categorie.IDFlux = 0 AND lignes_vente.IDSous_Catégorie = Sous_Categorie.IDSous_Catégorie AND Categorie.IDCatégorie = Lignes_Vente.IDCatégorie")
    c=curGDR.fetchall()
    for lignevente in c :
        Categorie = lignevente[0]
        Categorie = Categorie.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
        SousCategorie = lignevente[7]
        SousCategorie = SousCategorie.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
        IDCat = cat(Categorie)
        if IDCat == '9' :
            IDCat = souscatCycle(SousCategorie)
        if IDCat != '7' and IDCat != '5' and IDCat != '1':
            IDCat = souscatJeu(SousCategorie, IDCat)
        poids = lignevente[2]
        if poids == 0:
            poids = poidsProd(lignevente[6],ListePoids,n)
        curSQL.execute("INSERT INTO Lignes_vente (Id_catégorie,Montant,Poids,Taux_tva,Montant_tva,Id_vente, Id_Flux) values ('%s','%s','%s','%s','%s','%s','%s')" %\
                        (IDCat,lignevente[1],poids,lignevente[3],lignevente[4],venteoriginemax, 0)) 

    curGDR.execute("SELECT Categorie.Désignation,lignes_vente.montant,lignes_vente.poids,lignes_vente.tauxtva,lignes_vente.montanttva,lignes_vente.IDProduit FROM lignes_vente, Categorie WHERE Categorie.IDCatégorie = Lignes_Vente.IDCatégorie AND IDSous_Categorie = 0")
    c=curGDR.fetchall()
    for lignevente in c :
        Categorie = lignevente[0]
        Categorie = Categorie.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
        IDCat = cat(Categorie)
        poids = lignevente[2]
        if poids == 0:
            poids = 'NR'
        curSQL.execute("INSERT INTO Lignes_vente (Id_catégorie,Montant,Poids,Taux_tva,Montant_tva,Id_vente, Id_Flux) values ('%s','%s','%s','%s','%s','%s','%s')" %\
                        (IDCat,lignevente[1],poids,lignevente[3],lignevente[4],venteoriginemax, 0))

def Ville(ville, insee):

    ville = ville.upper().replace("-", " ").replace("'", "")
    ville = unidecode.unidecode(ville)
    
    for row in insee:
        if ville == row[1]:
            IdInsee = row[0]
            break
        else:
            IdInsee = 0

    return IdInsee

def catDico():
    test=csv.reader(open('catégorie.csv', "r", encoding='latin-1'), delimiter=',')
    next(test, None)
    MotClé = {}

    for row in test:
        MotClé[row[0]] = row[1]

    return MotClé

def fluxDico():
    test=csv.reader(open('flux.csv', "r", encoding='latin-1'), delimiter=',')
    next(test, None)
    MotClé = {}

    for row in test:
        MotClé[row[0]] = row[1]

    return MotClé

def cat(cat):
    
    MotClé = catDico()
    cat = unidecode.unidecode(cat)
    IDCat = 12
    for mot, Id in MotClé.items():
        if cat.find(mot) != -1:
            IDCat = Id       
            break
        else:
            continue

    return IDCat

def souscatCycle(souscat):
    MotCycle = CycleCsv()

    IDCat = 9
    souscat = unidecode.unidecode(souscat)
    for val in MotCycle:
        if souscat.find(val) != -1:
            IDCat = 11    
            break
        else:
            continue

    return IDCat

def souscatJeu(souscat, IDcat):
    MotCycle = JeuCsv()

    IDCat = IDcat
    souscat = unidecode.unidecode(souscat)
    for val in MotCycle:
        if souscat.find(val) != -1:
            IDCat = 7    
            break
        else:
            continue

    return IDCat

def flux(flux):
    MotClé = fluxDico()

    IDFlux = 9
    for mot, Id in MotClé.items():
        if flux.find(mot) != -1:
            IDFlux = Id
            break
        else:
            continue

    return IDFlux

def correct():
    curGDR.execute("SELECT Ville FROM Organisation")
    Ville = curGDR.fetchone()
    Ville = Ville[0].upper().replace("-"," ").replace("'", "")
    curSQL.execute("SELECT Longitude, Latitude FROM Insee WHERE Commune = ?", (Ville,))
    GPS = curSQL.fetchone()
    ID = IDStructure()
    correction_des_communes.correction(connect, curSQL, ID, GPS[0], GPS[1])

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

#######################################################################
##INSERTION DES DONNEES##
try:
    curGDR.execute("SELECT RaisonSociale FROM Organisation")
    RecyclerieNomGDR = curGDR.fetchone()

    curSQL.execute("INSERT OR IGNORE INTO Organisation (Recyclerie) VALUES (?) ", (RecyclerieNomGDR))
    print("Insertion en cours...")
    InsertTournee()
    print("Tournée insérée")
    insertComm()
    print("Commune insérée")
    correct()
    insertArr()
    print("Arrivage inséré")
    InsertProduit()
    print("Produit inséré")
    InsertVente()
    print("Vente inséré")

    #connect.commit()
    print("insertion des données effectué")
except IOError:
    print(IOError)

connect.close()
conn.close()
curGDR.close()

