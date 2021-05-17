import pypyodbc
import datetime
import csv
import sqlite3
import unidecode
import correction_des_communes
from math import *

#-----------------------------------------------------------------------------------------#
#fontions principales:

#fonction qui va chercher l'id de la structure dans la base sql
def IDStructure():
    curSQL.execute("SELECT Id_Recyclerie FROM Organisation WHERE Recyclerie = ?", (RecyclerieNomGDR))
    id_Struc = curSQL.fetchone()
    for row in id_Struc:
        id_Struc = row

    return id_Struc

#fonction qui va retourner la date selon la date max que contient la base sql et l'année actuelle
def verifAnnee(MaxDate, annee):
    date = str(annee) + '0101'
    if MaxDate > date:
        return date
    else:
        date = str(annee-1) + '0101'
        return date

#fonction d'insertion des communes
def insertComm():
    ID_Struc = IDStructure()
    curSQL.execute("SELECT Commune FROM Commune WHERE Id_Recyclerie = ?", (ID_Struc,))
    CommSQL = curSQL.fetchall()
    CommTab = list()
    for row in CommSQL:
        verif = row[0].upper()
        verif = verif.replace("'", "").replace("-", " ")
        CommTab.append(verif)

    curGDR.execute("SELECT Commune,CodePostal, Déchèterie, Apport, Domicile FROM Commune WHERE EnrActif = 1 AND (CodePostal = '' or CodePostal != '') ORDER BY CodePostal")
    CommuneList1 = curGDR.fetchall()
    for row in CommuneList1:
        Commune = row[0].upper()
        Commune = unidecode.unidecode(Commune)
        Commune = Commune.replace("'", "").replace("-", " ")
        Commune = Commune.strip(' ').replace(" ST ", " SAINT ").replace(" STE ", " SAINTE ")
        if Commune[:3] == "ST " or Commune[:4] =="STE ":
            Commune = Commune.replace("ST ", "SAINT ").replace("STE ","SAINTE ")
        CodePostal = row[1]
        if CodePostal == '':
            CodePostal = 'NR'
        codePost = str(CodePostal)
        codePost = codePost[:2] + '%'
        curSQL.execute("SELECT Id_Insee FROM Insee WHERE Commune LIKE '%s' AND Code LIKE '%s'"%\
                            (Commune+'%', codePost))
        id_insee = curSQL.fetchone()
        if id_insee == None:
            curSQL.execute("SELECT Id_Insee FROM Insee WHERE Commune LIKE '%s'"%\
                                (Commune+'%'))  
            id_insee = curSQL.fetchone()
        id_insee = str(id_insee).replace('(','').replace(')','').replace(',','')
        if Commune not in CommTab:
            curSQL.execute("INSERT INTO Commune (Commune, Code_postal, Id_Recyclerie, Id_Insee, Apport, Déchèterie, Domicile) VALUES('%s','%s','%s','%s','%s','%s','%s')" %\
                                    (Commune, CodePostal, ID_Struc, id_insee, row[3], row[2], row[4]))

#fonction d'insertion des arrivages
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

    curGDR.execute("SELECT to_char(Date,'YYYY/MM/DD'), Origine, Poids_total, IDArrivage FROM Arrivage WHERE IDCommune = 0 AND IDTournée = 0 AND date > %s AND date < %s " %\
            (verifDate,date))
    ArrivList3 = curGDR.fetchall()
    for row in ArrivList3:
        curSQL.execute("INSERT INTO Arrivage (Date, Id_commune, origine, poids_total, Id_recyclerie, Id_tournée) VALUES (?,?,?,?,?,?)", (row[0], 0, row[1], row[2], ID_Orga, 0))

        curSQL.execute("select max(Id_arrivage) from Arrivage")
        arrivagemax=curSQL.fetchone() [0]

        InsertProduit(ID_Orga, row[3], arrivagemax)

    curGDR.execute("SELECT to_char(Date,'YYYY/MM/DD'), Origine, Poids_total, Tournee.Intitulé, IDArrivage FROM Arrivage, Tournee WHERE IDCommune = 0 AND Tournee.IDTournée = Arrivage.IDTournée AND date > %s AND date < %s " %\
            (verifDate,date))
    ArrivList2 = curGDR.fetchall()
    for row in ArrivList2:
        tournée = row[3]
        tournée = tournée.replace("'", "").replace("-", " ")
        ID_tour = TournéeDic[tournée]
        curSQL.execute("INSERT INTO Arrivage (Date, Id_commune, origine, poids_total, Id_recyclerie, Id_tournée) VALUES (?,?,?,?,?,?)", (row[0], 0, row[1], row[2], ID_Orga, ID_tour))

        curSQL.execute("select max(Id_arrivage) from Arrivage")
        arrivagemax=curSQL.fetchone() [0]

        InsertProduit(ID_Orga, row[4], arrivagemax)

    curGDR.execute("SELECT to_char(Date,'YYYY/MM/DD'), Origine, Poids_total, Commune.Commune, IDArrivage FROM Arrivage, Commune WHERE Commune.IDCommune = Arrivage.IDCommune AND date > %s AND date < %s " %\
            (verifDate,date))
    ArrivList = curGDR.fetchall()
    for row in ArrivList:
        Comm = row[3]
        Comm = Comm.upper().replace("'","").replace("-"," ")
        Comm = unidecode.unidecode(Comm)
        Comm = Comm.strip(' ')
        Comm = Comm.replace(" ST ", " SAINT ").replace(" STE ", " SAINTE ")
        if Comm[:3] == "ST " or Comm[:4] =="STE ":
            Comm = Comm.replace("ST ", "SAINT ").replace("STE ","SAINTE ")
        ID_comm = CommTab[Comm]
        curSQL.execute("INSERT INTO Arrivage (Date, Id_commune, origine, poids_total, Id_recyclerie, Id_tournée) VALUES (?,?,?,?,?,?)", (row[0], ID_comm, row[1], row[2], ID_Orga, 0))

        curSQL.execute("select max(Id_arrivage) from Arrivage")
        arrivagemax=curSQL.fetchone() [0]

        InsertProduit(ID_Orga, row[4], arrivagemax)

#fonction qui va retourner une liste des poids et une autre liste pour le nombre des produits selon les sous_catégories de la base GDR
def ListPoids():

    curGDR.execute("SELECT IDSous_Catégorie FROM Sous_Categorie")
    IDSousCat = curGDR.fetchall() # On récupère les identifiants de sous_catégorie
    
    PoidsDico = {} # premier dictionnaire repertoriant les poids selon la sous_catégorie
    NombreDico = {} # deuxième dictionnaire repertoriant le nombre de produit selon la sous_catégorie
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

def InsertProduit(ID_Struc, IDArrivage, Arrivagemax):
    curGDR.execute("SELECT Produit.Nombre, Produit.Poids, Flux.Flux, Etat_produit.Désignation, Categorie.Désignation, Produit.IDArrivage, Sous_Categorie.Désignation FROM Flux, Produit, Etat_produit, Categorie, Sous_Categorie WHERE Produit.IDFlux = Flux.IDFlux AND Etat_produit.IDEtat_produit = Produit.IDEtat_produit AND Produit.IDCatégorie = Categorie.IDCatégorie AND Produit.IDSous_Catégorie = Sous_Categorie.IDSous_Catégorie AND IDArrivage = '%s'"%\
                                (IDArrivage))
    List = curGDR.fetchall()

    for row in List:
        Flux = row[2]
        Flux = Flux.upper().replace("'", "").replace("-", "").replace("/", "").replace(" ", "")
        IDFlux = flux(Flux)
        Categorie = row[4]
        Categorie = Categorie.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
        SousCategorie = row[6]
        SousCategorie = SousCategorie.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
        IDCat = cat(Categorie)
        if IDCat == '9' :
            IDCat = souscatCycle(SousCategorie)
        if IDCat != '7' and IDCat != '5' and IDCat != '1':
            IDCat = souscatJeu(SousCategorie, IDCat)
        curSQL.execute('INSERT INTO Produit (Orientation, Id_catégorie, Id_Flux, nombre, Id_recyclerie, Poids, Id_arrivage) VALUES (?,?,?,?,?,?,?)', (row[3], IDCat, IDFlux, row[0], ID_Struc, row[1], Arrivagemax))

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
        
    curGDR.execute("SELECT to_char(date,'YYYY/MM/DD'),code_postal,ville,montant_total,tauxremise, IDVente_Magasin from vente_magasin WHERE date > %s AND date < %s ORDER BY code_postal,ville" %\
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
            IdInsee = Ville(ville, insee, venteorigine[1])
                
        curSQL.execute("INSERT INTO Vente (Id_insee, Date, Code_Postal, Commune, Montant_total, TauxRemise, Id_recyclerie) VALUES('%s', '%s','%s','%s','%s','%s','%s') " %\
                        (IdInsee,venteorigine[0],venteorigine[1],ville,venteorigine[3],venteorigine[4], ID_Struc))

        curSQL.execute("select max(Id_Vente) from Vente ")
        venteoriginemax=curSQL.fetchone() [0]

        curGDR.execute("SELECT Categorie.Désignation,lignes_vente.montant,lignes_vente.poids,lignes_vente.tauxtva,lignes_vente.montanttva, Flux.Flux, lignes_vente.IDProduit, lignes_vente.IDSous_Catégorie, Sous_Categorie.Désignation FROM lignes_vente, Sous_Categorie, Flux, Categorie WHERE lignes_vente.IDSous_Catégorie = Sous_Categorie.IDSous_Catégorie AND Sous_Categorie.IDFlux = Flux.IDFlux AND Categorie.IDCatégorie = Lignes_Vente.IDCatégorie AND IDVente_Magasin = '%s'" %\
                                (venteorigine[5]))
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
        
        curGDR.execute("SELECT Categorie.Désignation,lignes_vente.montant,lignes_vente.poids,lignes_vente.tauxtva,lignes_vente.montanttva,lignes_vente.IDProduit, lignes_vente.IDSous_Catégorie, Sous_Categorie.Désignation FROM lignes_vente, Sous_Categorie, Categorie WHERE Sous_Categorie.IDFlux = 0 AND lignes_vente.IDSous_Catégorie = Sous_Categorie.IDSous_Catégorie AND Categorie.IDCatégorie = Lignes_Vente.IDCatégorie AND IDVente_Magasin = '%s'" %\
                                (venteorigine[5]))
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

        curGDR.execute("SELECT Categorie.Désignation,lignes_vente.montant,lignes_vente.poids,lignes_vente.tauxtva,lignes_vente.montanttva,lignes_vente.IDProduit FROM lignes_vente, Categorie WHERE Categorie.IDCatégorie = Lignes_Vente.IDCatégorie AND IDSous_Categorie = 0 AND IDVente_Magasin = '%s'" %\
                                (venteorigine[5]))
        c=curGDR.fetchall()
        for lignevente in c :
            
            Categorie = lignevente[0]
            Categorie = Categorie.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
            IDCat = cat(Categorie)
            poids = lignevente[2]
            if poids == 0:
                poids = 'NR'
            curSQL.execute("INSERT INTO Lignes_vente (Id_catégorie,Montant,Poids,Taux_tva,Montant_tva,Id_vente, Id_Flux) values ('%s','%s','%s','%s','%s','%s','%s')" %\
                            (IDCat,lignevente[1],poids,lignevente[3],lignevente[4],venteoriginemax, 'NR'))
        
        curGDR.execute("SELECT Categorie.Désignation,lignes_vente.montant,lignes_vente.poids,lignes_vente.tauxtva,lignes_vente.montanttva,lignes_vente.IDProduit FROM lignes_vente, Categorie WHERE Lignes_Vente.IDSous_Catégorie = 0 and  Categorie.IDCatégorie = Lignes_Vente.IDCatégorie AND IDVente_Magasin = '%s'" %\
                                (venteorigine[5]))
        c=curGDR.fetchall()
        for lignevente in c :
            
            Categorie = lignevente[0]
            Categorie = Categorie.upper().replace("'", "").replace("-", "").replace("/", "").strip(" ")
            IDCat = cat(Categorie)
            poids = lignevente[2]
            if poids == 0:
                poids = 'NR'
            curSQL.execute("INSERT INTO Lignes_vente (Id_catégorie,Montant,Poids,Taux_tva,Montant_tva,Id_vente, Id_Flux) values ('%s','%s','%s','%s','%s','%s','%s')" %\
                            (IDCat,lignevente[1],poids,lignevente[3],lignevente[4],venteoriginemax, 'NR'))

def Ville(ville, insee, code_postal):

    ville = ville.upper().replace("-", " ").replace("'", "")
    ville = unidecode.unidecode(ville)
    
    codePost = str(code_postal)[:2]
    for row in insee:
        if ville == row[1] and row[0].find(codePost) != -1:
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

#fonction pour réattribuer les produits de la catégorie "SPORTS ET LOISIRS" ayant comme sous_catégories "CYCLES", "VELO" ou autre à la catégorie "CYCLES"
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

# fonction pour réattribuer les catégories qui ont des sous_catégories ayant comme désignation "JEUX", "JOUETS" ou autre qui n'appartiennent pas à la catégorie "JEUX ET JOUETS" à la catégorie "JEUX ET JOUETS"
def souscatJeu(souscat, IDcat):
    MotCycle = JeuCsv()

    souscat = unidecode.unidecode(souscat) 
    for val in MotCycle:
        if souscat.find(val) != -1 and souscat != 'JEUNESSE': # si la sous_catégorie est identifié dans la liste et n'a pas pour désignation "JEUNESSE" alors on lui attribue la catégorie "JEUX ET JOUETS"
            IDcat = 7    
            break
        else:
            continue

    return IDcat

#fonction pour réattribuer les flux des bases GDR aux flux finaux de la base SQL
def flux(flux):
    MotClé = fluxDico()

    IDFlux = 1 # identifiant par defaut represantant le flux "TOUT VENANT" 
    for mot, Id in MotClé.items(): # recherche dans la liste si le flux existe et lui attribue un nouvel identifiant sinon prend l'identifiant par defaut
        if flux.find(mot) != -1:
            IDFlux = Id
            break
        else:
            continue

    return IDFlux

#fonction qui corrige les communes sans code insee
def correct():
    curGDR.execute("SELECT Ville FROM Organisation") #On va chercher la ville où la structure se situe
    Ville = curGDR.fetchone() [0]
    Ville = Ville.upper().replace("-"," ").replace("'", "")
    Ville = unidecode.unidecode(Ville)
    curSQL.execute("SELECT Longitude, Latitude FROM Insee WHERE Commune = '%s'"%\
                        (Ville)) #On va chercher les coordonnées de la ville
    GPS = curSQL.fetchone()
    #si il n'ya pas de coordonnée alors le programme va chercher dans le fichier des anciennes communes si la commune s'y trouve et retourne la nouvelle commune 
    if GPS == None:
        fichier = "ancienne_commune.csv"
        Csv=csv.reader(open(fichier, "r", encoding='latin-1'), delimiter=',')
        next(Csv, None)
        listing_insee={}
        for row in Csv:
            OldComm=row[3]
            if OldComm.find("'") :
                OldComm=OldComm.replace("'","")
            OldComm=unidecode.unidecode(OldComm.strip().upper()).replace("-", " ")
            NewComm=unidecode.unidecode(row[1].strip().upper()).replace("-", " ")
            listing_insee[OldComm] = NewComm
            for k,v in listing_insee.items() :
                if k == Ville :
                    Ville = v
                    break
        curSQL.execute("SELECT Longitude, Latitude FROM Insee WHERE Commune = '%s'"%\
                        (Ville))
        GPS = curSQL.fetchone()
    ID = IDStructure()
    correction_des_communes.correction(connect, curSQL, ID, GPS[0], GPS[1]) #appel d'une fonction du fichier correction_des_communes

#--------------------------------------------------------------------------------------------------------------
# Code principal

print("connexion en cours à la base de la recyclerie à extraire")
conn = pypyodbc.connect(DSN='Extraction')  # initialisation de la connexion au serveur odbc pour se connecter à GDR
curGDR = conn.cursor()
print("connexion ok\n")

print("connexion en cours à la grosse base de données")
connect = sqlite3.connect("finale.db") # initialisation de la connexion à la base sql
curSQL = connect.cursor()
print("connexion ok\n")

#######################################################################
##INSERTION DES DONNEES##
try:
    curGDR.execute("SELECT RaisonSociale FROM Organisation") # récupère le nom de l'organisation
    RecyclerieNomGDR = curGDR.fetchone()

    curSQL.execute("INSERT OR IGNORE INTO Organisation (Recyclerie) VALUES (?) ", (RecyclerieNomGDR)) # insère le nom de l'organisation dans la nouvelle base ou non si le nom existe déjà
    print("Insertion en cours...")
    InsertTournee()
    print("Tournée insérée  ===> 25%")
    insertComm()
    print("Commune insérée  ===> 50%")
    correct()
    insertArr()
    print("Arrivage inséré  ===> 75%")
    InsertVente()
    print("Vente inséré  ===> 100%")
    connect.commit()
    print("insertion des données effectué")
except IOError:
    print(IOError)

connect.close()
conn.close()
curGDR.close()

