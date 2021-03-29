import pypyodbc
import xlrd
import datetime
from xlwt import Workbook, easyxf
import unidecode

print("connexion en cours")
conn = pypyodbc.connect('DSN=test;')  # initialisation de la connexion au serveur
cur = conn.cursor()
print("connexion ok")

jourj = datetime.date.today()
annuel = jourj - datetime.timedelta(days=366)
annuel = str(annuel)
# annuel = annuel.replace("-","")
annuel = 20190101
fin = 20191231
print(annuel)

classeurinsee = xlrd.open_workbook("Extraction de données GDR.xlsx")
nomfeuilles = classeurinsee.sheet_names()
numfeuille = 2
feuilleinsee = classeurinsee.sheet_by_name(nomfeuilles[numfeuille])
codeinsee = {}

classeur = Workbook()
style = easyxf("font: bold 1; alignment: horizontal left")

synthesecollectes = ["Poids total", "nombre d'arrivages", "nombre de produits", "% du poids", "% du nombre de produits",
                     "poids moyen par arrivage", "nombre moyen de produits par arrivage", "poids moyen par produit"]
syntheseventes = ["Poids", "chiffre d'affaires (en €)", "nombre d'articles", "% du poids", "% du chiffre d'affaires",
                  "poids moyen par produit", "chiffre d'affaires moyen par produit"]
syntheseventescommune = ["Poids", "nombre de ventes", "nombre de produits", "chiffre d'affaires (en €)", "% du poids",
                         "% du chiffre d'affaires", "% du nombre de ventes", "poids moyen par produit",
                         "chiffre d'affaires moyen par produit "]
syntheseventescat = ["Poids", "nombre de ventes", "chiffre d'affaires (en €)", "nombre d'articles", "% du poids",
                     "% du chiffre d'affaires", "poids moyen par produit", "chiffre d'affaires moyen par produit "]

print("détection des flux")
listeflux = {}
cur.execute("SELECT idflux, flux FROM flux ")
listingflux = cur.fetchall()
for i in listingflux:
    listeflux[i[0]] = i[1]
cur.execute("SELECT flux, flux FROM flux ")
listingflux1 = cur.fetchall()
cur.execute("SELECT Désignation FROM Categorie")
cats = cur.fetchall()
cur.execute("SELECT Commune FROM Commune WHERE apport = 1")
comapport = cur.fetchall()
cur.execute("SELECT Commune FROM Commune WHERE domicile = 1")
comrdv = cur.fetchall()
cur.execute("SELECT Désignation FROM Etat_produit ")
orient = cur.fetchall()
print("détection des flux terminée")


def remplacement(mot):
    mot = mot.replace("'", "\\'")
    return (mot)


def calculsyntheseventes(x, ventes, filtre, objets):
    for i in objets:
        if i[0].find("'"):
            y = remplacement(i[0])
        if filtre == "comm":
            cur.execute(
                "SELECT SUM(Nombre), SUM(ROUND(poids,2)) ,SUM( ROUND( ( Lignes_Vente.Montant - ROUND( ( ( Lignes_Vente.Montant * Vente_Magasin.TauxRemise ) /  100) ,  3) ) ,  3) )  , COUNT(distinct(lignes_vente.idvente_magasin)) FROM Lignes_vente INNER JOIN Vente_magasin  ON lignes_vente.idvente_magasin=vente_magasin.idvente_magasin WHERE vente_magasin.date BETWEEN %s AND %s AND vente_magasin.ville = '%s' " % \
                (annuel, fin, y))
            poidsligne = cur.fetchone()
        elif filtre == "Catégories":
            cur.execute(
                "SELECT SUM(Nombre), SUM(ROUND(poids,2)),SUM( ROUND( ( Lignes_Vente.Montant - ROUND( ( ( Lignes_Vente.Montant * Vente_Magasin.TauxRemise ) /  100) ,  3) ) ,  3) ) , COUNT(distinct(lignes_vente.idvente_magasin)) FROM Lignes_vente INNER JOIN Vente_magasin  ON lignes_vente.idvente_magasin=vente_magasin.idvente_magasin INNER JOIN categorie ON  LIgnes_vente.idcatégorie=categorie.idcatégorie  WHERE vente_magasin.date BETWEEN %s AND %s AND categorie.Désignation = '%s' " % \
                (annuel, fin, y))
            poidsligne = cur.fetchone()
        elif filtre == "Flux":
            cur.execute("SELECT idflux FROM Flux WHERE Flux.flux = '%s' " % \
                        (y))
            fluxid = cur.fetchone()[0]
            cur.execute(
                "SELECT SUM(Lignes_vente.Nombre), SUM(ROUND(Lignes_vente.poids,2)) ,SUM( ROUND( ( Lignes_Vente.Montant - ROUND( ( ( Lignes_Vente.Montant * Vente_Magasin.TauxRemise ) /  100) ,  3) ) ,  3) ) ,Count(distinct(lignes_vente.IDvente_magasin)) FROM Lignes_vente INNER JOIN Vente_magasin  ON lignes_vente.idvente_magasin=vente_magasin.idvente_magasin INNER JOIN sous_categorie ON lignes_vente.idsous_catégorie=sous_categorie.idsous_catégorie WHERE vente_magasin.date > %s AND sous_categorie.idflux = %s " % \
                (annuel, fluxid))
            poidsligne = cur.fetchone()
        feuille3.write(x, 0, i[0])
        vv = ventes
        if poidsligne[0] and poidsligne[2]:
            feuille3.write(x, 1, poidsligne[1])
            feuille3.write(x, 2, poidsligne[3])
            feuille3.write(x, 3, poidsligne[0])
            feuille3.write(x, 4, poidsligne[2])
            feuille3.write(x, 5, int(poidsligne[1]) / vv[1] * 100)
            feuille3.write(x, 6, int(poidsligne[2]) / vv[2] * 100)
            feuille3.write(x, 8, poidsligne[1] / poidsligne[0])
        else:
            nombre = [1, 2, 3, 4, 5, 6, 7, 8]
            for chiffre in nombre:
                feuille3.write(x, chiffre, "aucune")
        x += 1
    if filtre == "Flux":
        cur.execute(
            "SELECT SUM(Lignes_vente.Nombre), SUM(ROUND(Lignes_vente.poids,2)) ,SUM( ROUND( ( Lignes_Vente.Montant - ROUND( ( ( Lignes_Vente.Montant * Vente_Magasin.TauxRemise ) /  100) ,  3) ) ,  3) ) ,Count(distinct(lignes_vente.IDvente_magasin)) FROM Lignes_vente INNER JOIN Vente_magasin  ON lignes_vente.idvente_magasin=vente_magasin.idvente_magasin INNER JOIN Sous_categorie ON Lignes_vente.idsous_catégorie = sous_categorie.idsous_catégorie  WHERE vente_magasin.date BETWEEN %s AND %s and idproduit = 0 and sous_categorie.idflux = 0 " % \
            (annuel, fin))
        poidsligne = cur.fetchone()
        vv = ventes
        if poidsligne[0] and poidsligne[2]:
            feuille3.write(x, 1, poidsligne[1])
            feuille3.write(x, 2, poidsligne[3])
            feuille3.write(x, 3, poidsligne[0])
            feuille3.write(x, 4, poidsligne[2])
            feuille3.write(x, 5, int(poidsligne[1]) / vv[1] * 100)
            feuille3.write(x, 6, int(poidsligne[2]) / vv[2] * 100)
            feuille3.write(x, 8, poidsligne[1] / poidsligne[0])
        else:
            nombre = [1, 2, 3, 4, 5, 6, 7, 8]
            for chiffre in nombre:
                feuille3.write(x, chiffre, "aucune")
        x += 1
    feuille3.write(x, 0, "total")
    feuille3.write(x, 1, ventes[1])
    feuille3.write(x, 4, ventes[2])
    x += 4
    return (x)


def syntheseventes():
    global feuille3
    cur.execute(
        "SELECT COUNT(idvente_magasin) , SUM(ROUND(poids_total,2)) , SUM(ROUND(Montant_total,2)) FROM Vente_magasin WHERE Date BETWEEN %s AND %s " % \
        (annuel, fin))
    ventes = cur.fetchone()
    x, y = 5, 1
    feuille3 = classeur.add_sheet("Synthèse Ventes")
    feuille3.write(1, 5, "Synthèse Ventes", style)
    for item in syntheseventescommune:
        feuille3.write(x, y, item, style)
        y += 1
    feuille3.write(x, 0, "Commune")
    x += 1
    cur.execute("select ville from Vente_Magasin where date BETWEEN %s AND %s group by ville " % \
                (annuel, fin))
    villes = cur.fetchall()
    x = calculsyntheseventes(x, ventes, "comm", villes)
    y = 1
    for item in syntheseventescommune:
        feuille3.write(x, y, item, style)
        y += 1
    feuille3.write(x, 0, "Catégories")
    x += 1
    cur.execute(
        "select Categorie.Désignation from Lignes_vente INNER JOIN Vente_Magasin ON Lignes_vente.idvente_magasin = Vente_magasin.idvente_magasin INNER JOIN Categorie ON Lignes_vente.idcatégorie=categorie.idcatégorie  where vente_magasin.date BETWEEN %s AND %s group by Categorie.désignation " % \
        (annuel, fin))
    vcats = cur.fetchall()
    x = calculsyntheseventes(x, ventes, "Catégories", vcats)
    y = 1
    for item in syntheseventescommune:
        feuille3.write(x, y, item, style)
        y += 1
    feuille3.write(x, 0, "Flux")
    x += 1
    x = calculsyntheseventes(x, ventes, "Flux", listingflux1)


def lecinsee():
    xinsee = 1
    print("lecture des codes insee")
    try:
        for lignes in range(feuilleinsee.nrows):
            cle = feuilleinsee.cell_value(xinsee, 1)
            cle = unidecode.unidecode(cle)
            cle = cle.lower().strip()
            valeur = feuilleinsee.cell_value(xinsee, 0)
            xinsee += 1
            codeinsee[cle] = valeur
    except:
        print("lecture des codes terminée")


def synthesecats(x, arrivages, an, listing, refprod):
    for arrivee in arrivages:
        for objet in listing:
            if objet[0].find("'"):
                objets = remplacement(objet[0])
            if listing == cats:
                cur.execute(
                    "SELECT SUM(ROUND(produit.poids,2)) ,COUNT(DISTINCT(arrivage.idarrivage)), SUM(produit.nombre) FROM produit INNER JOIN Categorie ON produit.IDCatégorie=Categorie.IDCatégorie INNER JOIN arrivage ON produit.IDarrivage = arrivage.idarrivage WHERE arrivage.date BETWEEN %s AND %s AND categorie.Désignation = '%s' AND arrivage.origine NOT LIKE '%s'" % \
                    (annuel, fin, objets, refprod))
            elif listing == listingflux1:
                cur.execute(
                    "SELECT SUM(ROUND(produit.poids,2)) ,COUNT(DISTINCT(arrivage.idarrivage)), SUM(produit.nombre) FROM produit INNER JOIN flux ON produit.IDflux=flux.IDflux INNER JOIN arrivage ON produit.IDarrivage = arrivage.idarrivage WHERE arrivage.date BETWEEN %s AND %s AND flux.Flux = '%s' AND arrivage.origine NOT LIKE '%s' " % \
                    (annuel, fin, objets, refprod))
            elif listing == orient:
                cur.execute(
                    "SELECT SUM(ROUND(produit.poids,2)) ,COUNT(DISTINCT(arrivage.idarrivage)), SUM(produit.nombre) FROM produit INNER JOIN Etat_produit ON produit.IDEtat_produit=Etat_produit.IDEtat_Produit INNER JOIN arrivage ON produit.IDarrivage = arrivage.idarrivage WHERE arrivage.date BETWEEN %s AND %s AND etat_produit.désignation = '%s' AND arrivage.origine NOT LIKE '%s' " % \
                    (annuel, fin, objets, refprod))
            elif listing == comrdv:
                cur.execute(
                    "SELECT SUM(ROUND(produit.poids,2)) ,COUNT(DISTINCT(arrivage.idarrivage)), SUM(produit.nombre) FROM produit INNER JOIN arrivage ON produit.IDarrivage=arrivage.IDarrivage INNER JOIN commune ON arrivage.idcommune = commune.idcommune WHERE arrivage.date BETWEEN %s AND %s AND commune.commune = '%s' AND commune.Domicile = 1 AND arrivage.origine = 'Rendez-Vous' " % \
                    (annuel, fin, objets))
            elif listing == comapport:
                cur.execute(
                    "SELECT SUM(ROUND(produit.poids,2)) ,COUNT(DISTINCT(arrivage.idarrivage)), SUM(produit.nombre) FROM produit INNER JOIN arrivage ON produit.IDarrivage=arrivage.IDarrivage INNER JOIN commune ON arrivage.IDcommune = commune.idcommune WHERE arrivage.date BETWEEN %s AND %s AND commune.commune = '%s' AND commune.Apport = 1 AND arrivage.origine = 'Apport sur Site' " % \
                    (annuel, fin, objets))
            retour = cur.fetchall()
            if not (retour):
                continue
            for info in retour:
                feuille2.write(x, 0, objet[0])
                if info[0]:
                    feuille2.write(x, 1, info[0])
                    feuille2.write(x, 2, info[1])
                    feuille2.write(x, 3, info[2])
                    feuille2.write(x, 4, round(info[0] * 100 / arrivee[0], 2))
                    feuille2.write(x, 5, round(info[2] * 100 / arrivee[1], 2))
                    feuille2.write(x, 6, round(info[0] / info[1], 2))
                    feuille2.write(x, 7, round(info[2] / info[1], 2))
                    feuille2.write(x, 8, round(info[0] / info[2], 2))
                else:
                    nombre = [1, 2, 3, 4, 5, 6, 7, 8]
                    for chiffre in nombre:
                        feuille2.write(x, chiffre, "aucune")
                x += 1
    x += 3
    return (x)


def syntheseorig():
    global feuille2
    refprod = "Réf"
    print("création et remplissage de la feuille synthèses ")
    x, y = 5, 1
    feuille2 = classeur.add_sheet("Synthèse")
    feuille2.write(1, 5, "Synthèse Collecte", style)
    for item in synthesecollectes:
        feuille2.write(4, y, item, style)
        y += 1
    feuille2.write(4, 0, "Origines", style)
    feuille2.write(5, 0, "Apport sur site", style)
    feuille2.write(6, 0, "Rendez-vous", style)
    feuille2.write(7, 0, "Déchèterie", style)
    feuille2.write(8, 0, "Total", style)
    orig = ["Apport sur Site", "Rendez-Vous", "Déchèterie"]
    cur.execute(
        "SELECT SUM(ROUND(produit.poids,2)) , SUM(produit.nombre), COUNT(DISTINCT(arrivage.idarrivage)) , min(arrivage.idarrivage) FROM produit INNER JOIN arrivage ON arrivage.idarrivage = produit.idarrivage WHERE arrivage.date BETWEEN %s AND %s AND arrivage.origine NOT LIKE '%s' " % \
        (annuel, fin, refprod))
    arrivages = cur.fetchall()
    for arrivee in arrivages:
        for tipe in orig:
            cur.execute(
                "SELECT SUM(ROUND(produit.poids,2)) , SUM(produit.nombre), COUNT(DISTINCT(arrivage.idarrivage)), AVG(produit.poids) FROM produit INNER JOIN arrivage ON arrivage.idarrivage = produit.idarrivage WHERE arrivage.date BETWEEN %s AND %s AND arrivage.origine = '%s' " % \
                (annuel, fin, tipe))
            arrivagesorig = cur.fetchall()
            for z in arrivagesorig:
                if z[0]:
                    feuille2.write(x, 1, z[0])
                    feuille2.write(x, 2, z[2])
                    feuille2.write(x, 3, z[1])
                    feuille2.write(x, 4, round(z[0] * 100 / arrivee[0], 2))
                    feuille2.write(x, 5, round(z[1] * 100 / arrivee[1], 2))
                    feuille2.write(x, 6, round(z[0] / z[2], 2))
                    feuille2.write(x, 7, round(z[1] / z[2], 2))
                    feuille2.write(x, 8, round(z[0] / z[1], 2))
                else:
                    nombre = [1, 2, 3, 4, 5, 6, 7, 8]
                    for chiffre in nombre:
                        feuille2.write(x, chiffre, "aucune")
                x += 1
        # on va remplir la ligne total
        feuille2.write(x, 1, arrivee[0])
        feuille2.write(x, 2, arrivee[2])
        feuille2.write(x, 3, arrivee[1])
        feuille2.write(x, 6, round(arrivee[0] / arrivee[2], 2))
        feuille2.write(x, 7, round(arrivee[1] / arrivee[2], 2))
        feuille2.write(x, 8, round(arrivee[0] / arrivee[1], 2))
        x, y = 12, 1
    for item in synthesecollectes:
        feuille2.write(11, y, item, style)
        y += 1
    feuille2.write(11, 0, "Catégories", style)
    x = synthesecats(x, arrivages, annuel, cats, refprod)
    y = 1
    for item in synthesecollectes:
        feuille2.write(x, y, item, style)
        y += 1
    feuille2.write(x, 0, "Flux", style)
    x += 1
    x = synthesecats(x, arrivages, annuel, listingflux1, refprod)
    y = 1
    for item in synthesecollectes:
        feuille2.write(x, y, item, style)
        y += 1
    feuille2.write(x, 0, "Orientation", style)
    x += 1
    x = synthesecats(x, arrivages, annuel, orient, refprod)
    y = 1
    for item in synthesecollectes:
        feuille2.write(x, y, item, style)
        y += 1
    feuille2.write(x, 0, "Commune (Rendez-Vous)", style)
    x += 1
    x = synthesecats(x, arrivages, annuel, comrdv, refprod)
    y = 1
    for item in synthesecollectes:
        feuille2.write(x, y, item, style)
        y += 1
    feuille2.write(x, 0, "Commune (Apport)", style)
    x += 1
    x = synthesecats(x, arrivages, annuel, comapport, refprod)
    syntheseventes()


def remplissagebdcollecte():
    global feuille1
    x, refprod = 5, "Réf%"
    print("saisie des collectes en cours")
    cur.execute(
        "SELECT IDarrivage, to_char(date,'YYYYMMDD') , origine , poids_total , commune.commune FROM arrivage INNER JOIN commune ON arrivage.idcommune=commune.idcommune WHERE date BETWEEN %s AND %s and origine NOT LIKE '%s'" % \
        (annuel, fin, refprod))
    arrivages = cur.fetchall()
    for i in arrivages:
        cur.execute("SELECT SUM(nombre) FROM Produit WHERE IDarrivage = %s " % \
                    (i[0]))
        Nb = cur.fetchone()[0]
        nominsee = str(i[4])
        nominsee = unidecode.unidecode(nominsee)
        nominsee = nominsee.replace(" ", "-").strip().lower()
        cinsee = codeinsee.get(nominsee)
        feuille1.write(x, 1, i[0])
        feuille1.write(x, 2, i[1])
        feuille1.write(x, 3, i[2])
        feuille1.write(x, 4, i[4])
        feuille1.write(x, 5, cinsee)
        feuille1.write(x, 9, Nb)
        feuille1.write(x, 10, i[3])
        x += 1
        cur.execute(
            "SELECT catégorie.Désignation,sous_catégorie.Désignation , flux.flux , produit.nombre,produit.idproduit,produit.poids FROM produit INNER JOIN Catégorie ON Produit.idcatégorie = catégorie.idcategorie INNER JOIN sous_catégorie ON produit.IDsous_catégorie = sous_catégorie.idsous_categorie INNER JOIN flux ON produit.idflux =flux.idflux WHERE produit.idarrivage = %s " % \
            (i[0]))
        produits = cur.fetchall()
        for prod in produits:
            feuille1.write(x, 0, prod[4])
            feuille1.write(x, 6, prod[2])
            feuille1.write(x, 7, prod[0])
            feuille1.write(x, 8, prod[1])
            feuille1.write(x, 9, prod[3])
            feuille1.write(x, 10, prod[5])
            x += 1
    print("saisie des collectes terminée")


def remplissagebdvente():
    global feuille
    x = 5
    print("Saisie des ventes en cours")
    cur.execute(
        "SELECT IDVente_Magasin , idclient , poids_total ,  Montant_total , Code_Postal, ville  FROM vente_magasin WHERE date BETWEEN '%s' AND %s " % \
        (annuel, fin))
    ventes = cur.fetchall()
    for i in ventes:
        cur.execute("SELECT SUM(nombre) FROM Lignes_vente WHERE idvente_magasin = %s " % \
                    (i[0]))
        nb = cur.fetchone()[0]
        if i[1] != 0:
            cur.execute("SELECT ville FROM client WHERE idclient = %s" % \
                        (i[1]))
            infoclient = cur.fetchone()[0]
        else:
            infoclient = "NC"
        if not i[4]:
            cinsee = "NC"
        else:
            nominsee = str(i[5])
            nominsee = unidecode.unidecode(nominsee)
            nominsee = nominsee.replace(" ", "-").strip().lower()
            cinsee = codeinsee.get(nominsee)
        feuille.write(x, 0, i[0])
        feuille.write(x, 1, infoclient)
        if not i[4]:
            feuille.write(x, 2, "NC")
        else:
            feuille.write(x, 2, i[4])
        feuille.write(x, 3, cinsee)
        feuille.write(x, 7, nb)
        feuille.write(x, 8, i[2])
        feuille.write(x, 9, i[3])
        feuille.write(x, 10, i[3])
        x += 1

    print("saisie des ventes terminée")


def prepabdcollecte():
    global feuille1
    feuille1 = classeur.add_sheet("BD_Collecte")
    feuille1.write(1, 5, "Collectes", style)
    feuille1.write(4, 0, "id produit", style)
    feuille1.write(4, 1, "id arrivage", style)
    feuille1.write(4, 2, "Date arrivage", style)
    feuille1.write(4, 3, "Origine arrivage (=rendez-vous, apport sur site ou déchèterie)", style)
    feuille1.write(4, 4,
                   "si déchèterie : nom de la déchèterie; si apport sur site ou rendez-vous = nom de la commune d'origine",
                   style)
    feuille1.write(4, 5, "Code insee", style)
    feuille1.write(4, 6, "Flux (tout-venant, DEEE, textile, eco-mobilier, bois, ferraille…)", style)
    feuille1.write(4, 7, "Catégorie", style)
    feuille1.write(4, 8, "Sous-Catégorie", style)
    feuille1.write(4, 9, "Quantité", style)
    feuille1.write(4, 10, "Poids", style)
    feuille1.write(4, 11, "Affectation", style)


def prepabdvente():
    global feuille
    feuille = classeur.add_sheet("BD_Vente")
    feuille.write(1, 5, "Ventes", style)
    feuille.write(4, 0, "Identifiant unique de la vente", style)
    feuille.write(4, 1, "Commune d'origine du client", style)
    feuille.write(4, 2, "Code postal de la commune d'origine", style)
    feuille.write(4, 3, "si c'est possible : retrouver le code insee de la commune (voir onglet insee)", style)
    feuille.write(4, 4, "Flux (tout-venant, DEEE, textile, eco-mobilier, bois, ferraille…)", style)
    feuille.write(4, 5, "Catégorie", style)
    feuille.write(4, 6, "Sous-Catégorie", style)
    feuille.write(4, 7, "Nombre de produits", style)
    feuille.write(4, 8, "Poids des produits", style)
    feuille.write(4, 9, "Montant HT", style)
    feuille.write(4, 10, "Montant TTC", style)
    feuille.write(4, 11, "IDProduit")


lecinsee()
prepabdvente()
remplissagebdvente()
prepabdcollecte()
remplissagebdcollecte()
syntheseorig()
classeur.save("extractionGDR.xlsx")
conn.close()
