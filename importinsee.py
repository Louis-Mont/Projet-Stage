import xlrd 

fichier = "classeur.xlsx"
classeur=xlrd.open_workbook(fichier)
nomfeuilles=classeur.sheet_names()
listing_insee={}

def dico_insee() :
    x,y=0,3
    feuille=classeur.sheet_by_name(nomfeuilles[0])
    for lignes in range(feuille.nrows) :
        try :
            a=feuille.cell_value(x,0)
            if a.find("'") :
                a=a.replace("'","\\'")
            a=a.strip().lower()
            b=feuille.cell_value(x,1)
            c=feuille.cell_value(x,2)
            listing_insee[a]= (b,c)
            x+=1
        except :
            print("fin")

def lec_dico(mot):
    for k,v in listing_insee.items() :
        if k == mot :
            print(k,v[0],v[1])



