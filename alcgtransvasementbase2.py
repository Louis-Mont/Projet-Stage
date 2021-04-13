import pypyodbc

conn=pypyodbc.connect("DSN=gdr2;")
cur=conn.cursor()

conn2=pypyodbc.connect("DSN=gdr;")
cur2=conn2.cursor()

cur.execute("SELECT indice,to_char(date,'YYYYMMDD'),montant,utilise,saisipar,to_char(dateutilise,'YYYYMMDD'),cast(heure as character(5)),caisse,cast(heureutilise as character(5)) FROM AVOIR ")
user=cur.fetchall()
#for i in user :
#    cur2.execute("INSERT INTO avoir (indice,to_char(date,'YYYYMMDD'),idclient,montant,utilise,saisipar,dateutilise,heure,caisse,heureutilise) VALUES ('%s',%s,0,%s,%s,'%s',%s,%s,'%s',%s) " %\
#        (i[0],i[1],i[2],i[3],i[4],i[5],i[6],i[7],i[8]))
cur.execute("select idcaisse,to_char(date,'YYYYMMDD') , Montant,cast(heure as character(5)),Identifiantcaisse,Nom_Operateur,Prenom_Operateur,Login_Operateur,Ecart from caisse where  operation = 'Ouverture' AND Date > 20201209")
a=cur.fetchall()
for caisse in a :
    cur2.execute("insert into caisse (Date,heure,IDentifiantCaisse,Nom_Operateur,Prenom_Operateur,Login_Operateur,Operation,Montant,Ecart,Cloture) values (%s,'%s','%s','%s','%s','%s','Ouverture',%s,%s,0)" %\
                 (caisse[1],caisse[3],caisse[4],caisse[5],caisse[6],caisse[7],caisse[2],caisse[8]))
    cur2.execute("select max(idcaisse) from caisse ")
    caissemax=cur2.fetchone()[0]
    cur.execute("select P1c,P2c,P5c,P10c,P20c,P50c,P1e,P2e,B5e,B10e,B20e,B50e,B100e FROM MonnaieCaisse WHERE idcaisse = %s " %\
                (caisse[0]))
    monnaie =cur.fetchall()
    cur2.execute("INSERT INTO MonnaieCaisse (idcaisse,P1c,P2c,P5c,P10c,P20c,P50c,P1e,P2e,B5e,B10e,B20e,B50e,B100e,B200e,B500e) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,0,0) " %\
                 (caissemax,monnaie[0],monnaie[1],monnaie[2],monnaie[3],monnaie[4],monnaie[5],monnaie[6],monnaie[7],monnaie[8],monnaie[9],monnaie[10],monnaie[11],monnaie[12]))
    cur.execute("select MIN(idcaisse) from caisse where idcaisse > %s AND operation='Fermeture' AND Login_operateur = '%s' " %\
                (caisse[0],caisse[7]))
    caissemin=cur.fetchone()[0]
    cur.execute("select to_char(date,'YYYYMMDD') ,cast(heure as character(5)) , Montant , Ecart , cloture from caisse where idcaisse = %s " %\
                (caissemin))
    z=cur.fetchall()
    for ferme in z :
        cur2.execute("insert into caisse (Date,heure,IDentifiantCaisse,Nom_Operateur,Prenom_Operateur,Login_Operateur,Operation,Montant,Ecart,Cloture) values (%s,'%s','%s','%s','%s','%s','Fermeture',%s,%s,%s)" %\
                 (ferme[0],ferme[1],caisse[4],caisse[5],caisse[6],caisse[7],ferme[2],ferme[3],ferme[4]))
    cur.execute("select idvente_magasin, to_char(date,'YYYYMMDD'), idclient,code_postal,ville,cast(heure as character(5)) ,montant_total,tauxremise,mode_reglement,poids_total,montant_especes,observations,idavoir,caisse from vente_magasin where idcaisse = %s" %\
                (caisse[0]))
    b=cur.fetchall()
    for venteorigine in b :
        ville=venteorigine[4]
        if ville.find("'") :
            ville=ville.replace("'","\\'")
        if len(venteorigine) == 11  :
            cur2.execute("insert into vente_magasin (date,idclient,code_postal,ville,heure,montant_total,tauxremise,mode_reglement,poids_total,montant_especes,idcaisse,caisse,idavoir) values(%s,0,'%s','%s','%s','%s','%s','%s','%s','%s','%s','%s',%s) " %\
                         (venteorigine[1],venteorigine[3],ville,venteorigine[5],venteorigine[6],venteorigine[7],venteorigine[8],venteorigine[9],venteorigine[10],caissemax,venteorigine[13],venteorigine[12]))
        else :
            cur2.execute("insert into vente_magasin (date,idclient,code_postal,ville,heure,montant_total,tauxremise,mode_reglement,poids_total,montant_especes,idcaisse,caisse) values('%s',0,'%s','%s','%s','%s','%s','%s','%s','%s','%s','%s') " %\
                         (venteorigine[1],venteorigine[3],ville,venteorigine[5],venteorigine[6],venteorigine[7],venteorigine[8],venteorigine[9],venteorigine[10],caissemax,venteorigine[12]))
        cur2.execute("select max(idvente_magasin) from vente_magasin  ")
        venteoriginemax=cur2.fetchone() [0]
        cur.execute ("select idlignes_vente,idproduit,idcategorie,idsous_categorie,montant,nombre,poids,saisie_categorieproduit,tauxremise,montant_remise,tauxtva,montanttva from lignes_vente where idvente_magasin = '%s' " %\
                     (venteorigine[0]))
        c=cur.fetchall()
        for lignevente in c :
            cur2.execute("insert into lignes_vente (idproduit,idcategorie,idsous_categorie,montant,nombre,poids,saisie_categorieproduit,tauxremise,montant_remise,tauxtva,montanttva,idvente_magasin) values ('%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s','%s')" %\
                         (lignevente[1],lignevente[2],lignevente[3],lignevente[4],lignevente[5],lignevente[6],lignevente[7],lignevente[8],lignevente[9],lignevente[10],lignevente[11],venteoriginemax))
        if venteorigine[8] == "Multiple" :
            cur.execute("select modereglement, montant from reglementmultiple where idvente_magasin = '%s' " %\
                        (venteorigine[0]))
            multi = cur.fetchall()
            for paiement in multi :
                cur2.execute("insert into reglementmultiple  (modereglement,montant,idvente_magasin) values ('%s',%s,%s)" %\
                             (paiement[0],paiement[1],venteoriginemax))

conn2.commit()
conn2.close()
conn.close()
