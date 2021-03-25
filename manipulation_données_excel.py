"""
Created on Mon Nov 23 17:22:21 2020

@author: Sawadogo Abdoul Raouf Wendyam Issoufou
"""

import pandas as pd


fichier_extraction_excel = 'Nom_fichier_extraction.xlsx'
fichier_remplissage_excel = 'Nom_fichier_remplissage'
fichier_intermediaire = 'fichier_temporaire.xlsx' #permet le fonctionnement du programme aussi bien dans le cas ou le fichier de remplissage est protégé ou pas
données_extraction = pd.read_excel(fichier_extraction_excel, sheet_name = 'Nom de la page ou les données sont situées') #lecture des données sur les différents fichiers
données_remplissage = pd.read_excel(fichier_remplissage_excel, sheet_name = 'Nom de la page ou les données sont situées')
données_temporaire = pd.read_excel(fichier_intermediaire, sheet_name = 'fichier_temporaire')

df = pd.DataFrame({'produits':[],
                   'ventes':[],
                   'stocks':[],
                   'CDE':[]})#création d'une feuille de donnée avec le nom des colonnes voulues


def comporte_caracteres_communs(produit_1,produit_2,caracteres_a_verifier):
    """Fonction permettant de comparer les premiers caractères de deux noms
       attributes:
           produit_1: premier produit
           produit_2: deuxième produit
           caractères a verifier: nombre de caractères a verifier
        returns:
           booléen: True si les caractères sont les meme, false sinon
    """
    
    for index in range (0,caracteres_a_verifier):
          
         
            
        if not produit_1[index].lower() == produit_2[index].lower(): #on vérifie que premiers caractères des deux produits correspondent
                
            return False
    return True


def identifier_donnees(feuille,ligne,colonne_ventes,colonne_stocks,colonne_ventes,colonne_commandes):
    """Fonction permettant d'identifier les données a ajouter sur la feuillle de donnée ici les ventes, les stocks et les commandes
       attributs:
           feuille: feuille de donnée
           ligne:ligne au niveau de la feuille
           colonne_ventes:colonne des ventes
           colonne_stocks:colonne du stock
           colonne_commandes: colonne des commandes
        returns:
           liste:liste des données identifiées
    """

    ventes = float(feuille.iloc[ligne, colonne_ventes])   
    stocks = float(feuille.iloc[ligne, colonne_stocks])
    commandes = float(feuille.iloc[ligne, colonne_commandes])

    valeurs_donnees =[ventes, stocks, commandes] #liste des valeurs identifiées

    return valeurs_donners

    def sauvegarder_donnees(feuille,ligne_2,fichier,valeurs):
        """Fonction permettant de sauvegarder les données identifiées
           attributes:
               feuille: feuille de données pour l'écriture des valeurs
               ligne_2: ligne d'écriture des valeurs
               fichier: fichier de sauvegarde
               valeurs: liste des valeurs
               
        """
        df.loc[ligne_2,'ventes'] = valeurs[1]
        df.loc[ligne_2,'stocks'] = valeurs[2]
        df.loc[ligne_2,'CDE'] = valeurs[3]
        writer = pd.ExcelWriter(fichier)
        df.to_excel(writer,'fichier temporaire',index=False)
        writer.save()

        print ('ventes enregistrées =', valeurs[1])
        print ('stocks enregistrées =', valeurs[2])
        print ('cde enregistrées =', valeurs[3])






print(données_extraction)
print(données_remplissage)


for ligne_remplissage in range ('première_ligne','dernière_ligne'):

    produit_2 = données_remplissage.iloc[ligne_remplissage,colonne_nom_produit]#locallisation du nom du produit a l'aide des données affichées
    
    df.loc[ligne_remplissage,'produits'] = produit_2 #ajout du produit dans la feuille de données à la colonne produit
    writer = pd.ExcelWriter('fichier_temporaire.xlsx')#écriture du contenu de la feuille dans le fichier intermediaire et sauvegarde
    df.to_excel(writer,'fichier_temporaire',index=False)
    writer.save()
    
    ligne_extraction = #entrez le numéro de la première ligne du fichier d'extraction
        
    
    while ligne_extraction <= #entrez le numéro de la dernière ligne du fichier d'extraction :
        
        prod_1 = données_extraction.iloc[ligne_1,1] #identification du produit dont les données seront extraites
                     
        if comporte_caracteres_communs(produit_1,produit_2,'nombres de caractères a verifier'):
                   
                
            #affichage des produits correspondants     
            print(prod_2)
            print(prod_1)

            #validation de l'utilisateur        
            rep = input("Le couple correspond ? (o=oui,n=non) :")
                    
                    
            if rep == 'o':              
               
               liste_valeurs = identifier_données(données_extraction,ligne_extraction,'colonne ventes','colonne stocks','colonne commandes',)
                             
               ligne_extraction += 1

               sauvegarder_donnees(df,ligne_remplissage,fichier_intermediaire,liste_valeurs)
                
                    
                
            else :
                    
                ligne_extraction += 1
        
        
        else:
         
                     
            ligne_extraction += 1


#
# a la fin du programme, on ouvre le fichier intermédiaire et on copie son contenu dans le fichier de remplissage
#
#