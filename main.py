# Codé par THEVAKUMARAN Krishnarash, LASNAMI Sara
# Optimisation Combinatoire
# Bin Packing (Methode First Fit et First Fit Decreasing)

import xlsxwriter
import numpy as np
import random

""" On implémente la classe Sac """
class Sac(object):
    def __init__(self):
        # La liste qui va contenir les objets qui vont rentrer dans le sac
        self.objets = []
        self.poids_total = 0

    # La fonction ajout_objets permet d'ajouter un objet dans le sac
    def ajout_objets(self, objet):
        self.objets.append(objet)
        self.poids_total += objet

""" Algorithme de la méthode First Fit """
def first_fit(liste_objets, capacite):
    sacs = []
    # Liste contenant les objets dans un ordre aléatoire
    liste_aleatoire = np.random.permutation(liste_objets)
    liste_objets = liste_aleatoire.tolist()

    for objet_liste_random in liste_objets:
        # pour chaque objets, nous cherchons s'il y a un sac libre où l'objet peut être placé.
        for objets in sacs:
            # Si le poids_total de notre sac et le poids de l'objet ne depasse pas la capacité de notre sac
            if objets.poids_total + objet_liste_random <= capacite:
                # nous ajoutons l'objet dans le sac avec la fonction ajout_objets
                objets.ajout_objets(objet_liste_random)
                break
        # Sinon : c'est à dire s'il n'y a pas de sac libre où l'objet peut être placé.
        else:
            # Nous ouvrons donc un nouveau sac et y ajoutons l'objet.
            objets = Sac()
            objets.ajout_objets(objet_liste_random)
            sacs.append(objets)

    return sacs

""" Algorithme de la méthode First Fit Decreasing """
def first_fit_decreasing(liste_objets, capacite_sac):
    # Trie la liste des objets dans un ordre décroissant avec la methode sorted et reverse = True.
    liste_decroissant = sorted(liste_objets, reverse=True)
    sacs =[]
    for objet_liste_dec in liste_decroissant:
        # pour chaque objets, nous cherchons s'il y a un sac libre où l'objet peut être placé.
        for objets in sacs:
            # Si le poids_total de notre sac et le poids de l'objet ne depasse pas la capacité de notre sac
            if objets.poids_total + objet_liste_dec <= capacite_sac:
                # nous ajoutons l'objet dans le sac avec la fonction ajout_objets
                objets.ajout_objets(objet_liste_dec)
                break
        # Sinon : c'est à dire s'il n'y a pas de sac libre où l'objet peut être placé.
        else:
            # Nous ouvrons donc un nouveau sac et y ajoutons l'objet.
            objets = Sac()
            objets.ajout_objets(objet_liste_dec)
            sacs.append(objets)
    return sacs

""" Cette fonction prend la liste des objets et la capacité du sac et renvoi la solution optimal des majorants """
def solution_optimal(liste_objets, capacite_sac):
    poids_total = 0
    for objet_liste in liste_objets:
        # Nous ouvrons donc un nouveau sac et y ajoutons l'objet et chaque ajout d'objet, le poids de l'objet s'ajoutera au poids total
        objets = Sac()
        objets.ajout_objets(objet_liste)
        poids_total += objets.poids_total
    # Solution Optimal est le poids total divise par la capacite du sac
    solution = poids_total/(capacite_sac)
    return solution

""" Cette fonction genere un fichier excel dans notre cas, pour inserer la capacite, le nombre d'objet et la liste des objets aleatoire
 pour ensuite savoir les nombres de sac utilise avec la methode ffd"""
def generation_fichier(fichier):
    # Creation d'un fichier excel et d'une feuille excel
    workbook = xlsxwriter.Workbook(fichier)
    worksheet = workbook.add_worksheet("Resultat")

    # Insertion des headers
    worksheet.write("A1", "Les tables")
    worksheet.write("B1", "Capacite du sac")
    worksheet.write("C1", "Nombre d'objet")
    worksheet.write("D1", "Solution Optimal")
    worksheet.write("E1", "Solution Algorithme First Fit")
    worksheet.write("F1", "Efficacite Algorithme First Fit")
    worksheet.write("G1", "Solution Algorithme First Fit Decreasing")
    worksheet.write("H1", "Efficacite Algorithme First Fit Decreasing")
    worksheet.write("I1", "FF < FFD")

    # Pour la generation de 10 tests
    for i in range(2,101):
        j=100
        # Instantiation des valeurs (ligne, capacite, nombre d'objets, liste d'objets, solution optimal)
        capacite = random.randint(1,200)
        j+=50
        nbre_objet = random.randint(10,j)
        liste_objet = [random.randint(0,capacite) for p in range(0,nbre_objet)]
        bin_package_first_fit_decreasing = first_fit_decreasing(liste_objet, capacite)
        bin_package_first_fit_random = first_fit(liste_objet, capacite)
        solution = solution_optimal(liste_objet, capacite)

        # Ecriture des resultats numeriques sur l'excel
        worksheet.write("A"+str(i), "Tables "+str(i))
        worksheet.write("B"+str(i), capacite)
        worksheet.write("C"+str(i), nbre_objet)
        worksheet.write("D"+str(i), round(solution))
        worksheet.write("E"+str(i), len(bin_package_first_fit_random))
        worksheet.write("F"+str(i), (len(bin_package_first_fit_random) / round(solution)))
        worksheet.write("G"+str(i), len(bin_package_first_fit_decreasing))
        worksheet.write("H"+str(i), (len(bin_package_first_fit_decreasing) / round(solution)))
        if((((len(bin_package_first_fit_random) / round(solution)) <= len(bin_package_first_fit_decreasing) / round(solution)))):
            worksheet.write("I" + str(i), "True")
        else :
            worksheet.write("I" + str(i), "False")

    # Fermeture de l'excel
    workbook.close()

#Appel de fonction de generation de fichier excel (Resultats numeriques sur ces excels)
generation_fichier("ResultatsNumeriques.xlsx")




