# Codé par THEVAKUMARAN Krishnarash, LASNAMI Sara
# Optimisation Combinatoire
# Bin Packing (Methode First Fit Decreasing)

import openpyxl

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

""" Cette fonction prend un fichier (excel de preference) qu'on renseigne en parametre
et renvoie la liste des objets et la capacité du sac. """
def extraction_fichier(fichier):
    fichier1 = openpyxl.load_workbook(fichier)
    feuille = fichier1.active
    liste_objets = []
    # Recupere le nombre d'objet qui est dans la 2eme ligne de la 1ere colonne
    nbre_objets = feuille.cell(row = 2, column = 2).value
    # Recupere la capacite du sac qui est dans la 2eme ligne de la 2eme colonne
    capacite = feuille.cell(row = 3, column = 2).value
    for i in range(2,nbre_objets+2):
        # Recupere le poids des objets de la 3eme colonne
        case_objet = feuille.cell(row = 1, column = i)
        liste_objets.append(case_objet.value)
    return liste_objets,capacite

def solution_optimal(liste_objets, capacite_sac):
    poids_total = 0.0
    for objet_liste in liste_objets:
        # Nous ouvrons donc un nouveau sac et y ajoutons l'objet et chaque ajout d'objet, le poids de l'objet s'ajoutera au poids total
        objets = Sac()
        objets.ajout_objets(objet_liste)
        poids_total += objets.poids_total
    # Solution Optimal est le poids total divise par la capacite du sac
    solution = poids_total/float(capacite_sac)
    return solution

print("Utilisation de l'algorithme FFD avec des valeurs rentre dans le code")
bin_package_first_fit_decreasing_1 = first_fit_decreasing([3,4,4,3,3,2,1], 10)
solution = solution_optimal([3,4,4,3,3,2,1], 10)
print("Le nombre de sac utilise avec la méthode First Fit Decreasing : ", len(bin_package_first_fit_decreasing_1), "sacs")
print("La solution optimale : ", solution, "sacs")

print("\nUtilisation de l'algorithme FFD avec un excel")
resultat = extraction_fichier("ValeurTest_FFD1.xlsx")
liste_objet, capacite = resultat[0], resultat[1]
bin_package_first_fit_decreasing = first_fit_decreasing(liste_objet, capacite)
solution = solution_optimal(liste_objet, capacite)

print("Le nombre de sac utilise avec la méthode First Fit Decreasing : ", len(bin_package_first_fit_decreasing), "sacs")
print("La solution optimale : ", solution, "sacs")

print("\nL'efficacite de l'algorithme approche de la methode bin packing ffd :",(len(bin_package_first_fit_decreasing) - solution))


