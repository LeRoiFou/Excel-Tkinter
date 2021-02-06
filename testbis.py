"""
Dans ce programme, on apprend à récupérer les données d'un fichier Excel comprenant 2 onglets dont chacun d'eux contient
un échantillonnage des données de la liasse fiscale

L'objectif est de récupérer dans chacun des onglets les données de la cellule A1 à la cellule D4 d'un fichier Calc de
Libreoffice converti en un fichier .xlsx d'Excel (fichier -> enregistrer sous -> Type : Excel 2007-365) et d'afficher
les données dans deux onglets distints sous la forme de champs de saisies

Éditeur : Laurent REYNAUD
Date : 06-02-2021
"""

from tkinter import *
from tkinter import ttk  # pour les onglets
from tkinter import filedialog  # pour la fenêtre d'ouverture du fichier à récupérer
from openpyxl import load_workbook  # pour récupérer les données Excel

root = Tk()
root.title('Récupération des données Excel')

"""Configuration des onglets"""
my_notebook = ttk.Notebook(root)
my_notebook.pack()

"""Fonction permettant d'ouvrir un fichier dans l'ordinateur"""
my_program = filedialog.askopenfilename()  # ouverture de la boîte de dialogue

"""Chargement de la feuille active d'Excel"""
wb = load_workbook(my_program)


class OpenProgram(Frame):
    """Classe permettant de récupérer les données du fichier Excel et de les afficher dans Python avec le module Tkinter
    """

    def __init__(self, master, number):  # number en paramètre qui constitue le n° de l'onglet d'Excel
        super().__init__(master)
        self.pack()
        self.widgets(number)  # reprise du paramètre 'number' ci-dessus

    def widgets(self, number):  # reprise du paramètre 'number' ci-dessus

        """Assignation de la feuille d'Excel (feuille = worksheet = ws)"""
        ws = wb[wb.sheetnames[number]]  # number = numéro de l'onglet du Fichier Excel

        """Saisie de la plage de cellules des données à récupérer du fichier Excel"""
        range_excel = ws['A1':'D4']

        """Assignation d'une liste"""
        my_list = []

        """Affichage des données"""
        for cell in range_excel:
            for x in cell:
                my_list.append(x.value)  # ajout des données de la feuille Excel dans une liste

        """Assignation de variables"""
        height = 5
        width = 5
        nb = 0

        """Ajout des champs de saisis"""
        for i in range(1, height):  # rows
            for j in range(1, width):  # columns
                cell = Entry(self, justify='center')
                cell.grid(row=i, column=j, ipady=5)
                cell.insert(0, my_list[nb])  # insertion des données de la liste 'my_list' dans les champs de saisies
                nb += 1  # incrémentation +1 des données de la liste 'my_list'


"""Configuration des cadres"""
frame_2050 = Frame(my_notebook)
frame_2051 = Frame(my_notebook)

"""Ajout des onglets"""
my_notebook.add(frame_2050, text='Actif du bilan')
my_notebook.add(frame_2051, text='Passif du bilan')

open_2050 = OpenProgram(frame_2050, 0)  # en argument : l'onglet de tkinter et le n° d'onglet d'Excel
open_2051 = OpenProgram(frame_2051, 1)  # en argument : l'onglet de tkinter et le n° d'onglet d'Excel
root.mainloop()
