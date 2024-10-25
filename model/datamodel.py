import pandas as pd

class Employes:
    def __init__(self):
        self.tableName = "employes"
        self.columns = [
                        'id_employe', 
                        'prenom', 
                        'nom'
                        ]
        self.columnWithoutKey = [
                        'prenom', 
                        'nom'
                        ]
        self.key = 'id_employe'
        self.data_to_upsert = None
    
    def lire_xlsx(self, nom_du_fichier):
        try:
            df = pd.read_excel(nom_du_fichier)
            iterables = [df[elt] for elt in self.columns]
            self.data_to_upsert = list(zip(*iterables))
            return df
        except:
            print("le fichier excel ne peut pas etre lu")
            
class Chantiers:
    def __init__(self):
        self.tableName = "chantiers"
        self.columns = [
                        'id_chantier', 
                        'zone'
                        ]
        self.columnWithoutKey = [
                        'zone'
                        ]
        self.key = 'id_chantier'
        self.data_to_upsert = None
    
    def lire_xlsx(self, nom_du_fichier):
        try:
            df = pd.read_excel(nom_du_fichier)
            iterables = [df[elt] for elt in self.columns]
            self.data_to_upsert = list(zip(*iterables))
            return df
        except:
            print("le fichier excel ne peut pas etre lu")