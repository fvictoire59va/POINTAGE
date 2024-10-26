import pandas as pd

class Entity:
    def __init__(self):
        pass
    
    def dynamic_groupby(df, group_col, value_col, agg_func):
        
        return df.groupby(group_col)[value_col].agg(agg_func)
    
    def lire_xlsx(self, nom_du_fichier):
        try:
            df = pd.read_excel(nom_du_fichier)
            df = df.fillna(0)
            df = df.query("id_employe != 0")
            # print(df)
            # aggregated_df = self.dynamic_groupby(df, self.key, self.columnWithoutKey, 'sum')
            # print(df)
            iterables = [df[elt] for elt in self.columns]
            self.data_to_upsert = list(zip(*iterables))
            # print(self.data_to_upsert)
            return df
        except:
            print("le fichier excel ne peut pas etre lu")

class Employes(Entity):
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
        self.key = ['id_employe']
        self.data_to_upsert = None
    
    def lire_xlsx(self, nom_du_fichier):
        super().lire_xlsx(nom_du_fichier)
            
class Chantiers(Entity):
    def __init__(self):
        self.tableName = "chantiers"
        self.columns = [
                        'id_chantier', 
                        'zone'
                        ]
        self.columnWithoutKey = [
                        'zone'
                        ]
        self.key = ['id_chantier']
        self.data_to_upsert = None
    
    def lire_xlsx(self, nom_du_fichier):
        super().lire_xlsx(nom_du_fichier)
            
class Pointages(Entity):
    def __init__(self):
        self.tableName = "pointages"
        self.columns = [
                        'id_employe',
                        'date',
                        'id_chantier',
                        'nb_heure',
                        'chauffeur',
                        'trajet',
                        'transport'
                        ]
        self.columnWithoutKey = [
                        'nb_heure',
                        'chauffeur',
                        'trajet',
                        'transport'
                        ]
        self.key = [
                        'id_employe',
                        'date',
                        'id_chantier'
                    ]
        self.data_to_upsert = None
    
    def lire_xlsx(self, nom_du_fichier):
        super().lire_xlsx(nom_du_fichier)