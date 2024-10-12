# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np

from xhtml2pdf import pisa
from jinja2 import Environment, FileSystemLoader
# from io import BytesIO

month_names_fr = {
    1: 'janvier', 
    2: 'février', 
    3: 'mars', 
    4: 'avril',
    5: 'mai', 
    6: 'juin', 
    7: 'juillet', 
    8: 'août',
    9: 'septembre', 
    10: 'octobre', 
    11: 'novembre', 
    12: 'décembre'
}

def lire_xlsx(nom_du_fichier):
    df = pd.read_excel(nom_du_fichier)
    if 'date' in df.columns:
        df['mois'] = df['date'].dt.month.map(month_names_fr)
        df['weeknumber'] = df['date'].dt.isocalendar().week
    else:
        pass
    return df

# Utility function
def convert_html_to_pdf(template, df, output_filename, unite_temps):
    source_html = template.render(df=df, unite_temps=unite_temps)
    result_file = open(output_filename, "w+b")
    pisa_status = pisa.CreatePDF(
            source_html,                # the HTML to convert
            dest=result_file)           # file handle to recieve result
    result_file.close()

    return pisa_status.err

if __name__ == "__main__":
    df_pointage = lire_xlsx('data/pointage.xlsx')
    df_employe = lire_xlsx('data/employe.xlsx')

    # aggregation des lignes par employé, mois, semaine, jour
    aggregated_df_jour = df_pointage.groupby(['id_employe', 'mois', 'weeknumber', 'date']).agg({'nb_heure': 'sum', 'conducteur': 'max', 'passager': 'max'})
    aggregated_df_jour['repas'] = np.where(aggregated_df_jour['nb_heure'] >= 5, 1, 0)
    aggregated_df_jour = aggregated_df_jour.reset_index()
    merged_df = pd.merge(aggregated_df_jour, df_employe, on=['id_employe' ], how='left')
    merged_df = merged_df.rename(columns={'nb_heure': 'heures',
                       'conducteur': 'chauffeur',
                       'weeknumber': 'semaine',
                       'passager': 'passager'
                     })
    
    # recuperation du template pour la creation du tableau
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('template/pointage.html')
    
    # constitution d'une collection pour parcourir le tableau
    dfs_par_id = dict(list(merged_df.groupby('id_employe')))

    # pour chaque employé, genere un pdf constitué d'un tableau journalier
    for id, df in dfs_par_id.items():
        df.drop('id_employe', axis=1, inplace=True)
        df['date'] = df['date'].dt.strftime('%d-%m-%Y')
        df = df[['prenom', 'nom', 'mois', 'semaine', 'date', 'heures', 'chauffeur', 'passager', 'repas']]
        df = df.sort_values(by='date')
        df = df.fillna(0)
        convert_html_to_pdf(template, df, "result/" + str(df['prenom'].unique()[0]) + "_" + str(df['nom'].unique()[0]) + "_" + str(df['mois'].unique()[0]) + "_decompte_heures_jours.pdf", "jour")

    # aggrégé a la semaine pour calculer les heures supplémentaires, par employé
    merged_df['id'] = merged_df['prenom'] + "_" + merged_df['nom']
    merged_df_semaine = merged_df.groupby(['id', 'mois', 'semaine']).agg({'prenom': 'min', 'nom': 'min' , 'heures': 'sum', 'chauffeur': 'sum', 'passager': 'sum', 'repas': 'sum'})
    merged_df_semaine['heures supp'] = np.where(merged_df_semaine['heures'] > 35 , merged_df_semaine['heures'] - 35, 0)
    df = merged_df_semaine.reset_index()
    dfs_par_id = dict(list(df.groupby('id')))
    
    # pour chaque employé, genere un pdf constitué d'un tableau hebdomadaire
    for id, df in dfs_par_id.items():
        df = df[['prenom', 'nom', 'mois', 'semaine', 'heures', 'chauffeur', 'passager', 'repas', 'heures supp']]
        convert_html_to_pdf(template, df, "result/pointage_semaine_" + str(df['prenom'].unique()[0]) + "_" + str(df['nom'].unique()[0]) + ".pdf", "semaine")
    
    # aggrégé au mois pour obtenir la fiche de salaire, par employé
    merged_df_mois = merged_df_semaine.groupby(['id', 'mois']).agg({'prenom': 'min', 'nom': 'min' , 'heures': 'sum', 'chauffeur': 'sum', 'passager': 'sum', 'repas': 'sum', 'heures supp': 'sum'})
    df = merged_df_semaine.reset_index()   
    dfs_par_id = dict(list(df.groupby('id')))
    
    # pour chaque employé, genere un pdf constitué d'un tableau mensuel
    for id, df in dfs_par_id.items():
        df = df[['prenom', 'nom', 'mois', 'heures', 'chauffeur', 'passager', 'repas', 'heures supp']]
        convert_html_to_pdf(template, df, "result/pointage_mois_" + str(df['prenom'].unique()[0]) + "_" + str(df['nom'].unique()[0]) + ".pdf", "mois") 
