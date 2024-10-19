# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np

from xhtml2pdf import pisa
from jinja2 import Environment, FileSystemLoader
import pprint

# pd.set_option('future.no_silent_downcasting', True)

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
def convert_html_to_pdf(template, df, output_filename, unite_temps, df_zone=None):
    df = df.round(2)
    if df_zone is not None:
        source_html = template.render(df=df, df_zone=df_zone, unite_temps=unite_temps)
    else:
        source_html = template.render(df=df, df_zone=None, unite_temps=unite_temps)
    
    result_file = open(output_filename, "w+b")
    pisa_status = pisa.CreatePDF(
            source_html,                # the HTML to convert
            dest=result_file)           # file handle to recieve result
    result_file.close()

    return pisa_status.err

if __name__ == "__main__":
    df_pointage = lire_xlsx('data/POINTAGE_JUIN_2024.xlsx')
    df_employe = lire_xlsx('data/employe.xlsx')

    # aggregation des lignes par employé, mois, semaine, jour
    columns = {'nb_heure': 'heures',
               'conducteur': 'chauffeur',
               'weeknumber': 'semaine',
               'passager': 'passager'
              }
    aggregated_df_jour = df_pointage.groupby(['id_employe', 'mois', 'weeknumber', 'date']).agg({'nb_heure': 'sum', 'conducteur': 'max', 'passager': 'max'})
    aggregated_df_jour['repas'] = np.where(aggregated_df_jour['nb_heure'] >= 5, 1, 0)
    aggregated_df_jour = aggregated_df_jour.reset_index()
    merged_df = (aggregated_df_jour
                 .merge(df_employe, on='id_employe', how='left')
                 .rename(columns=columns)
                )
    
    # calcul des zones a la journée
    # prendre la zone la plus pres quand il y a plusieurs chantiers sur la même journée
    # faire une liste de zone par employé et par jour
    df_chantier = lire_xlsx('data/chantier.xlsx')
    df_pointage = df_pointage[['date', 'id_employe', 'id_chantier']]

    merged_df_zone = pd.merge(df_pointage, df_chantier, on=['id_chantier' ], how='left')
    merged_df_zone.drop('id_chantier', axis=1, inplace=True)
    dfs_par_id = dict(list(merged_df_zone.groupby(['id_employe', 'date'])))
    df_zone = pd.DataFrame(columns=['date', 'id_employe', 'zone'])
    for id, df in dfs_par_id.items():
        index_min = df['zone'].idxmin()       
        df_row = df.loc[[index_min]]
        df_zone = pd.concat([df_zone, df_row])

    # l'information concernant la zone est rattachée a la journée de l'employé
    merged_df = pd.merge(merged_df, df_zone, on=['date', 'id_employe' ], how='left')
    merged_df = merged_df.fillna(0).infer_objects(copy=False)
    
    # recuperation du template pour la creation du tableau
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('template/pointage.html')
    
    # constitution d'une collection pour parcourir le tableau
    dfs_par_id = dict(list(merged_df.groupby('id_employe')))

    # pour chaque employé, genere un pdf constitué d'un tableau journalier
    for id, df in dfs_par_id.items():
        df.drop('id_employe', axis=1, inplace=True)
        df['date'] = df['date'].dt.strftime('%d-%m-%Y')
        df = df[['prenom', 'nom', 'mois', 'semaine', 'date', 'heures', 'chauffeur', 'passager', 'repas', 'zone']]
        df = df.sort_values(by='date')
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
    
    df_mois_zone = merged_df.groupby(['id_employe', 'mois']).agg({'zone': list}).reset_index()
    dfs_zone_par_id = dict(list(df_mois_zone.groupby('id_employe')))
    for id, df in dfs_zone_par_id.items():
        df = df[['mois', 'zone', 'id_employe']]
        df_zone_dec = pd.DataFrame(columns=['zone 1', 'zone 2', 'zone 3', 'zone 4', 'zone 5', 'zone 6', 'zone 7'])
        list_zone = df.iloc[0, 1]
        df_zone_dec.loc[len(df_zone_dec)] = [list_zone.count(1), list_zone.count(2), list_zone.count(3), list_zone.count(4), list_zone.count(5), list_zone.count(6), list_zone.count(7)]
        df_transposed = df_zone_dec.T
    
    # aggrégé au mois pour obtenir la fiche de salaire, par employé
    merged_df_mois = merged_df_semaine.groupby(['id', 'mois']).agg({'prenom': 'min', 'nom': 'min' , 'heures': 'sum', 'chauffeur': 'sum', 'passager': 'sum', 'repas': 'sum', 'heures supp': 'sum'})
    df = merged_df_mois.reset_index()   
    dfs_par_id = dict(list(df.groupby('id')))
    # pour chaque employé, genere un pdf constitué d'un tableau mensuel
    for id, df in dfs_par_id.items():
        df = df[['prenom', 'nom', 'mois', 'heures', 'chauffeur', 'passager', 'repas', 'heures supp']]
        convert_html_to_pdf(template, df, "result/pointage_mois_" + str(df['prenom'].unique()[0]) + "_" + str(df['nom'].unique()[0]) + ".pdf", "mois", df_transposed) 
    
    # controle des incohérences
    # controle 1 
    # si la somme des heures > 5 alors il n'y a pas de repas
    
    # controle 2 
    # si la zone < 0 ou zone > 7 alors elle est erronée
    
    # controle 3 
    # 