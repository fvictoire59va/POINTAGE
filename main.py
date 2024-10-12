# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np

from xhtml2pdf import pisa
from jinja2 import Environment, FileSystemLoader
from io import BytesIO

month_names_fr = {
    1: 'janvier', 2: 'février', 3: 'mars', 4: 'avril',
    5: 'mai', 6: 'juin', 7: 'juillet', 8: 'août',
    9: 'septembre', 10: 'octobre', 11: 'novembre', 12: 'décembre'
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
def convert_html_to_pdf(template, df, output_filename, titre):
    source_html = template.render(df=df, titre=titre)
    result_file = open(output_filename, "w+b")
    pisa_status = pisa.CreatePDF(
            source_html,                # the HTML to convert
            dest=result_file)           # file handle to recieve result
    result_file.close()

    return pisa_status.err

if __name__ == "__main__":
    df_pointage = lire_xlsx('data/pointage.xlsx')
    df_employe = lire_xlsx('data/employe.xlsx')

    aggregated_df_jour = df_pointage.groupby(['id_employe', 'mois', 'weeknumber', 'date']).agg({'nb_heure': 'sum', 'conducteur': 'max', 'passager': 'max'})
    aggregated_df_jour['repas'] = np.where(aggregated_df_jour['nb_heure'] >= 5, 1, 0)
    aggregated_df_jour = aggregated_df_jour.reset_index()
    merged_df = pd.merge(aggregated_df_jour, df_employe, on=['id_employe' ], how='left')
    merged_df = merged_df.rename(columns={'nb_heure': 'heures',
                       'conducteur': 'chauffeur',
                       'weeknumber': 'semaine',
                       'passager': 'passager'
                     })
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('template/pointage_jour.html')
    dfs_par_id = dict(list(merged_df.groupby('id_employe')))
    pisa.showLogging()
    df_check = pd.DataFrame()
    df_check_global = pd.DataFrame()
    for id, df in dfs_par_id.items():
        df.drop('id_employe', axis=1, inplace=True)
        df['date'] = df['date'].dt.strftime('%d-%m-%Y')
        df = df[['prenom', 'nom', 'mois', 'semaine', 'date', 'heures', 'chauffeur', 'passager', 'repas']]
        df = df.sort_values(by='date')
        df = df.fillna(0)
        df_check = df[['date', 'prenom','nom']].copy()
        # df_check['check chauffeur passager'] = np.where(df['chauffeur'] + df['passager'] > 1 , 'ERREUR chauffeur ou passager', 'OK')
        # df_check['Check zone>7'] = np.where(df['zone'] > 7 , 'la zone ne peut pas etre > 7', 'OK')
        # df_check['Check zone<1'] = np.where(df['zone'] < 1 , 'la zone ne peut pas etre < 1', 'OK')
        # df_check=df_check[(df_check['check chauffeur passager'] != 'OK') | 
        #             (df_check['Check zone>7'] != 'OK') |  
        #             (df_check['Check zone<1'] != 'OK')
        #            ]
        # df_check_global=pd.concat([df_check_global, df_check], ignore_index=True)
        convert_html_to_pdf(template, df, str(df['prenom'].unique()[0]) + "_" + str(df['nom'].unique()[0]) + "_" + str(df['mois'].unique()[0]) + "_decompte_heures_jours.pdf", "jour")
    # convert_html_to_pdf(template, df_check_global, "check.pdf")

    # aggrégé a la semaine pour calculer les heures supplémentaires, par employé
    merged_df['id'] = merged_df['prenom'] + "_" + merged_df['nom']
    merged_df_semaine = merged_df.groupby(['id', 'mois', 'semaine']).agg({'prenom': 'min', 'nom': 'min' , 'heures': 'sum', 'chauffeur': 'sum', 'passager': 'sum', 'repas': 'sum'})
    merged_df_semaine['heures supp'] = np.where(merged_df_semaine['heures'] > 35 , merged_df_semaine['heures'] - 35, 0)
    dfs_par_id = dict(list(merged_df_semaine.groupby('id')))
    for id, df in dfs_par_id.items():
        print(df)
        df = df[['prenom', 'nom', 'mois', 'semaine', 'heures', 'chauffeur', 'passager', 'repas']]
        convert_html_to_pdf(template, df, "pointage_semaine_" + str(df['prenom'].unique()[0]) + "_" + str(df['nom'].unique()[0]) + ".pdf", "semaine")

    
    # aggrégé au mois pour obtenir la fiche de salaire, par employé
    # merged_df_mois = merged_df_semaine.groupby(['prenom', 'nom', 'mois']).agg({'heures': 'sum', 'chauffeur': 'sum', 'passager': 'sum', 'repas': 'sum', 'heures supp': 'sum'})
    # convert_html_to_pdf(template, merged_df_mois, "pointage_mois.pdf")   
    # print(merged_df_mois)
