# -*- coding: utf-8 -*-
import pandas as pd
import numpy as np
import warnings

from xhtml2pdf import pisa
from jinja2 import Environment, FileSystemLoader
import pprint

# pd.set_option('future.no_silent_downcasting', True)
warnings.simplefilter(action='ignore', category=FutureWarning)
pd.options.mode.chained_assignment = None

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
    annee = '2024'
    mois = 'juin'
    df_pointage = lire_xlsx('data/input/POINTAGE_JUIN_2024.xlsx')
    df_employe = lire_xlsx('data/referentiel/employe.xlsx')
    df_chantier = lire_xlsx('data/referentiel/chantier.xlsx')
    
    # recuperation du template pour la creation du tableau
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('template/pointage.html')
    
    # aggregation des lignes par employé, mois, semaine, jour
    columns = {'nb_heure': 'heures',
               'chauffeur': 'chauffeur',
               'weeknumber': 'semaine',
               'trajet': 'trajet'
              }
    
    aggregats = {'nb_heure': 'sum', 
                 'chauffeur': 'max', 
                 'trajet': 'max',
                 'transport': 'max',
                 'zone': 'max'
           }
    
    pivot = ['id_employe', 
             'mois', 
             'weeknumber', 
             'date'
             ]
    
    aggregated_df_jour = (df_pointage
                          .merge(df_chantier, on=['id_chantier' ], how='left')
                          .query("id_chantier != 'abs'") 
                          .groupby(pivot)
                          .agg(aggregats)
                          .reset_index()
                          .merge(df_employe, on='id_employe', how='left')
                          .rename(columns=columns)
                          .fillna(0).infer_objects(copy=False)
                        )   

    aggregated_df_jour['repas'] = np.where(aggregated_df_jour['heures'] < 5 , 0, 1)
    
    id_employe_list = df_employe['id_employe'].tolist()
    for df_employe in id_employe_list:
        df = aggregated_df_jour[aggregated_df_jour['id_employe']==df_employe]
        df['date'] = df['date'].dt.strftime('%d-%m-%Y')
        df = df[['prenom', 'nom', 'mois', 'semaine', 'date', 'heures', 'chauffeur', 'trajet', 'transport', 'repas', 'zone']]
        df = df.sort_values(by='date')
        path_filename = "output/"+ annee + "/" + mois + "/jour/" + str(df['prenom'].unique()[0]) + "_" + str(df['nom'].unique()[0]) + ".pdf"
        convert_html_to_pdf(template, df, path_filename, "jour")

    # # aggrégé a la semaine pour calculer les heures supplémentaires, par employé
    aggregats = {'prenom': 'min', 
                 'nom': 'min' , 
                 'heures': 'sum', 
                 'chauffeur': 'sum', 
                 'trajet': 'sum', 
                 'transport': 'sum',
                 'repas': 'sum'
                }
    pivot = ['id_employe', 
             'mois', 
             'semaine'
            ]
    merged_df_semaine = (aggregated_df_jour
                         .groupby(pivot)
                         .agg(aggregats)
                         .reset_index()
    )
    merged_df_semaine['heures supp'] = np.where(merged_df_semaine['heures'] > 35 , merged_df_semaine['heures'] - 35, 0)
    columns = ['prenom', 
               'nom', 
               'mois', 
               'semaine', 
               'heures', 
               'chauffeur', 
               'trajet', 
               'transport',
               'repas', 
               'heures supp'
            ]
    # pour chaque employé, genere un pdf constitué d'un tableau hebdomadaire
    for df_employe in id_employe_list:
        df = merged_df_semaine[merged_df_semaine['id_employe']==df_employe]
        df = df[columns]
        df = df.sort_values(by='semaine')
        path_filename = "output/"+ annee + "/" + mois + "/semaine/" + str(df['prenom'].unique()[0]) + "_" + str(df['nom'].unique()[0]) + ".pdf"
        convert_html_to_pdf(template, df, path_filename, "semaine")
    
    # aggrégé au mois pour obtenir la fiche de salaire, par employé

    aggregat = {'prenom': 'min', 
                'nom': 'min' , 
                'heures': 'sum', 
                'chauffeur': 'sum', 
                'trajet': 'sum', 
                'transport': 'sum',
                'repas': 'sum', 
                'heures supp': 'sum'
            }
    
    pivot = ['id_employe', 
             'mois'
            ]
    merged_df_mois = (merged_df_semaine
                      .groupby(pivot)
                      .agg(aggregat)
                      .reset_index()
    )
    
    columns = ['prenom', 'nom', 'mois', 'heures', 'chauffeur', 'trajet', 'transport', 'repas', 'heures supp']
    for df_employe in id_employe_list:
        
        # calcul du nombre de zone
        df = aggregated_df_jour[aggregated_df_jour['id_employe']==df_employe]
        df = df[['zone', 'chauffeur', 'trajet', 'transport']]
        aggregat = {'chauffeur': 'sum', 
                    'trajet': 'sum',
                    'transport': 'sum'
                }
        df_zone = (df
                .groupby('zone')
                .agg(aggregat)
                .reset_index()
                .astype(int)
        )
        
        df = merged_df_mois[merged_df_mois['id_employe']==df_employe]
        df = df[columns]
        path_filename = "output/"+ annee + "/" + mois + "/mois/" + str(df['prenom'].unique()[0]) + "_" + str(df['nom'].unique()[0]) + ".pdf"
        convert_html_to_pdf(template, df, path_filename, "mois", df_zone) 