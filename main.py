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

# Fonction pour générer un PDF pour un employé
def generate_pdf_for_employee(employe, template):
    html = template.render(employe=employe)
    result = BytesIO()
    pdf = pisa.pisaDocument(BytesIO(html.encode("UTF-8")), result)
    if not pdf.err:
        result.seek(0)
        return result.getvalue()
    else:
        print(f"Erreur lors de la création du PDF pour {employe['Nom']}: {pdf.err}")
        return None

if __name__ == "__main__":
    df_pointage = lire_xlsx('data/pointage.xlsx')
    # aggrégé a la journée et par employé
    aggregated_df_jour = df_pointage.groupby(['id_employe', 'weeknumber','date']).agg({'nb_heure': 'sum', 'conducteur': 'sum', 'passager': 'sum'})
    aggregated_df_jour['repas'] = np.where(aggregated_df_jour['nb_heure'] >= 5, 1, 0)

    df_employe = lire_xlsx('data/employe.xlsx')

    # jointure sur la table de depart sur la semaine et l'employe
    merged_df = pd.merge(df_pointage, df_employe, on=['id_employe'], how='left')

    merged_df = merged_df.rename(columns={'nb_heure': 'heures',
                    'id_employe': 'employe',
                    'id_chantier': 'chantier',
                    'conducteur': 'chauffeur',
                    'weeknumber': 'semaine',
                    'passager': 'passager'
                    })

    # print(merged_df)

    # Configuration de Jinja2 pour charger le template
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('template/pointage_jour.html')

    dfs_par_id = dict(list(merged_df.groupby('employe')))

    for id, df in dfs_par_id.items():
        pdf_content = generate_pdf_for_employee(df, template)
