# -*- coding: utf-8 -*-
import model.datamodel as model
import model.config as conf
import model.database as db
import logging

if __name__ == "__main__":
    # chargement des parametres du traitement
    config = conf.Config()
    config.load('config.ini')

    # initialisation de la base de données
    bdd = db.Postgresql(config)
    
    # mise a jour de la table employes dans la bdd
    employes = model.Employes()
    employes.lire_xlsx(config.employes_path)
    bdd.upsert(employes)

    # mise a jour de la table chantiers dans la bdd
    chantiers = model.Chantiers()
    chantiers.lire_xlsx(config.chantiers_path)
    bdd.upsert(chantiers)

