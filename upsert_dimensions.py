# -*- coding: utf-8 -*-
import model.datamodel as model
import model.config as conf
import model.database as db
import logging
import os
import sys

if __name__ == "__main__":
    # Configuration basique du logger
    logging.basicConfig(level=logging.DEBUG, 
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                        filename=os.path.basename(sys.argv[0]) + '.log',
                        filemode='w')
    
    # Si vous voulez aussi afficher les logs dans la console
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    formatter = logging.Formatter('%(name)-12s: %(levelname)-8s %(message)s')
    console.setFormatter(formatter)
    logging.getLogger('').addHandler(console)
    logger = logging.getLogger(__name__)
    
    # chargement des parametres du traitement
    logger.info("Load configuration file....")
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
    
    # mise a jour de la table chantiers dans la bdd
    pointages = model.Pointages()
    pointages.lire_xlsx(config.pointages_path)
    bdd.upsert(pointages)    
    
    # Exemples d'utilisation
    # logger.debug("Debug message")
    # logger.info("This is an info message")
    # logger.warning("Warning: something unexpected happened")
    # logger.error("An error occurred")
    # logger.critical("Critical error, application might crash")