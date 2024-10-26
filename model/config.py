import configparser

class Config:
    def __init__(self):
        self.host = None
        self.port = None
        self.password = None
        self.bdd = None
        self.employes_path = None
        
    def load(self, configPath):
        # chargement du fichier de conf
        config = configparser.ConfigParser()
        config.read(configPath)
        # chargement des variables de la base de données
        self.host = config.get('DATABASE', 'host')
        self.port = config.get('DATABASE', 'port')
        self.password = config.get('DATABASE', 'password')
        self.bdd = config.get('DATABASE', 'bdd')
        self.employes_path = config.get('EXCEL_FILES', 'employes')
        self.chantiers_path = config.get('EXCEL_FILES', 'chantiers')
        self.pointages_path = config.get('EXCEL_FILES', 'pointages')