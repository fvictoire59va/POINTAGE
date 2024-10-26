import psycopg2
from psycopg2.extras import execute_values
import logging
from functools import wraps

class Postgresql:
    def __init__(self, config):
        self.config = config
    
    def log_decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            logging.info(f"Calling function '{func.__name__}'")
            result = func(*args, **kwargs)
            logging.info(f"Function '{func.__name__}' finished execution")
            return result
        return wrapper
    
    @log_decorator
    def upsert(self, className):
        try:
            connection = psycopg2.connect(
                user="postgres",
                password=self.config.password,
                host=self.config.host,
                port=self.config.port,
                database=self.config.bdd
            )
            UpdateExp = ""
            cursor = connection.cursor()
            print("Connexion réussie à la base de données PostgreSQL")
            
            labels = ", ".join(className.columns)
            for valeur in className.columns:
                if UpdateExp == "":
                    UpdateExp = str(valeur) + " = EXCLUDED." + str(valeur) 
                else:
                    UpdateExp = UpdateExp + ", " + str(valeur) + " = EXCLUDED." + str(valeur)
            
            upsert_command =f"""
                                INSERT INTO {className.tableName} ({labels}) VALUES %s
                                ON CONFLICT ({ ', '.join(f'{key}' for key in className.key)}) DO 
                                UPDATE SET {UpdateExp}
                            """
            # print(upsert_command)
            execute_values(cursor, upsert_command, className.data_to_upsert)
            print("L'upsert s'est executé avec succes")
            connection.commit()
            
        except (Exception, psycopg2.Error) as error:
            print("Erreur lors de la connexion à PostgreSQL", error)
            
        finally:
            if connection:
                cursor.close()
                connection.close()
                print("Connexion PostgreSQL fermée")