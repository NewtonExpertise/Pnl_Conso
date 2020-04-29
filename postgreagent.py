import psycopg2
import configparser
import logging


class PostgreAgent:
    """
    Pilotage d'une connexion a une db postgresql
    les paramètres sont passés dans le fichier .ini
    """
    def __init__(self, args):
        self.connection = ""
        self.cursor = ""
        self.args = args
        logging.debug(f"Connexion postgre")

    def _connect(self):
        try:
            self.connection = psycopg2.connect(**self.args)
            self.cursor = self.connection.cursor()
            dbname = self.connection.get_dsn_parameters()["dbname"]
            logging.debug(f"acces à la db {dbname}")
        except (Exception, psycopg2.InterfaceError) as error:
            logging.error(f"Echec connexion : {error}")
            return 0
        except (Exception, psycopg2.DatabaseError) as error:
            logging.error(f"Echec database : {error}")
            return 0

    def _close(self):
        logging.debug("fermeture de la connexion à postgre")
        if self.connection:
            self.connection.commit()
            self.connection.close()

    def __enter__(self):
        """Pour context manager"""
        self._connect()
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self._close()


    def query(self, sql):
        data = []
        try:
            self.cursor.execute(sql)
        except (Exception, psycopg2.Error) as error:
            logging.error(f"Echec requête : {error}")
            return data
        for row in self.cursor:
            data.append([x for x in row])
        return data
    
    def execute(self, sql):
        try:
            self.cursor.execute(sql)
            return True
        except (Exception, psycopg2.Error) as error:
            logging.error(f"Echec requête : {error}")
            return False

    def query(self, sql):
        data = []
        try:
            self.cursor.execute(sql)
            return True
        except (Exception, psycopg2.Error) as error:
            logging.error(f"Echec requête : {error}")
            return False

    def manage(self, sql):
        try:
            logging.debug("Running query")
            self.cursor.execute(sql)
        except (Exception, psycopg2.Error) as error:
            logging.error(f"Echec requête : {error}")
            return 0

    def table_exists(self, table):
        val = False
        sql = """
        SELECT table_name FROM information_schema.tables where table_schema='public' ;
        """
        self.cursor.execute(sql)
        table_list = [x[0] for x in self.cursor.fetchall()]
        if table in table_list:
            val = True
        return val


if __name__ == "__main__":

    from collections import OrderedDict

    logging.basicConfig(
        level=logging.DEBUG,
        format="%(module)s %(funcName)s\t%(levelname)s - %(message)s",
    )
    table = "paie_compteurs"

    conf = OrderedDict([('host', '10.0.0.17'), ('user', 'admin'), ('password', 'Zabayo@@'), ('port', '5432'), ('dbname', 'newton')])
    with PostgreAgent(conf) as db:
        if db.connection:
            if db.table_exists(table):
                logging.debug(f"table {table} exists")
