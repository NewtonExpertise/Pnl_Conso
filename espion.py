import getpass
import logging
from collections import OrderedDict
from datetime import datetime
from postgreagent import PostgreAgent


def update_espion(dossier = "", base = "", argStr = ""):

    conf = OrderedDict(
        [
            ('host', '10.0.0.17'), 
            ('user', 'admin'), 
            ('password', 'Zabayo@@'), 
            ('port', '5432'), 
            ('dbname', 'outils')
            ]
        )

    horodat = datetime.now()
    collab = getpass.getuser()
    table = "espion"
    values = [collab, horodat, dossier, base, argStr]

    sql = """
    INSERT INTO espion (collab, horodat, dossier, base, args) 
    VALUES (%s, %s, %s, %s, %s);
    """

    with PostgreAgent(conf) as db:
        if db.connection:
            if db.table_exists(table):
                logging.debug(f"table {table} exists")        
                db.cursor.execute(sql, values)

# if __name__ == "__main__":
#     import logging
#     update_espion("FORM05", "abc")

