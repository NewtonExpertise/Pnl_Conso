import logging
import os
import sys
import pyodbc
from datetime import datetime


class QuadraSetEnv(object):
    def __init__(self, ipl_file):

        self.cpta = ""
        self.paie = ""
        self.gi = ""
        self.conn = ""
        self.cur = ""

        with open(ipl_file, "r") as f:
            lines = f.readlines()

        for line in lines:
            line = line.rstrip().replace("\\", "/")
            if "=" in line:
                key, item = line.split("=")[0:2]
                if key == "RACDATACPTA":
                    self.cpta = item.upper()
                elif key == "RACDATAPAIE":
                    self.paie = item.upper()
                elif key == "RACDATAGI":
                    self.gi = item.upper()
        if self.cpta:
            logging.debug("Acces fichier ipl OK")

    def make_db_path(self, type_dossier, num_dossier):
        type_dossier = type_dossier.upper()
        num_dossier = num_dossier.upper()
        db_path = ""
        if (
            type_dossier == "DC"
            or type_dossier.startswith("DA")
            or type_dossier.startswith("DS")
        ):
            db_path = os.path.join(self.cpta, type_dossier, num_dossier, "qcompta.mdb")

        elif type_dossier == "PAIE":
            db_path = os.path.join(self.paie, num_dossier, "qpaie.mdb")
        return os.path.abspath(db_path)

    def chemins_cpta(self, categ="D", tail=""):
        """
        Renvoie la liste des chemin menant vers les dossiers comptables découverts
        dans le dossier database/cpta (dc, archives, situations)
        tail : permet de rajouter un nom de fichier/dossier à la fin du chemin
        """
        liste = []
        with os.scandir(self.cpta) as it1:
            for entry in it1:
                name = entry.name.upper()
                if name.startswith(categ) and entry.is_dir():
                    base = os.path.join(self.cpta, name)
                    with os.scandir(base) as it2:
                        for entry in it2:
                            if entry.is_dir():
                                final = os.path.abspath(os.path.join(base, entry.name))
                                liste.append(final)

        if tail:
            liste_chemins_ = [os.path.join(x, tail) for x in liste]
            liste_chemins = [x for x in liste_chemins_ if os.path.isfile(x)]
        else:
            liste_chemins = liste

        return liste_chemins
    
    def chemins_paie(self, bannis=[]):
        """
        Renvoi la liste des chemins vers les bases paie
        bannis = permet d'ajouter une liste de dossiers exclus
        """
        liste = []
        with os.scandir(self.paie) as itr:
            dossList = [x.name for x in itr]
        accepted = [x for x in dossList if x not in bannis]
        
        for item in accepted:
            qpaie = os.path.abspath(os.path.join(self.paie, item, "qpaie.mdb"))
            if os.path.isfile(qpaie):
                liste.append(qpaie)
        return liste

    def gi_list_clients(self, bannis=False, sortis=False):

        mdb_path = os.path.join(self.gi, "0000", "qgi.mdb")
        constr = "Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + mdb_path
        logging.debug("openning qgi : {}".format(mdb_path))

        if sortis:
            adden = "C.DateSortie>=#1899-12-30# "
        else:
            adden = "C.DateSortie=#1899-12-30# "

        sql = """
            SELECT I.Code, I.Nom 
            FROM Intervenants I 
            INNER JOIN Clients C 
            ON I.Code=C.Code 
            WHERE I.IsClient='1' 
            AND {}
            ORDER BY I.Nom
            """.format(adden)
        try:
            self.conn = pyodbc.connect(constr, autocommit=True)
            self.cur = self.conn.cursor()
            self.cur.execute(sql)
            data = list(self.cur)
        except pyodbc.Error:
            logging.error(
                ("erreur requete base {} \n {}".format(mdb_path, sys.exc_info()[1]))
            )
            return False

        if not bannis:
            purge = [x for x in data if x[0] not in self.bannis()]
            return purge

        return data

    def get_rs(self, code_dossier):
        rs = ""
        liste_dossiers = self.gi_list_clients()
        for code, nom in liste_dossiers:
            if code_dossier == code:
                rs = nom
        return rs

    def recent_cpta(self, dossier="", depth=0, tail="qcompta.mdb"):
        """
        Renvoie la liste des chemins vers les bases d'un dossier spécifique
        Cela comprend DC, DA et DS
        L'historique est limité par l'argument depth
        """
        liste = []
        anneeASonder = [str(datetime.now().year - x) for x in range(0,depth)]

        with os.scandir(self.cpta) as it1:
            for entry in it1:
                nameMill = entry.name.upper()
                for annee in anneeASonder:
                    if annee in nameMill :
                        base = os.path.join(self.cpta, nameMill)
                        with os.scandir(base) as it2:
                            for entry in it2:
                                if (entry.is_dir() and dossier == entry.name.upper()):
                                    liste.append((
                                        nameMill,
                                        os.path.abspath(os.path.join(base, entry.name, tail))))
        liste.sort(reverse=True)
        dcpath = self.make_db_path("DC", dossier)
        if os.path.isfile(dcpath):
            liste.insert(0, ("DC", dcpath))
        return liste        

    def bannis(self):
        bannis = [
                "000000",
                "000075",
                "000101", # NEWTON EXPERTISE
                "000175",
                "000206", # RECORD SHOP
                "000272", # DRAKKAR
                "000284",
                "000402",
                "000423",
                "000431",
                "000455",
                "000490",
                "000500", # NEWTON AUDIT
                "000511",
                "000689",
                "000700", # A&M EXPERTISE
                "000745",
                "016113",
                "016352",
                "016586",
                "016695",
                "016709",
                "016714",
                "016717",
                "ADTEST",
                "ASUP06",
                "COP334",
                "demo",
                "DEMO14",
                "FORM07",
                "FORM0X",
                "IMPORT",
                "NRTEST",
                "N00101",
                "PMCDO1",
                "PMCDON",
                "propre",
                "TPRDLO",
                "TSTJLF",
                "TSTVPN",
                "TYPE",
                "TYPE1",
                "X00101",
                "X00396",
                "X00691",
                "X00705",
                "X00752",
                "X00844",
                "X00882",
                "X00886",
                "X00949",
                "X00953",
                "X0JOHN",
                "X15091",
                "XBATIM",
                "XOUTIL",
                "Y00670",
                "Z00130",
                "ZOUTAA",
                "ZRESTO"
            ]
        return bannis    

# if __name__ == "__main__":
#     ipl = r"\\srvquadra\qappli\quadra\database\client\quadra.ipl"
#     Q = QuadraSetEnv(ipl)
#     gilist = Q.gi_list_clients(sortis=False)
#     listeDossier = {}
#     for code, rs in gilist:
#         listeDossier.update({f"{rs} ({code})" : code}) 
#     print(listeDossier)
    # for code, nom in gilist:
    #     if nom.startswith("DRAK"):
    #         print(code, nom)

if __name__ == "__main__":
    ipl = r"\\srvquadra\qappli\quadra\database\client\quadra.ipl"
    Q = QuadraSetEnv(ipl)
    for item in (Q.recent_cpta("000181", 3)):
        print(item)