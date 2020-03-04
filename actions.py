import os
from Query_pnl import get_raison_social, MacDo_Groupe, get_periode_exercice, ectriture_analytique
import xlwings as xw
from datetime import timedelta,datetime
from mdbagent import MdbConnect
import re
import configparser as cp
from quadraenv import QuadraSetEnv
from collections import defaultdict
import sys

class Operateur_PNL():

    def __init__(self):

        self.ws = xw.sheets.active
        self.wb = self.ws.book
        ressources = os.path.dirname(sys.argv[0])
        config = cp.ConfigParser()
        try:
            config.read(os.path.join(ressources,'conf_operateur_pnl.ini'),encoding="utf-8")
        except Exception as e :
            print(str(e))
        self.path_ipl = config.get('Path', 'path_ipl')
        # appelation des dossiers comptable
        self.dossier_annuel = config.get('millesime', 'dossier_annuel')
        self.dossier_archive = config.get('millesime', 'dossier_archive')
        # Appelation des feuilles excels
        self.ws_ecritureN = config.get('sheets', 'ws_ecritureN')
        self.ws_ecritureN1 = config.get('sheets', 'ws_ecritureN1')
        self.ws_conso = config.get('sheets', 'ws_conso')
        self.TCDN = config.get('sheets', 'TCDN')
        self.TCDN1 = config.get('sheets', 'TCDN1')

        self.info_processN = config.get('info_traitement', 'N')
        self.info_processN1 = config.get('info_traitement', 'N1')

        self.ColCentreN1 = config.get('Plage_cells', 'ColCentreN1')
        self.ColSoldeN1 = config.get('Plage_cells', 'ColSoldeN1')
        self.EcritureN1 = config.get('Plage_cells', 'EcritureN1')
        self.ColEnseigneN1 = config.get('Plage_cells', 'ColEnseigneN1')
        self.ColCentreN = config.get('Plage_cells', 'ColCentreN')
        self.ColSoldeN = config.get('Plage_cells', 'ColSoldeN')
        self.EcritureN = config.get('Plage_cells', 'EcritureN')
        self.ColEnseigneN = config.get('Plage_cells', 'ColEnseigneN')
        self.FiltreEnseigneN = config.get('Plage_cells', 'FiltreEnseigneN')
        self.FiltreEnseigneN1 = config.get('Plage_cells', 'FiltreEnseigneN1')

    def PnL_consolide(self, mdbpath, sheet_name):
        """
        Alimente le fichier excel des P&L Conso
        """
        bases_location_firme = self.get_group_location(mdbpath)
        _, nom_groupe = MacDo_Groupe(mdbpath)
        ws_conso = self.wb.sheets[self.ws_conso]
        Nb_enseigne = len(bases_location_firme)
        i=0
        try:
            for client, infos in bases_location_firme.items():
                i+=1
                self.conso_ectriture_analytique(infos['path_bdd'], infos['fin_exercice'], sheet_name, infos['raison_social'], nom_groupe)
                if sheet_name == self.ws_ecritureN:
                    ws_conso.range("K6").number_format='0,00%'
                    ws_conso.range("K6").value = i/Nb_enseigne
                elif sheet_name == self.ws_ecritureN1:
                    ws_conso.range("M6").number_format='0,00%'
                    ws_conso.range("M6").value = i/Nb_enseigne

        except KeyError as e:
            print(e)

    def get_group_location(self, mdbpath):
        # 1 on récupère la date de cloture du dossier sélectionné.
        date_exercice = get_periode_exercice(mdbpath)

        # Etablissement d'une plage permettant de capter l'ensemble des dossiers sur l'année civile d'un groupe
        debut_plage = datetime(year=date_exercice['fin'].year, month=1, day=1)
        fin_plage = self.end_of_month(datetime(year=date_exercice['fin'].year, month=12, day=1))

        # on récupère les codes dossiers du groupe sélectionné.
        groupe_mcdo, _ = MacDo_Groupe(mdbpath)

        # création d'un défaut dict contenant bdd_name et le path correspondant au code client.
        bases_location_firme = defaultdict(list)
        for client in groupe_mcdo:
            liste_tuple_base_path = self.makeListeDossier(client) # on récupère le chemin vers la BDD
            for base, path in liste_tuple_base_path:
                bases_location_firme[client].append((base, path))
        bases_location_firme = dict(bases_location_firme) # conversion du defaultdict en dict

        for clients, BDD in bases_location_firme.items():
            for bdd_name, path_bdd in BDD:
                if self.dossier_annuel in os.path.basename(os.path.dirname(os.path.dirname(path_bdd))).upper() or self.dossier_archive in os.path.basename(os.path.dirname(os.path.dirname(path_bdd))).upper():
                    date_exercice = get_periode_exercice(path_bdd)  # on récupère les périodes d'exercies
                    raison_social = get_raison_social(path_bdd)

                    if debut_plage < date_exercice['fin'] and date_exercice['fin'] <= fin_plage:
                        # mise a jour du doctionnaire
                        bases_location_firme[clients] = {
                            'fin_exercice':date_exercice['fin'],
                            'path_bdd': path_bdd,
                            'raison_social' : raison_social}

        return bases_location_firme

    def clear_pnl_conso(self ):
        """
        Réinitialise le PNL Conso.
        """
        ws_conso = self.wb.sheets[self.ws_conso]
        ws_conso.range("J6").clear()
        ws_conso.range("L6").clear()
        ws_conso.range("K6").value = 0
        ws_conso.range("M6").value = 0
        ws_conso.range("J5").clear()
        ws_conso.range("J7").clear()
        ws_conso.range("A1").clear()
        ws_conso.range("B2").clear()

        self.wb.sheets[self.ws_ecritureN].clear()
        self.wb.sheets[self.ws_ecritureN1].clear()
        nbligne = max(ws_conso.cells(self.ws.api.rows.count, "J").end(-4162).row+1, ws_conso.cells(self.ws.api.rows.count, "K").end(-4162).row+1)
        ws_conso.range("J12:K"+str(nbligne)).clear()

    def conso_ectriture_analytique(self, mdbpath, fin_exercice, sheet_name, Client, Nom_groupe):
        """
        Ecrit dans une feuilles excel défini par Sheet_name les écritures analytique d'un dossier.
        """

        data = ectriture_analytique(mdbpath, fin_exercice, Client)

        if data:

            ws_conso = self.wb.sheets[self.ws_conso]
            ws_conso.range("A1").value = Nom_groupe
            if sheet_name == self.ws_ecritureN:
                ws_conso.range("B2").value = fin_exercice

            # insertion du listing des sociétées traitées dans la page P&L conso pour information.
            if sheet_name == self.ws_ecritureN:
                nbligne = ws_conso.cells(self.ws.api.rows.count, "J").end(-4162).row +1
                ws_conso.range('J'+str(nbligne)).value = Client
            elif sheet_name == self.ws_ecritureN1:
                nbligne = ws_conso.cells(self.ws.api.rows.count, "K").end(-4162).row +1
                ws_conso.range('K'+str(nbligne)).value = Client

            try:
                ws_E = self.wb.sheets.add(sheet_name)
            except:
                ws_E = self.wb.sheets[sheet_name]

            nbligne = ws_E.cells(self.ws.api.rows.count, "A").end(-4162).row +1
            if nbligne >3:

                ws_E.range('A'+str(nbligne)).value = data[1:]
                ws_E.autofit()
            else:
                # formatage
                ws_E.range('H:K').number_format='@'
                ws_E.range('L:L').number_format='jj/mm/aaaa'
                ws_E.range('C:C').number_format='@'
                ws_E.range('E:G').number_format='# ##0,00'
                ws_E.range('A1').value = data
                ws_E.range('A:N').api.AutoFilter(VisibleDropDown=True)
                ws_E.autofit()

    def set_plage_cellule_pnl_conso(self):
        """
        Etablissement des plage de valeur
        """
        ws_conso = self.wb.sheets[self.ws_conso]
        ws_en = self.wb.sheets[self.ws_ecritureN]
        ws_e1n = self.wb.sheets[self.ws_ecritureN1]
        ws_TCDN = self.wb.sheets[self.TCDN]
        ws_TCD1N = self.wb.sheets[self.TCDN1]    

        nbligne = max(ws_conso.cells(self.ws.api.rows.count, "J").end(-4162).row,ws_conso.cells(self.ws.api.rows.count, "K").end(-4162).row)
        # set plages enseigne name
        ws_conso.range('L12:L'+str(nbligne)).name = self.FiltreEnseigneN
        ws_conso.range('M12:M'+str(nbligne)).name = self.FiltreEnseigneN1
        # set plages N-1
        nbligne = ws_e1n.cells(self.ws.api.rows.count, "A").end(-4162).row
        ws_e1n.range('I1:I'+str(nbligne)).name = self.ColCentreN1
        ws_e1n.range('G1:G'+str(nbligne)).name = self.ColSoldeN1
        ws_e1n.range('N1:N'+str(nbligne)).name = self.ColEnseigneN1
        ws_e1n.range('A1:N'+str(nbligne)).name = self.EcritureN1
        # set plage N
        nbligne = ws_en.cells(self.ws.api.rows.count, "A").end(-4162).row
        ws_en.range('I1:I'+str(nbligne)).name = self.ColCentreN
        ws_en.range('G1:G'+str(nbligne)).name = self.ColSoldeN
        ws_en.range('N1:N'+str(nbligne)).name = self.ColEnseigneN
        ws_en.range('A1:N'+str(nbligne)).name = self.EcritureN
        # refresh pivot Table
        try:
            ws_TCDN.api.PivotTables(self.TCDN).PivotCache().Refresh()
            
        except Exception as e:
            print(e)
        try:
            ws_TCD1N.api.PivotTables(self.TCDN1).PivotCache().Refresh()
        except Exception as e:
            print(e)

        
    def end_of_month(self, dt0):
        """
        Renvoi le dernier jour du mois de la date donnée
        prend un datetime objet
        """
        dt1 = dt0.replace(day=1)
        dt2 = dt1 + timedelta(days=32)
        dt3 = dt2.replace(day=1) - timedelta(days=1)
        return dt3

    def add_sheet_new_name(self, wb, nom):
        """
        Génère une feuille excel avec un nom unique
        nb : une feuille excel ne peut contenir que 31 caractères
        """
        nom = nom[:29]
        increment = 0
        sheet = [sheet.name for sheet in wb.sheets]
        if nom in sheet:

            for sheet_name in sheet:
                try:
                    name, compteur , _ =  re.split(nom+r'(\d+)',sheet_name) 
                except ValueError:
                    compteur =0
                if increment< int(compteur):
                    increment = int(compteur)

            new_name = nom + str(increment+1)
        else:
            new_name = nom 
        new_ws = wb.sheets.add(new_name)
        new_ws.book.activate(True)

        return new_ws

    def makeListeDossier(self, code_dossier):
        """
        Prépare la liste des dossiers
        """
        qenv = QuadraSetEnv(self.path_ipl)
        dbList = qenv.recent_cpta(code_dossier, depth=3)

        return dbList


if __name__ == "__main__":
    import pprint
    import xlwings as xw
    pp=pprint.PrettyPrinter(indent=4)
    mdb=  r'\\srvquadra\Qappli\Quadra\DATABASE\cpta\DC\000424\qcompta.mdb'
    mdb2 = r'\\srvquadra\Qappli\Quadra\DATABASE\cpta\DA2018\000424\qcompta.mdb'
    Qgi = r'\\srvquadra\Qappli\Quadra\DATABASE\gi\0000\Qgi.mdb'
    # pp.pprint(get_mois_exercice(mdb))
    # import xlwings as xw
    # ws = xw.sheets.active
    # wb=ws.book
    # pp.pprint(makeListeDossier('000954'))
    # pp.pprint(PnL_data(mdb, '31/12/2019', '000954'))
    # PnL_consolide(mdb)
    # data = ecritures_analytiques(mdb)
    # print(get_raison_social(mdb))
    # data = PnL_data_groupe(mdb, '31/12/2019', '000954')
    # data = [list(data) for data in data]
    # ws = xw.sheets.active
    # xw.Range("A1").value=data

    # print(MacDo_Groupe(mdb))

    OP = Operateur_PNL()

    sheet_nameN = "Ecriture N"
    sheet_name1N = "Ecriture N-1"
    print('start clear')
    OP.clear_pnl_conso()
    print('end cleat')
    OP.PnL_consolide(mdb ,sheet_nameN)
    OP.PnL_consolide(mdb2, sheet_name1N)
    OP.set_plage_cellule_pnl_conso()