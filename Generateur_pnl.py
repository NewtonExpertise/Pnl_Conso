from quadraenv import QuadraSetEnv
import xlwings as xw
from mdbagent import MdbConnect
import os
from espion import update_espion
import win32ui



def pnl_excel():
    """

    """

    ipl =  r"\\srvquadra\qappli\quadra\database\client\quadra.ipl"
    Q = QuadraSetEnv(ipl)
    ws = xw.sheets.active
    
    code_client = ws.range('O2').value
    try :
        code_client = str(int(code_client))
    except:
        pass
    if code_client and len(code_client)<7:
        code_client = code_client.zfill(6)
        raison_social = Q.get_rs(code_client)
        ws.range("A1").value = raison_social
        dossierN = ws.range('O4').value
        dossier1N = ws.range('Q4').value

        bases = Q.recent_cpta(dossier=code_client, depth=3)
        bases_name = [ base[0] for base in bases]
        str_bases_name = ';'.join(bases_name)

        xw.Range('O4').api.validation.delete()
        xw.Range('O4').clear()
        xw.Range('O4').api.validation.add(3,1,3,str_bases_name)

        xw.Range('Q4').api.validation.delete()
        xw.Range('Q4').clear()
        xw.Range('Q4').api.validation.add(3,1,3,str_bases_name)
        path_N = False
        path_1N = False
        if dossierN and dossierN != "Dossier N":
            for nom, path in bases:
                if nom == dossierN:
                    path_N =  path
            ecritures_analytiques(path_N,"Ecritures_N")
            codes_journaux(path_N)
            xw.Range('B2').value = end_exercice(path_N)
            xw.Range('E1').value = end_exercice(path_N)
        else:
            ws.range('O4').value = "Dossier N"

        if dossier1N and dossier1N!="Dossier N-1":
            for nom, path in bases:
                if nom == dossier1N:
                    path_1N =  path
            ecritures_analytiques(path_1N,"Ecritures_N-1")
            xw.Range('E2').value = end_exercice(path_1N)
            
        else:
            ws.range('Q4').value = "Dossier N-1"

        if path_N or path_1N:
            win32ui.MessageBox("Fin de traitement", "Information")

        if path_N and path_1N:
            update_espion(code_client, dossierN+' - '+dossier1N , "pnl_excel_clt_bilan")
        elif path_N:
            update_espion(code_client, dossierN , "pnl_excel_clt_bilan")
        

    else:
        ws.range('O2').value = "Num client"




def PnL_consolide(mdbpath):
    """
    Renvoie vers le tableur la balance analytique de l'exercice sélectionné
    """
    sql = """
    SELECT 
    CUM.Centre AS CodeAna, 
    CUM.Solde AS SoldeCumul 
    FROM 
    (SELECT 
    Centre, 
    SUM(MontantTenuDebit) - SUM(MontantTenuCredit) AS Solde 
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture>=#2019-01-01# 
    AND PeriodeEcriture<=#2019-12-31# 
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    GROUP BY Centre) CUM 
    """
    
    # Récupération data
    with MdbConnect(mdbpath) as mdb:
        info, data = mdb.queryInfoData(sql)
    headers = [x[0] for x in info]
    data.insert(0, headers)

    return data



def ecritures_analytiques(mdbpath, sheet_name):
    """
    Renvoie vers le tableur la listes des écritures analytiques
    """

    sql = f"""
    SELECT
        E.CodeJournal AS Journal, 
        DateSerial(Year(E.PeriodeEcriture), Month(E.PeriodeEcriture), E.JourEcriture) AS DateEcr,
        E.NumeroCompte AS Compte, E.Libelle, E.MontantTenuDebit AS Debit, E.MontantTenuCredit AS Credit, 
        (E.MontantTenuDebit-E.MontantTenuCredit) AS Solde,
        E.NumeroPiece AS Piece, A.Centre, E.RefImage, E.CodeOperateur AS Oper, E.DateSysSaisie
    FROM
        (
            SELECT 
                TypeLigne, NumUniq, NumeroCompte, CodeJournal,  Folio, LigneFolio, 
                PeriodeEcriture, JourEcriture, NumLigne, Libelle, MontantTenuDebit, MontantTenuCredit, 
                NumeroPiece, CodeOperateur, DateSysSaisie, RefImage 
            FROM Ecritures 
            WHERE TypeLigne='E' 
            AND (NumeroCompte LIKE '6%' OR NumeroCompte LIKE '7%')) E
    LEFT JOIN
        (
            SELECT 
                TypeLigne, CodeJournal, Folio, LigneFolio, PeriodeEcriture, JourEcriture, NumLigne, Centre 
            FROM Ecritures WHERE TypeLigne='A') A
    ON E.CodeJournal=A.CodeJournal
    AND E.Folio=A.Folio
    AND E.LigneFolio=A.LigneFolio
    AND E.PeriodeEcriture=A.PeriodeEcriture
    """
    # Récupération data
    with MdbConnect(mdbpath) as mdb:
        info, data = mdb.queryInfoData(sql)
    headers = [x[0] for x in info]
    data.insert(0, headers)
    
    bnligne=len(data)
    ws = xw.sheets.active
    wb = ws.book
    ws_E = wb.sheets(sheet_name)

    # formatage
    ws_E.clear()
    ws_E.range('H:K').number_format='@'
    ws_E.range('L:L').number_format='jj/mm/aaaa'
    ws_E.range('C:C').number_format='@'
    ws_E.range('E:G').number_format='# ##0,00'
    ws_E.range('A1').value = data
    ws_E.autofit()

def codes_journaux(mdbpath):
    sql="""
    SELECT Code from Journaux ORDER BY Code;
    """
    with MdbConnect(mdbpath) as mdb:
        data = mdb.query(sql)
    set1 = {x[0] for x in data}

    # Requête sur la base des paramètres généraux QcomptaC
    drive, _ = os.path.splitdrive(mdbpath)
    QcomptaC = os.path.abspath(os.path.join(drive, "quadra/database/cpta/qcomptac.mdb"))
    with MdbConnect(QcomptaC) as mdb:
        data = mdb.query(sql)
    set2 = {x[0] for x in data}

    fullset = sorted(set1.union(set2))
    bnligne=len(fullset) +10
    xw.Range('N10:N60').clear()
    xw.Range('O10:O60').clear()
    xw.Range('N10:N'+str(bnligne)).name = 'FiltreJourn1'
    xw.Range('O10:O'+str(bnligne)).name = 'FiltreJourn2'
    xw.Range('N10').options(transpose = True).value = fullset

def end_exercice(mdbpath):
    sql = """
    SELECT DebutExercice, FinExercice, DateLimiteSaisie
    FROM Dossier1
    """
    with MdbConnect(mdbpath) as mdb:
        periode = mdb.query(sql)
    fin_exercice = periode[0][1]
    return fin_exercice
    

    

if __name__ == "__main__":
    import pprint
    pp = pprint.PrettyPrinter(indent=4)

    path = r'\\srvquadra\Qappli\Quadra\DATABASE\cpta\DC\T00752\QCompta.mdb'
    pp.pprint(balance_analytique(path))