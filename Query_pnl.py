from mdbagent import MdbConnect
import configparser as cp
import os
import sys

ressources = os.path.dirname(sys.argv[0])
config = cp.ConfigParser()
try:
    config.read(os.path.join(ressources,'conf_operateur_pnl.ini'),encoding="utf-8")
except Exception as e :
    print(str(e))
path_GI = config.get('Path', 'path_GI')
path_ipl = config.get('Path', 'path_ipl')



"""

    Query sur les dossiers clients (Raison sociale, groupement de société, dates du dossier, etc ...)

"""

def get_raison_social(mdbpath):
    sql = """
    SELECT RaisonSociale
    FROM Dossier1
    """
    with MdbConnect(mdbpath) as mdb:
        RaisonSociale = mdb.query(sql)[0][0]
    return RaisonSociale

def MacDo_Groupe(mdbpath):
    """
    retourn une liste de clients appartenant au même groupe
    """
    code_client = os.path.basename(os.path.dirname(mdbpath))
    sql = f"""
    SELECT c.Code, CodeRegroupement
    FROM Clients as c
    WHERE CodeRegroupement = (SELECT CodeRegroupement FROM Clients WHERE Code = '{code_client}')
    AND c.DateSortie = #30/12/1899#
    AND c.Code in (SELECT i.Code FROM Intervenants as i WHERE i.Enseigne = 'MAC DO')
    """
    Groupe = []
    with MdbConnect(path_GI) as mdb:
        data = mdb.query(sql)
    for codeDossier in data:
        Groupe.append(codeDossier[0])
    Nom =  data[0][1]
    return Groupe, Nom


def get_periode_exercice(QcomptaC):
    """
    Renvoie la listes des mois de l'exercice.
    """
    sql = """
    SELECT DebutExercice, FinExercice, DateLimiteSaisie
    FROM Dossier1
    """
    with MdbConnect(QcomptaC) as mdb:
        periode = mdb.query(sql)
    
    for debut, fin , limite in periode:
        periode_ex = {"debut":debut, "fin":fin, "limite":limite}
    return periode_ex



    



"""

    Query sur les données comptables des dossiers.

"""

def ectriture_analytique(mdbpath, fin_exercice, Client):
    """
    Ecrit dans une feuilles excel défini par Sheet_name les écritures analytique d'un dossier.
    """
    sql = f"""
    SELECT
        ''''&E.CodeJournal AS Journal,
        DateSerial(Year(E.PeriodeEcriture), Month(E.PeriodeEcriture), E.JourEcriture) AS DateEcr,
        ''''&E.NumeroCompte AS Compte, ''''&E.Libelle as Libelle, E.MontantTenuDebit AS Debit, E.MontantTenuCredit AS Credit,
        (E.MontantTenuDebit-E.MontantTenuCredit) AS Solde,
        ''''&E.NumeroPiece AS Piece, A.Centre, ''''&E.RefImage as RefImage,
        ''''&E.CodeOperateur AS Oper, E.DateSysSaisie as DateSysSaisie, ''''&e.TypeLigne as TypeLigne, '{Client}' AS CLIENT
    FROM
        (
            SELECT
                TypeLigne, NumUniq, NumeroCompte, CodeJournal,  Folio, LigneFolio,
                PeriodeEcriture, JourEcriture, NumLigne, Libelle, MontantTenuDebit, MontantTenuCredit,
                NumeroPiece, CodeOperateur, DateSysSaisie, RefImage
            FROM Ecritures
            WHERE TypeLigne='E'
            AND (NumeroCompte LIKE '6%' OR NumeroCompte LIKE '7%')
            AND PeriodeEcriture <= #{fin_exercice}#) E
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
    data = [list(d) for d in data]
    if data:
        data.insert(0, headers)
    return data


def PnL_data_analytique(mdbpath, cloture, code_dossier):
    sql = f"""
    SELECT '01' as Ligne, '000954' as Societe,'001 - Ventes nettes total' as Poste, (SUM(MontantTenuDebit)-SUM(MontantTenuCredit)) AS Montant
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='002' OR Centre='042')
    UNION
    SELECT '02', '000954' as codeclient, '002 - Ventes de produits alimentaires', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='002')
    UNION
    SELECT '03', '000954' as codeclient, '003 - Food Cost : Achat nourriture', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='003')
    UNION
    SELECT '04', '000954' as codeclient,'004 - Food Cost : Repas', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='004')
    UNION
    SELECT '05',  '000954' as codeclient,'005 - Food Cost : Déchets', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='005')
    UNION
    SELECT '06',  '000954' as codeclient,'006 - COUT TOTAL DE LA NOURRITURE', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='003' OR Centre='004' OR Centre='005')
    UNION
    SELECT '07',  '000954' as codeclient,'007 - Paper', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='007')
    UNION
    SELECT '08',  '000954' as codeclient,'008 - COUT TOTAL DES PRODUITS VENDUS', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='007' OR Centre='003' OR Centre='004' OR Centre='005')
    UNION
    SELECT '09',  '000954' as codeclient,'009 - BENEFICE BRUT', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='002' OR Centre='007' OR Centre='003' OR Centre='004' OR Centre='005')
    UNION
    SELECT '10',  '000954' as codeclient,'010 - Main d '&''''&'œuvre équipiers', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='010')
    UNION
    SELECT '11',  '000954' as codeclient,'011 - Salaires managers', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='011')
    UNION
    SELECT '12',  '000954' as codeclient,'012 - Charges sociales managers', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%'
    OR NumeroCompte LIKE '7%')
    AND (Centre='012')
    UNION
    SELECT '13',  '000954' as codeclient,'013 - Frais de voyage', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='013')
    UNION
    SELECT '14',  '000954' as codeclient,'014 - Publicité GIE', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='014')
    UNION
    SELECT '15',  '000954' as codeclient,'015 - Promotion locale', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='015')
    UNION
    SELECT '16',  '000954' as codeclient,'016 - Services extérieurs', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures
    WHERE TypeLigne='A'
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%'
    OR NumeroCompte LIKE '7%')
    AND (Centre='016')
    UNION
    SELECT '17',  '000954' as codeclient,'017 - Uniformes', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='017')
    UNION
    SELECT '18',  '000954' as codeclient,'018 - Fournitures d '&''''&'exploitation', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='018')
    UNION
    SELECT '19',  '000954' as codeclient,'019 - Entretien et réparations d '&''''&'équipement', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='019')
    UNION
    SELECT '20',  '000954' as codeclient,'020 - Electricité gaz téléphone eau', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='020')
    UNION
    SELECT '21',  '000954' as codeclient,'021 - Fournitures de bureau', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='021')
    UNION
    SELECT '22',  '000954' as codeclient,'022 - Ecarts de caisse', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='022')
    UNION
    SELECT '23',  '000954' as codeclient,'023 - Divers', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='023')
    UNION
    SELECT '24',  '000954' as codeclient,'TOTAL DEPENSES  CONTROLABLES' , SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre = '010' OR Centre = '011' OR 
        Centre = '012' OR Centre = '013' OR 
        Centre = '014' OR Centre = '015' OR 
        Centre = '016' OR Centre = '017' OR 
        Centre = '018' OR Centre = '019' OR 
        Centre = '020' OR Centre = '021' OR 
        Centre = '022' OR Centre = '023')
    UNION
    SELECT '25',  '000954' as codeclient,  '024 - P.A.C.', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre = '002' OR Centre = '007' OR
        Centre = '003' OR Centre = '004' OR
        Centre = '005' OR Centre = '010' OR
        Centre = '011' OR Centre = '012' OR
        Centre = '013' OR Centre = '014' OR
        Centre = '015' OR Centre = '016' OR
        Centre = '017' OR Centre = '018' OR
        Centre = '019' OR Centre = '020' OR
        Centre = '021' OR Centre = '022' OR
        Centre = '023')
    UNION
    SELECT '26',  '000954' as codeclient,'030 - Redevance standard', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='030')
    UNION
    SELECT '27',  '000954' as codeclient,'031 - Redevance services', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='031')
    UNION
    SELECT '28',  '000954' as codeclient,'032 - Frais comptables et juridiques', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='032')
    UNION
    SELECT '29',  '000954' as codeclient,'033 - Assurance', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='033')
    UNION
    SELECT '30',  '000954' as codeclient,'034 - Taxes et permis', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='034')
    UNION
    SELECT '31',  '000954' as codeclient,'035 - Perte (gain) cession d '&''''&'actif', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='035')
    UNION
    SELECT '32',  '000954' as codeclient,'036 - Dépréciation amortissement', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='036')
    UNION
    SELECT '33',  '000954' as codeclient,'037 - Frais financiers et charges d '&''''&'intérêts', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='037')
    UNION
    SELECT '34',  '000954' as codeclient,'038 - Revenus financiers', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='038')
    UNION
    SELECT '35',  '000954' as codeclient,'039 - Dépenses (revenus) divers', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='039')
    UNION
    SELECT '36',  '000954' as codeclient,'040 - TOTAL DEPENSES NON CONTROLABLES', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre = '030' OR Centre = '031' OR 
        Centre = '032' OR Centre = '033' OR 
        Centre = '034' OR Centre = '035' OR 
        Centre = '036' OR Centre = '037' OR 
        Centre = '038' OR Centre = '039')
    UNION
    SELECT '37',  '000954' as codeclient,'041 - TOTAL DES DEPENSES', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre ='007' OR Centre ='003' OR 
        Centre = '004' OR Centre = '005' OR
        Centre = '010' OR Centre = '011' OR 
        Centre = '012' OR Centre = '013' OR 
        Centre = '014' OR Centre = '015' OR 
        Centre = '016' OR Centre = '017' OR 
        Centre = '018' OR Centre = '019' OR 
        Centre = '020' OR Centre = '021' OR 
        Centre = '022' OR Centre = '023' OR
        Centre = '030' OR Centre = '031' OR 
        Centre = '032' OR Centre = '033' OR 
        Centre = '034' OR Centre = '035' OR 
        Centre = '036' OR Centre = '037' OR 
        Centre = '038' OR Centre = '039'
        )
    UNION
    SELECT '38',  '000954' as codeclient,'042 - Ventes non-alimentaires', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='042')
    UNION
    SELECT '39',  '000954' as codeclient,'043 - Coûts non-alimentaires', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='043')
    UNION
    SELECT '40',  '000954' as codeclient,'044 - RESULTAT NET NON-ALIMENTAIRE', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='042' OR Centre='043')
    UNION
    SELECT '41',  '000954' as codeclient,'045 - REVENU NET D '&''''&'EXPLOITATION', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre = '042' OR Centre = '043' OR
        Centre = '030' OR Centre = '031' OR
        Centre = '032' OR Centre = '033' OR
        Centre = '034' OR Centre = '035' OR
        Centre = '036' OR Centre = '037' OR
        Centre = '038' OR Centre = '039' OR
        Centre = '002' OR Centre = '007' OR
        Centre = '003' OR Centre = '004' OR
        Centre = '005' OR Centre = '010' OR
        Centre = '011' OR Centre = '012' OR
        Centre = '013' OR Centre = '014' OR
        Centre = '015' OR Centre = '016' OR
        Centre = '017' OR Centre = '018' OR
        Centre = '019' OR Centre = '020' OR
        Centre = '021' OR Centre = '022' OR
        Centre = '023')
    UNION
    SELECT '42',  '000954' as codeclient,'050 - Salaire locataire-gérant', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='050')
    UNION
    SELECT '43',  '000954' as codeclient,'051 - Autres dépenses', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='051')
    UNION
    SELECT '44',  '000954' as codeclient,'052 - Dépenses de bureau', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='052')
    UNION
    SELECT '45',  '000954' as codeclient,'TOTAL DES FRAIS D '&''''&'ADMINISTRATION', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='050' OR Centre='051' OR Centre='052')
    UNION
    SELECT '46',  '000954' as codeclient,'053 - REVENU NET AVANT IMPOTS', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre = '042' OR Centre = '043' OR
        Centre = '030' OR Centre = '031' OR
        Centre = '032' OR Centre = '033' OR
        Centre = '034' OR Centre = '035' OR
        Centre = '036' OR Centre = '037' OR
        Centre = '038' OR Centre = '039' OR
        Centre = '002' OR Centre = '007' OR
        Centre = '003' OR Centre = '004' OR
        Centre = '005' OR Centre = '010' OR
        Centre = '011' OR Centre = '012' OR
        Centre = '013' OR Centre = '014' OR
        Centre = '015' OR Centre = '016' OR
        Centre = '017' OR Centre = '018' OR
        Centre = '019' OR Centre = '020' OR
        Centre = '021' OR Centre = '022' OR
        Centre = '023' OR Centre = '050' OR
        Centre='051' OR Centre='052')
    UNION
    SELECT '47',  '000954' as codeclient,'054 - Impôt sur les sociétés', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre='054')
    UNION
    SELECT '48',  '000954' as codeclient,'REVENU NET APRES IMPOTS', SUM(MontantTenuDebit)-SUM(MontantTenuCredit)
    FROM Ecritures 
    WHERE TypeLigne='A' 
    AND PeriodeEcriture<=#{cloture}#
    AND (NumeroCompte LIKE '6%' 
    OR NumeroCompte LIKE '7%') 
    AND (Centre = '042' OR Centre = '043' OR
        Centre = '030' OR Centre = '031' OR
        Centre = '032' OR Centre = '033' OR
        Centre = '034' OR Centre = '035' OR
        Centre = '036' OR Centre = '037' OR
        Centre = '038' OR Centre = '039' OR
        Centre = '002' OR Centre = '007' OR
        Centre = '003' OR Centre = '004' OR
        Centre = '005' OR Centre = '010' OR
        Centre = '011' OR Centre = '012' OR
        Centre = '013' OR Centre = '014' OR
        Centre = '015' OR Centre = '016' OR
        Centre = '017' OR Centre = '018' OR
        Centre = '019' OR Centre = '020' OR
        Centre = '021' OR Centre = '022' OR
        Centre = '023' OR Centre = '050' OR
        Centre = '051' OR Centre = '052' OR
        Centre='054')
    """
    # Récupération data
    with MdbConnect(mdbpath) as mdb:
        info, data = mdb.queryInfoData(sql)
    headers = [x[0] for x in info]
    data.insert(0, headers)

    return data


def Balance_annuel_soldes_analytiques(mdbpath):
    """
    Renvoie vers le tableur la balance annuel des soldes annalytiques
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

