from quadraenv import QuadraSetEnv
import xlwings as xw
from mdbagent import MdbConnect
import os
from espion import update_espion
from actions import Operateur_PNL
import configparser as cp
import sys

OP = Operateur_PNL()
# OP.clear_pnl_conso
ressources = os.path.dirname(sys.argv[0])
config = cp.ConfigParser()
try:
    config.read(os.path.join(ressources,'conf_operateur_pnl.ini'),encoding="utf-8")
except Exception as e :
    print(str(e))
path_ipl = config.get('Path', 'path_ipl')

ws_ecritureN = config.get('sheets', 'ws_ecritureN')
ws_ecritureN1 = config.get('sheets', 'ws_ecritureN1')

no_baseN_selected = config.get('conf_sheet', 'no_baseN_selected')
no_baseN1_selected = config.get('conf_sheet', 'no_baseN1_selected')
no_client_selected = config.get('conf_sheet', 'no_client_selected')

info_processN = config.get('info_traitement', 'N')
info_processN1 = config.get('info_traitement', 'N1')
Start = config.get('info_traitement', 'debut')
End = config.get('info_traitement', 'fin')
# type d'opération pour posgres espion
operation = config.get('operation', 'conso')

def pnl_conso_excel():
    

    Q = QuadraSetEnv(path_ipl)
    ws = xw.sheets.active
    wb= ws.book
    OP.clear_pnl_conso()
    code_client = ws.range('K2').value
    try :
        code_client = str(int(code_client))
    except:
        pass
    if code_client and len(code_client)<7:
        code_client = code_client.zfill(6)

        raison_social = Q.get_rs(code_client)

        dossierN = ws.range('k4').value
        dossier1N = ws.range('M4').value

        bases = Q.recent_cpta(dossier=code_client, depth=3)
        bases_name = [ base[0] for base in bases]
        str_bases_name = ';'.join(bases_name)

        xw.Range('k4').api.validation.delete()
        xw.Range('k4').clear()
        xw.Range('k4').api.validation.add(3,1,3,str_bases_name)

        xw.Range('M4').api.validation.delete()
        xw.Range('M4').clear()
        xw.Range('M4').api.validation.add(3,1,3,str_bases_name)
        path_N = False
        path_1N = False
        if dossierN and dossierN != no_baseN_selected:
            ws.range('J5').value = Start
            for nom, path in bases:
                if nom == dossierN:
                    path_N =  path
            ws.range("J6").value = info_processN
            ws.range("K6").value = 0
            OP.PnL_consolide(path_N, ws_ecritureN)
            
        else:
            ws.range('k4').value = no_baseN_selected

        if dossier1N and dossier1N != no_baseN1_selected:
            ws.range('J5').value = Start
            for nom, path in bases:
                if nom == dossier1N:
                    path_1N =  path
            ws.range("L6").value = info_processN1
            ws.range("M6").value = 0
            OP.PnL_consolide(path_1N, ws_ecritureN1)
            
            
        else:
            ws.range('M4').value = no_baseN1_selected

        if path_N or path_1N:
            print("BINGO")
            OP.set_plage_cellule_pnl_conso()
            OP.set_manual_controle()
            ws.range("J7").value = End

        if path_N and path_1N:
            update_espion(code_client, dossierN+' - '+dossier1N)
        elif path_N:
            update_espion(code_client, dossierN)
        

    else:
        ws.range('k2').value = no_client_selected


if __name__ == "__main__":
    pnl_conso_excel()