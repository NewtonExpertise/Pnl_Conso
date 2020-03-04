from tkinter import *
from tkinter import messagebox
from tkinter.ttk import Combobox
import actions
from quadraenv import QuadraSetEnv
from espion import update_espion
from datetime import datetime
from time import strftime, strptime
import locale
import os
import configparser as cp

config = cp.ConfigParser()
config.read('conf_operateur_pnl.ini',encoding="utf-8")
path_ipl = config.get('Path', 'path_ipl')

locale.setlocale(locale.LC_TIME,'')

class Application(Frame):

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.couleur = "#E4AB5B"
        self.pack()
        self.create_widgets()

    def create_widgets(self):

        # Création des widgets
        self.var_dossiers = StringVar()
        self.var_dossiers.set("")
        self.lab_dossier = Label(self, text="Dossiers \ud83d\udd0e",font=('Helvetica', 12) , foreground='orange')
        self.lab_select_dossier = Label(self, text="\ud83d\udc49",font=('Helvetica', 14, 'bold'),foreground='orange')
        self.lab_baseN = Label(self, text="N : \ud83d\udc49",font=('Helvetica', 14, 'bold'),foreground='orange')
        self.lab_base1N = Label(self, text="N-1 : \ud83d\udc49",font=('Helvetica', 14, 'bold'),foreground='orange')
        self.lab_select_action = Label(self, text="\ud83d\udc49",font=('Helvetica', 14,'bold'),foreground='orange')
   
        self.saisie1 = Entry(self, width=25, textvariable=self.var_dossiers, cursor='question_arrow')
        self.saisie1.focus_set()
        self.liste_dossiers = Listbox(self, width=25, height=8, selectbackground=self.couleur, cursor="hand2")
        self.liste_basesN = Listbox(self, width=25, height=5, selectbackground=self.couleur, cursor="hand2")
        self.liste_basesN.config(height=0)

        self.liste_bases1N = Listbox(self, width=25, height=5, selectbackground=self.couleur, cursor="hand2")
        self.liste_bases1N.config(height=0)

        self.liste_actions = Listbox(self,width=25, selectbackground=self.couleur, cursor="hand2")
        self.liste_actions.config(height=0)
    
        # Positions
        self.lab_dossier.grid(row=0, column=0, padx=10, pady=3,sticky='e')
        self.saisie1.grid(row=0, column=1, padx=10, pady=3)
        self.liste_dossiers.grid(row=1, column=1, padx=10, pady=3)
        self.lab_select_dossier.grid(row=1, column=0, padx=10, pady=3,sticky='e')
        self.liste_basesN.grid(row=2, column=1, padx=10, pady=3)
        self.liste_bases1N.grid(row=3, column=1, padx=10, pady=3)
        self.liste_actions.grid(row=4, column=1, padx=10, pady=3)

        # Actions Binding
        self.liste_dossiers.bind("<ButtonRelease-1>", self.makeListeBase)
        self.liste_basesN.bind("<ButtonRelease-1>", self.setMdbPathN)
        self.liste_bases1N.bind("<ButtonRelease-1>", self.setMdbPath1N)
        self.liste_actions.bind("<ButtonRelease-1>", self.setAction)

        
        # Callback pour filtrage de la liste dossiers
        self.var_dossiers.trace("w", lambda name, index,
                                mode: self.filter_list_dossier())

        # Dictionnaires des actions
        self.dispatch = {
            actions.PnL_consolide.__name__: actions.PnL_consolide,
           
        }
        for i, action in enumerate(self.dispatch.keys()):
            self.liste_actions.insert(i, action)

        self.liste_actions.configure(state=DISABLED)
        self.makeListeDossier()
        self.filter_list_dossier()

    def setMdbPathN(self, e):
        """
        Définit le chemin complet vers la base qcompta (mdb)
        """
        self.lab_select_action.grid(row=3, column=0, padx=10, pady=3,sticky='e')
        index, = self.liste_basesN.curselection()
        self.base = self.liste_basesN.get(index)
        for base, chemin in self.dbList:
            if self.base == base:
                self.mdbN = chemin
        if self.mdbN and self.mdb1N:
            self.liste_actions.configure(state=NORMAL)

    def setMdbPath1N(self, e):
        """
        Définit le chemin complet vers la base qcompta (mdb)
        """
        self.lab_select_action.grid(row=3, column=0, padx=10, pady=3,sticky='e')
        index, = self.liste_bases1N.curselection()
        self.base = self.liste_bases1N.get(index)
        for base, chemin in self.dbList:
            if self.base == base:
                self.mdb1N = chemin
        if self.mdbN and self.mdb1N:
            self.liste_actions.configure(state=NORMAL)


    def makeListeDossier(self):
        """
        Prépare le liste des dossiers
        """
        self.dicDossier = {}
        ipl = path_ipl
        self.qenv = QuadraSetEnv(ipl)
        for code, rs in self.qenv.gi_list_clients():
            label = f"{rs} ({code})"
            self.dicDossier.update({label: code})

    def makeListeBase(self, e):
        """
        Prépare la liste des bases (DC, DA, ...)
        """
        self.mdb1N = False
        self.mdbN = False
        self.liste_actions.configure(state=DISABLED)
        self.lab_select_action.grid_forget()
   
    
        self.lab_baseN.grid(row=2, column=0, padx=10, pady=3,sticky='e')
        self.lab_base1N.grid(row=3, column=0, padx=10, pady=3,sticky='e')
        self.liste_basesN.delete(0, END)
        self.liste_bases1N.delete(0, END)
        index, = self.liste_dossiers.curselection()
        value = self.liste_dossiers.get(index)
        self.code_dossier = self.dicDossier[value]
        self.dbList = self.qenv.recent_cpta(self.code_dossier, depth=3)
        for i, (base, _) in enumerate(self.dbList):
            self.liste_basesN.insert(i, base)
            self.liste_bases1N.insert(i, base)

    def filter_list_dossier(self):
        """
        Filtrage auto de la liste des dossiers
        """
        search_term = self.var_dossiers.get()
        lbox_list = [x for x in self.dicDossier.keys()]
        self.liste_dossiers.delete(0, END)
        for item in lbox_list:
            if search_term.lower() in item.lower():
                self.liste_dossiers.insert(END, item)

    def setAction(self, e):
        """
        Sélection du programmes qui sera lancé
        """
        # self.liste_actions.configure(state='disabled')
        index, = self.liste_actions.curselection()
        value = self.liste_actions.get(index)

        self.dispatch[value](self.mdbN)
        self.dispatch[value](self.mdb1N)
        messagebox.showinfo("Annonce", "Export terminé")
        update_espion(self.code_dossier, self.base, f"operateur;{value}")
        sys.exit()
    
    def setAction_periode(self, e):
        """
        Sélection du programmes qui sera lancé avec une période choisie
        """
        # mois sélectionné :
        select_mois = self.combobox_periode.get()
        print(select_mois)
        select_mois = actions.end_of_month(datetime.strptime(select_mois, "%Y-%B"))
        self.dispatch[self.select_action](self.mdbN, select_mois)
        messagebox.showinfo("Annonce", "Export terminé")
        update_espion(self.code_dossier, self.base, f"operateur;{self.select_action}")
        sys.exit()


root = Tk()
root.title('Opérateur Excel')
root.wm_attributes("-topmost", 1)
ressources = os.path.dirname(sys.argv[0])
root.iconbitmap(os.path.join(ressources,"IMG/favicon.png"))
app = Application(master=root)
app.mainloop()
