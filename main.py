#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import shutil
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

class CNCManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Programme CNC Manager - Session : admin [Admin]")
        self.root.geometry("1100x700")
        self.root.minsize(900, 600)

        # Style général
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # En-tête du logiciel
        header_frame = tk.Frame(self.root, bg="#1a3d6c", height=40)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        title_label = tk.Label(
            header_frame, 
            text="PROGRAMME CNC MANAGER", 
            bg="#1a3d6c", 
            fg="white", 
            font=("Arial", 12, "bold")
        )
        title_label.pack(pady=8)

        # Barre de menu supérieure
        menubar = tk.Menu(self.root)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Quitter", command=self.root.quit)
        menubar.add_cascade(label="Fichier", menu=file_menu)

        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Paramètres de communication", command=lambda: messagebox.showinfo("Info", "Module paramètres RS232"))
        menubar.add_cascade(label="Outils", menu=tools_menu)

        admin_menu = tk.Menu(menubar, tearoff=0)
        admin_menu.add_command(label="Gestion des utilisateurs", command=lambda: messagebox.showinfo("Admin", "Connecté en tant qu'Administrateur"))
        menubar.add_cascade(label="Administration", menu=admin_menu)

        self.root.config(menu=menubar)

        # Création du Notebook (Onglets principaux)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Initialisation des différents onglets
        self.init_tab_list()
        self.init_tab_comparison()
        self.init_tab_backup()
        self.init_tab_of()
        self.init_tab_traceability()
        self.init_tab_transfer()

    def init_tab_list(self):
        """1. Liste programme usinage"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="1. Liste programme usinage")
        
        lbl = tk.Label(tab, text="Répertoire des programmes ISO / G-code", font=("Arial", 11, "bold"))
        lbl.pack(anchor="w", padx=10, pady=10)

        # Tableau des programmes
        columns = ("Nom du fichier", "Date de modification", "Taille", "Machine ciblée")
        self.tree_programs = ttk.Treeview(tab, columns=columns, show="headings")
        for col in columns:
            self.tree_programs.heading(col, text=col)
            self.tree_programs.column(col, width=200)
        
        self.tree_programs.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Données fictives d'exemple
        sample_data = [
            ("0021A.iso", "2026-09-15 10:30", "14 Ko", "NUM 1060"),
            ("P_SURF_01.cnc", "2026-09-14 16:22", "32 Ko", "NUM 1060"),
            ("FOIL_CARBON.iso", "2026-09-12 09:15", "45 Ko", "NUM 1060")
        ]
        for item in sample_data:
            self.tree_programs.insert("", tk.END, values=item)

    def init_tab_comparison(self):
        """2. Comparaison de Dossiers / Fichiers (Style WinMerge)"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="2. Comparaison de Dossiers / Fichiers")

        # Sous-onglets ou sous-cadres de comparaison
        comp_notebook = ttk.Notebook(tab)
        comp_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Sous-onglet : Comparaison côte-à-côte
        sub_tab_file = ttk.Frame(comp_notebook)
        comp_notebook.add(sub_tab_file, text="Comparaison côte-à-côte de Fichiers (Style WinMerge)")

        selection_frame = tk.LabelFrame(sub_tab_file, text="Sélection des fichiers G-code à comparer")
        selection_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(selection_frame, text="Fichier A :").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_file_a = tk.Entry(selection_frame, width=45)
        self.entry_file_a.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(selection_frame, text="Parcourir", command=self.browse_file_a).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(selection_frame, text="Fichier B :").grid(row=0, column=3, sticky="w", padx=5, pady=5)
        self.entry_file_b = tk.Entry(selection_frame, width=45)
        self.entry_file_b.grid(row=0, column=4, padx=5, pady=5)
        tk.Button(selection_frame, text="Parcourir", command=self.browse_file_b).grid(row=0, column=5, padx=5, pady=5)

        action_frame = tk.Frame(sub_tab_file)
        action_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Button(action_frame, text="▶ Comparer Fichiers Texte", bg="#e1f5fe", command=self.compare_files).pack(side=tk.LEFT, padx=5)
        tk.Button(action_frame, text="▲ Différence Précédente", state=tk.DISABLED).pack(side=tk.LEFT, padx=5)
        tk.Button(action_frame, text="▼ Différence Suivante", state=tk.DISABLED).pack(side=tk.LEFT, padx=5)
        
        self.lbl_status_comp = tk.Label(action_frame, text="Aucune comparaison", fg="gray")
        self.lbl_status_comp.pack(side=tk.RIGHT, padx=5)

        # Zone d'affichage des différences
        text_frame = tk.Frame(sub_tab_file)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.text_comparison = tk.Text(text_frame, wrap=tk.NONE, font=("Courier New", 10))
        scrollbar_y = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.text_comparison.yview)
        scrollbar_x = ttk.Scrollbar(text_frame, orient=tk.HORIZONTAL, command=self.text_comparison.xview)
        
        self.text_comparison.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.text_comparison.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def init_tab_backup(self):
        """3. Configuration & Planification des Sauvegardes"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="3. Configuration & Planification des Sauvegardes")
        
        frame = tk.LabelFrame(tab, text="Paramètres de sauvegarde automatique (Cobian / Interne)")
        frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)

        tk.Label(frame, text="Dossier Source des programmes :").grid(row=0, column=0, sticky="w", padx=10, pady=10)
        self.entry_src = tk.Entry(frame, width=50)
        self.entry_src.grid(row=0, column=1, padx=10, pady=10)
        self.entry_src.insert(0, "E:/sauvegarde-usinage")

        tk.Label(frame, text="Dossier de Destination (Archive) :").grid(row=1, column=0, sticky="w", padx=10, pady=10)
        self.entry_dest = tk.Entry(frame, width=50)
        self.entry_dest.grid(row=1, column=1, padx=10, pady=10)
        self.entry_dest.insert(0, "D:/Backup_CNC_Archive")

        tk.Button(frame, text="Lancer une sauvegarde immédiate", bg="#c8e6c9", command=self.run_backup).grid(row=2, column=1, sticky="w", padx=10, pady=20)

    def init_tab_of(self):
        """4. Ordres de Fabrication (OF)"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="4. Ordres de Fabrication (OF)")
        tk.Label(tab, text="Gestion et Suivi des Ordres de Fabrication (Sidi Bou Ali)", font=("Arial", 11)).pack(padx=10, pady=20)

    def init_tab_traceability(self):
        """5. Traçabilité & Suivi"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="5. Traçabilité & Suivi")
        tk.Label(tab, text="Journal des modifications et historique des transferts machine", font=("Arial", 11)).pack(padx=10, pady=20)

    def init_tab_transfer(self):
        """6. Transfert CNC (RS232 / NUM 1060)"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="6. Transfert CNC (RS232 / NUM 1060)")
        
        frame = tk.Label(tab, text="Interface de liaison série avec le contrôleur NUM 1060", font=("Arial", 11, "bold"))
        frame.pack(padx=10, pady=15)

        btn_send = tk.Button(tab, text="Envoyer le programme vers la CNC", bg="#ffccbc", width=30, height=2)
        btn_send.pack(pady=10)

        btn_recv = tk.Button(tab, text="Recevoir depuis la CNC (DNC)", bg="#b2dfdb", width=30, height=2)
        btn_recv.pack(pady=10)

    def browse_file_a(self):
        filename = filedialog.askopenfilename(initialdir="E:/sauvegarde-usinage")
        if filename:
            self.entry_file_a.delete(0, tk.END)
            self.entry_file_a.insert(0, filename)

    def browse_file_b(self):
        filename = filedialog.askopenfilename(initialdir="E:/sauvegarde-usinage")
        if filename:
            self.entry_file_b.delete(0, tk.END)
            self.entry_file_b.insert(0, filename)

    def compare_files(self):
        path_a = self.entry_file_a.get()
        path_b = self.entry_file_b.get()

        if not path_a or not path_b:
            messagebox.showwarning("Attention", "Veuillez sélectionner les deux fichiers à comparer.")
            return

        try:
            with open(path_a, 'r', encoding='utf-8', errors='ignore') as fa:
                lines_a = fa.readlines()
            with open(path_b, 'r', encoding='utf-8', errors='ignore') as fb:
                lines_b = fb.readlines()

            self.text_comparison.delete("1.0", tk.END)
            
            max_len = max(len(lines_a), len(lines_b))
            diff_count = 0

            for i in range(max_len):
                line_a = lines_a[i].strip() if i < len(lines_a) else "<fin de fichier>"
                line_b = lines_b[i].strip() if i < len(lines_b) else "<fin de fichier>"

                if line_a == line_b:
                    formatted_line = f"{i+1:04d} | {line_a:<40} | {line_b}\n"
                    self.text_comparison.insert(tk.END, formatted_line)
                else:
                    diff_count += 1
                    formatted_line = f"{i+1:04d} * {line_a:<40} * {line_b}\n"
                    self.text_comparison.insert(tk.END, formatted_line)

            self.lbl_status_comp.config(text=f"{diff_count} différence(s) trouvée(s)")
            messagebox.showinfo("Succès", "Comparaison terminée avec succès.")

        except Exception as e:
            messagebox.showerror("Erreur", fImpossible de lire les fichiers :\n{e}")

    def run_backup(self):
        src = self.entry_src.get()
        dest = self.entry_dest.get()
        try:
            if not os.path.exists(dest):
                os.makedirs(dest)
            
            # Copie simple du dossier
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_folder = os.path.join(dest, f"backup_{timestamp}")
            shutil.copytree(src, backup_folder)
            messagebox.showinfo("Sauvegarde", f"Sauvegarde réalisée avec succès dans :\n{backup_folder}")
        except Exception as e:
            messagebox.showerror("Erreur Sauvegarde", fUne erreur est survenue :\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CNCManagerApp(root)
    root.mainloop()
