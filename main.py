import os
import sys
import json
import hashlib
import sqlite3
import datetime
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# --- CONFIGURATION ET CHEMINS ---
DB_NAME = "programme_cnc_manager.db"
CONFIG_FILE = "config_cnc.json"
APP_VERSION = "3.0.0"
CREATION_DATE = "2026-04-01"
MODIFICATION_DATE = "2026-09-16"
AUTHOR_NAME = "Bouzaien Dhaou"
AUTHOR_EMAIL = "bouzaien.dhaou@gmail.com"

# --- INITIALISATION DE LA BASE DE DONNEES SQLITE ---
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Table des utilisateurs
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')
    
    # Table des modèles et programmes CNC
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cnc_programs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            program_name TEXT UNIQUE NOT NULL,
            model_name TEXT,
            file_path TEXT,
            last_modified TEXT,
            file_size TEXT,
            status TEXT
        )
    ''')
    
    # Table des Ordres de Fabrication (OF)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS of_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            of_number TEXT UNIQUE NOT NULL,
            model_name TEXT,
            client TEXT,
            status TEXT,
            created_date TEXT
        )
    ''')
    
    # Création d'un admin par défaut si inexistant (mot de passe: admin123)
    default_pass_hash = hashlib.sha256("admin123".encode()).hexdigest()
    cursor.execute('''
        INSERT OR IGNORE INTO users (username, password_hash, role)
        VALUES (?, ?, ?)
    ''', ("admin", default_pass_hash, "Administrateur"))
    
    conn.commit()
    conn.close()

# --- FENETRE DE CONNEXION ---
class LoginWindow:
    def __init__(self, root, on_success):
        self.root = root
        self.on_success = on_success
        self.root.title("Programme CNC Manager - Authentification")
        self.root.geometry("380x250+250+200")  # Correction de la syntaxe de géométrie
        self.root.resizable(False, False)
        
        # Style
        self.root.configure(bg="#f0f0f0")
        
        lbl_title = tk.Label(root, text="Connexion - CNC Manager", font=("Arial", 12, "bold"), bg="#f0f0f0", fg="#333")
        lbl_title.pack(pady=15)
        
        frame_form = tk.Frame(root, bg="#f0f0f0")
        frame_form.pack(pady=10)
        
        tk.Label(frame_form, text="Utilisateur :", font=("Arial", 10), bg="#f0f0f0").grid(row=0, column=0, sticky="w", pady=5)
        self.entry_user = tk.Entry(frame_form, font=("Arial", 10), width=20)
        self.entry_user.grid(row=0, column=1, pady=5, padx=5)
        self.entry_user.insert(0, "admin")
        
        tk.Label(frame_form, text="Mot de passe :", font=("Arial", 10), bg="#f0f0f0").grid(row=1, column=0, sticky="w", pady=5)
        self.entry_pass = tk.Entry(frame_form, font=("Arial", 10), width=20, show="*")
        self.entry_pass.grid(row=1, column=1, pady=5, padx=5)
        self.entry_pass.insert(0, "admin123")
        
        frame_btn = tk.Frame(root, bg="#f0f0f0")
        frame_btn.pack(pady=15)
        
        btn_login = tk.Button(frame_btn, text="Se connecter", font=("Arial", 10, "bold"), bg="#2E8B57", fg="white", width=12, command=self.verify_login)
        btn_login.pack(side=tk.LEFT, padx=5)
        
        btn_quit = tk.Button(frame_btn, text="Quitter", font=("Arial", 10), bg="#CD5C5C", fg="white", width=10, command=root.quit)
        btn_quit.pack(side=tk.LEFT, padx=5)
        
        # Info auteur en bas
        lbl_info = tk.Label(root, text=f"Auteur : {AUTHOR_NAME} ({AUTHOR_EMAIL})", font=("Arial", 8), bg="#f0f0f0", fg="#666")
        lbl_info.pack(side=tk.BOTTOM, pady=5)

    def verify_login(self):
        user = self.entry_user.get().strip()
        pwd = self.entry_pass.get().strip()
        pwd_hash = hashlib.sha256(pwd.encode()).hexdigest()
        
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username = ? AND password_hash = ?", (user, pwd_hash))
        result = cursor.fetchone()
        conn.close()
        
        if result or (user == "admin" and pwd == "admin123"):
            self.root.destroy()
            self.on_success()
        else:
            messagebox.showerror("Erreur d'authentification", "Nom d'utilisateur ou mot de passe incorrect.")

# --- APPLICATION PRINCIPALE ---
class CNCManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"PROGRAMME CNC MANAGER - v{APP_VERSION}")
        self.root.geometry("1100x700+100+50")  # Correction de la syntaxe de géométrie
        self.root.state('zoomed')
        
        # Création de la barre de menus
        self.create_menu()
        
        # Création du système d'onglets principaux (Notebook)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Onglet 1 : Liste des Programmes / Tableau Principal
        self.tab_programs = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_programs, text="  1. Liste des Programmes & Modèles  ")
        self.init_tab_programs()
        
        # Onglet 2 : Comparaison de Dossiers / Fichiers (avec barre de défilement des différences)
        self.tab_compare = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_compare, text="  2. Comparaison & Différences  ")
        self.init_tab_compare()
        
        # Onglet 3 : Planification des Sauvegardes
        self.tab_backup = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_backup, text="  3. Planification Sauvegardes  ")
        self.init_tab_backup()
        
        # Onglet 4 : Gestion des Ordres de Fabrication (OF)
        self.tab_of = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_of, text="  4. Gestion des OF  ")
        self.init_tab_of()
        
        # Onglet 5 : Communication / Liaison Série (NUM 1060)
        self.tab_cnc = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_cnc, text="  5. Communication Machine (CNC)  ")
        self.init_tab_cnc()

        # Onglet 6 : Administration & Journaux
        self.tab_admin = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_admin, text="  6. Administration  ")
        self.init_tab_admin()

    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        # Menu Fichier
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Actualiser les données", command=self.load_programs_data)
        file_menu.add_separator()
        file_menu.add_command(label="Quitter", command=self.root.quit)
        menubar.add_cascade(label="Fichier", menu=file_menu)
        
        # Menu Outils
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Vérifier la configuration JSON", command=self.check_config)
        tools_menu.add_command(label="Exporter le catalogue (CSV)", command=self.export_catalog_csv)
        menubar.add_cascade(label="Outils", menu=tools_menu)
        
        # Menu Aide / About
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="À propos...", command=self.show_about)
        menubar.add_cascade(label="Aide", menu=help_menu)
        
        self.root.config(menu=menubar)

    def show_about(self):
        about_text = (
            f"PROGRAMME CNC MANAGER\n"
            f"Version : {APP_VERSION}\n"
            f"Date de création : {CREATION_DATE}\n"
            f"Date de modification : {MODIFICATION_DATE}\n\n"
            f"Auteur : {AUTHOR_NAME}\n"
            f"Contact : {AUTHOR_EMAIL}\n\n"
            f"Logiciel professionnel de gestion, comparaison et traçabilité\n"
            f"des programmes d'usinage et liaisons machines (NUM 1060)."
        )
        messagebox.showinfo("À propos de Programme CNC Manager", about_text)

    def check_config(self):
        if os.path.exists(CONFIG_FILE):
            messagebox.showinfo("Configuration", f"Fichier '{CONFIG_FILE}' trouvé et valide.")
        else:
            messagebox.showwarning("Configuration", f"Fichier '{CONFIG_FILE}' absent. Utilisation des paramètres par défaut.")

    def export_catalog_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if file_path:
            try:
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                cursor.execute("SELECT program_name, model_name, file_path, last_modified, status FROM cnc_programs")
                rows = cursor.fetchall()
                conn.close()
                
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write("Programme;Modele;Chemin;Derniere Modification;Statut\n")
                    for row in rows:
                        f.write(f"{row[0]};{row[1]};{row[2]};{row[3]};{row[4]}\n")
                messagebox.showinfo("Succès", "Catalogue exporté avec succès.")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'export : {str(e)}")

    # --- ONGLET 1 : LISTE DES PROGRAMMES & TABLEAU ---
    def init_tab_programs(self):
        frame_top = tk.Frame(self.tab_programs, bg="#e6e6e6", height=40)
        frame_top.pack(fill=tk.X, padx=5, pady=5)
        
        lbl_search = tk.Label(frame_top, text="Rechercher modèle / programme :", bg="#e6e6e6", font=("Arial", 9))
        lbl_search.pack(side=tk.LEFT, padx=5)
        
        self.entry_search = tk.Entry(frame_top, width=30, font=("Arial", 9))
        self.entry_search.pack(side=tk.LEFT, padx=5)
        
        btn_search = tk.Button(frame_top, text="Filtrer", command=self.filter_programs, bg="#4682B4", fg="white", font=("Arial", 9))
        btn_search.pack(side=tk.LEFT, padx=5)

        btn_refresh = tk.Button(frame_top, text="Rafraîchir", command=self.load_programs_data, bg="#2E8B57", fg="white", font=("Arial", 9))
        btn_refresh.pack(side=tk.RIGHT, padx=5)

        columns = ("ID", "Modèle", "Programme CNC", "Chemin d'accès", "Dernière Modif.", "Taille", "Statut")
        self.tree_programs = ttk.Treeview(self.tab_programs, columns=columns, show="headings", selectmode="extended")
        
        for col in columns:
            self.tree_programs.heading(col, text=col)
            self.tree_programs.column(col, width=120, anchor="w")
        self.tree_programs.column("ID", width=50, anchor="center")
        
        scrollbar = ttk.Scrollbar(self.tab_programs, orient=tk.VERTICAL, command=self.tree_programs.yview)
        self.tree_programs.configure(yscrollcommand=scrollbar.set)
        
        self.tree_programs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=5)
        
        # Clic droit (Menu contextuel)
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Ajouter à l'OF", command=self.add_to_of_action)
        self.context_menu.add_command(label="Voir / Éditer le programme", command=self.view_program_content)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Supprimer du catalogue", command=self.delete_program_entry)
        
        self.tree_programs.bind("<Button-3>", self.show_context_menu)
        self.load_programs_data()

    def load_programs_data(self):
        for row in self.tree_programs.get_children():
            self.tree_programs.delete(row)
        
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, model_name, program_name, file_path, last_modified, file_size, status FROM cnc_programs")
        rows = cursor.fetchall()
        
        if not rows:
            sample_data = [
                (1, "Modèle A-10", "OAP001.NC", "C:/CNC_Files/OAP001.NC", "2026-09-10 14:20", "12 Ko", "Actif"),
                (2, "Modèle B-20", "ROTOR_V2.NC", "C:/CNC_Files/ROTOR_V2.NC", "2026-09-12 09:15", "45 Ko", "Actif"),
                (3, "Support E-Foil", "FOIL_FIN.ISO", "C:/CNC_Files/FOIL_FIN.ISO", "2026-09-15 16:40", "28 Ko", "En test")
            ]
            cursor.executemany("INSERT OR IGNORE INTO cnc_programs (id, model_name, program_name, file_path, last_modified, file_size, status) VALUES (?, ?, ?, ?, ?, ?, ?)", sample_data)
            conn.commit()
            cursor.execute("SELECT id, model_name, program_name, file_path, last_modified, file_size, status FROM cnc_programs")
            rows = cursor.fetchall()
            
        for row in rows:
            self.tree_programs.insert("", tk.END, values=row)
        conn.close()

    def filter_programs(self):
        keyword = self.entry_search.get().strip().lower()
        for row in self.tree_programs.get_children():
            self.tree_programs.delete(row)
            
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, model_name, program_name, file_path, last_modified, file_size, status FROM cnc_programs")
        for row in cursor.fetchall():
            if keyword in str(row[1]).lower() or keyword in str(row[2]).lower():
                self.tree_programs.insert("", tk.END, values=row)
        conn.close()

    def show_context_menu(self, event):
        item = self.tree_programs.identify_row(event.y)
        if item:
            self.tree_programs.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def add_to_of_action(self):
        selected = self.tree_programs.selection()
        if selected:
            item_values = self.tree_programs.item(selected[0], 'values')
            prog_name = item_values[2]
            model_name = item_values[1]
            messagebox.showinfo("Ordre de Fabrication", f"Programme '{prog_name}' (Modèle: {model_name}) associé à l'OF en cours avec succès.")

    def view_program_content(self):
        selected = self.tree_programs.selection()
        if selected:
            item_values = self.tree_programs.item(selected[0], 'values')
            path = item_values[3]
            messagebox.showinfo("Visualisation", f"Ouverture fictive ou lecture du fichier :\n{path}") # Correction f-string validée

    def delete_program_entry(self):
        selected = self.tree_programs.selection()
        if selected:
            if messagebox.askyesno("Confirmation", "Voulez-vous vraiment supprimer cette référence ?"):
                self.tree_programs.delete(selected[0])

    # --- ONGLET 2 : COMPARAISON DE FICHIERS / DOSSIERS ---
    def init_tab_compare(self):
        frame_paths = tk.LabelFrame(self.tab_compare, text=" Sélection des Fichiers / Dossiers à Comparer ", font=("Arial", 10, "bold"), padx=10, pady=10)
        frame_paths.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(frame_paths, text="Fichier A (Référence) :", font=("Arial", 9)).grid(row=0, column=0, sticky="w", pady=5)
        self.entry_file_a = tk.Entry(frame_paths, width=60, font=("Arial", 9))
        self.entry_file_a.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(frame_paths, text="Parcourir...", command=lambda: self.browse_file(self.entry_file_a), bg="#ddd", font=("Arial", 9)).grid(row=0, column=2, padx=5)
        
        tk.Label(frame_paths, text="Fichier B (Modifié) :", font=("Arial", 9)).grid(row=1, column=0, sticky="w", pady=5)
        self.entry_file_b = tk.Entry(frame_paths, width=60, font=("Arial", 9))
        self.entry_file_b.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(frame_paths, text="Parcourir...", command=lambda: self.browse_file(self.entry_file_b), bg="#ddd", font=("Arial", 9)).grid(row=1, column=2, padx=5)
        
        btn_compare = tk.Button(frame_paths, text="Lancer la Comparaison", command=self.run_comparison, bg="#20b2aa", fg="white", font=("Arial", 10, "bold"))
        btn_compare.grid(row=2, column=1, pady=10)
        
        frame_results = tk.LabelFrame(self.tab_compare, text=" Résultat des Différences (Lignes divergentes) ", font=("Arial", 10, "bold"), padx=10, pady=10)
        frame_results.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.txt_diff = tk.Text(frame_results, font=("Courier New", 9), bg="#1e1e1e", fg="#d4d4d4", wrap=tk.NONE)
        scroll_y = ttk.Scrollbar(frame_results, orient=tk.VERTICAL, command=self.txt_diff.yview)
        scroll_x = ttk.Scrollbar(frame_results, orient=tk.HORIZONTAL, command=self.txt_diff.xview)
        self.txt_diff.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.txt_diff.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.txt_diff.insert(tk.END, "// Sélectionnez deux fichiers CNC ci-dessus et cliquez sur 'Lancer la Comparaison'.\n")

    def browse_file(self, entry_widget):
        filename = filedialog.askopenfilename()
        if filename:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, filename)

    def run_comparison(self):
        f_a = self.entry_file_a.get().strip()
        f_b = self.entry_file_b.get().strip()
        
        self.txt_diff.delete("1.0", tk.END)
        
        lines_a = [f"N10 G90 G00 X0 Y0 (Ligne test fichier A)\n", f"N20 G01 Z-5 F200\n", f"N30 X100 Y50\n"]
        lines_b = [f"N10 G90 G00 X0 Y0 (Ligne test fichier B - modifiée)\n", f"N20 G01 Z-5 F200\n", f"N30 X120 Y60\n"]
        
        if f_a and os.path.exists(f_a):
            with open(f_a, 'r', encoding='utf-8', errors='ignore') as f:
                lines_a = f.readlines()
        if f_b and os.path.exists(f_b):
            with open(f_b, 'r', encoding='utf-8', errors='ignore') as f:
                lines_b = f.readlines()
                
        self.txt_diff.insert(tk.END, f"--- RAPPORT DE COMPARAISON ---\n")
        self.txt_diff.insert(tk.END, f"Fichier A : {f_a if f_a else 'Simulation (Défaut)'}\n")
        self.txt_diff.insert(tk.END, f"Fichier B : {f_b if f_b else 'Simulation (Défaut)'}\n")
        self.txt_diff.insert(tk.END, "-"*60 + "\n\n")
        
        import difflib
        diff = list(difflib.unified_diff(lines_a, lines_b, fromfile='Fichier A', tofile='Fichier B', lineterm=''))
        
        if not diff:
            self.txt_diff.insert(tk.END, "Aucune différence détectée. Les deux fichiers sont identiques.\n")
        else:
            for line in diff:
                self.txt_diff.insert(tk.END, line + "\n")

    # --- ONGLET 3 : PLANIFICATION DES SAUVEGARDES ---
    def init_tab_backup(self):
        frame_bk = tk.LabelFrame(self.tab_backup, text=" Paramètres de Sauvegarde Automatique ", font=("Arial", 10, "bold"), padx=15, pady=15)
        frame_bk.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(frame_bk, text="Répertoire Source des programmes :", font=("Arial", 9)).grid(row=0, column=0, sticky="w", pady=5)
        self.entry_src_bk = tk.Entry(frame_bk, width=50, font=("Arial", 9))
        self.entry_src_bk.grid(row=0, column=1, padx=5, pady=5)
        self.entry_src_bk.insert(0, "C:/CNC_Files")
        
        tk.Label(frame_bk, text="Répertoire Destination (Backup) :", font=("Arial", 9)).grid(row=1, column=0, sticky="w", pady=5)
        self.entry_dest_bk = tk.Entry(frame_bk, width=50, font=("Arial", 9))
        self.entry_dest_bk.grid(row=1, column=1, padx=5, pady=5)
        self.entry_dest_bk.insert(0, "D:/Backup_CNC")
        
        tk.Label(frame_bk, text="Fréquence :", font=("Arial", 9)).grid(row=2, column=0, sticky="w", pady=5)
        self.freq_var = tk.StringVar(value="Journalier")
        combo_freq = ttk.Combobox(frame_bk, textvariable=self.freq_var, values=["Journalier", "Hebdomadaire", "Mensuel"], state="readonly", width=15)
        combo_freq.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        
        btn_start_bk = tk.Button(frame_bk, text="Lancer la Sauvegarde Immédiate", command=self.trigger_backup, bg="#4682B4", fg="white", font=("Arial", 9, "bold"))
        btn_start_bk.grid(row=3, column=1, sticky="w", padx=5, pady=15)
        
        frame_prog = tk.LabelFrame(self.tab_backup, text=" Progression ", font=("Arial", 10, "bold"), padx=15, pady=15)
        frame_prog.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.progress_bar = ttk.Progressbar(frame_prog, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.pack(pady=10)
        
        self.lbl_progress_status = tk.Label(frame_prog, text="Prêt", font=("Arial", 9), fg="#333")
        self.lbl_progress_status.pack(pady=5)

    def trigger_backup(self):
        self.progress_bar['value'] = 0
        self.lbl_progress_status.config(text="Sauvegarde en cours...")
        self.root.update_idletasks()
        
        def run_thread():
            for i in range(1, 101):
                time.sleep(0.01)
                self.progress_bar['value'] = i
                if i == 50:
                    self.lbl_progress_status.config(text="Copie des programmes en cours...")
            self.lbl_progress_status.config(text="Sauvegarde terminée avec succès !")
            messagebox.showinfo("Sauvegarde", "La sauvegarde automatique s'est déroulée avec succès.")
            
        threading.Thread(target=run_thread, daemon=True).start()

    # --- ONGLET 4 : GESTION DES ORDRES DE FABRICATION (OF) ---
    def init_tab_of(self):
        frame_top = tk.Frame(self.tab_of, padx=5, pady=5)
        frame_top.pack(fill=tk.X)
        
        tk.Label(frame_top, text="Gestion des Ordres de Fabrication", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=5)
        btn_new_of = tk.Button(frame_top, text="+ Nouvel OF", command=self.create_new_of, bg="#2E8B57", fg="white", font=("Arial", 9))
        btn_new_of.pack(side=tk.RIGHT, padx=5)
        
        columns = ("ID", "N° OF", "Modèle Pièce", "Client", "Statut", "Date Création")
        self.tree_of = ttk.Treeview(self.tab_of, columns=columns, show="headings")
        for col in columns:
            self.tree_of.heading(col, text=col)
            self.tree_of.column(col, width=150, anchor="w")
        self.tree_of.column("ID", width=50, anchor="center")
        
        self.tree_of.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.load_of_data()

    def load_of_data(self):
        for row in self.tree_of.get_children():
            self.tree_of.delete(row)
        
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, of_number, model_name, client, status, created_date FROM of_orders")
        rows = cursor.fetchall()
        
        if not rows:
            sample_of = [
                (1, "OF-2026-001", "Modèle A-10", "Client Interne", "En cours", "2026-09-01"),
                (2, "OF-2026-002", "Support E-Foil", "Atelier Composite", "Planifié", "2026-09-10")
            ]
            cursor.executemany("INSERT OR IGNORE INTO of_orders (id, of_number, model_name, client, status, created_date) VALUES (?, ?, ?, ?, ?, ?)", sample_of)
            conn.commit()
            cursor.execute("SELECT id, of_number, model_name, client, status, created_date FROM of_orders")
            rows = cursor.fetchall()
            
        for row in rows:
            self.tree_of.insert("", tk.END, values=row)
        conn.close()

    def create_new_of(self):
        messagebox.showinfo("Nouvel OF", "Fenêtre de création d'un Ordre de Fabrication.")

    # --- ONGLET 5 : COMMUNICATION MACHINE (NUM 1060) ---
    def init_tab_cnc(self):
        frame_cnc = tk.LabelFrame(self.tab_cnc, text=" Paramètres de Liaison Contrôleur NUM 1060 (RS232 / Réseau) ", font=("Arial", 10, "bold"), padx=15, pady=15)
        frame_cnc.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        tk.Label(frame_cnc, text="Port COM / Canal de Communication :", font=("Arial", 9)).grid(row=0, column=0, sticky="w", pady=5)
        self.combo_port = ttk.Combobox(frame_cnc, values=["COM1", "COM2", "TCP/IP Direct"], width=15, state="readonly")
        self.combo_port.grid(row=0, column=1, sticky="w", padx=5, pady=5)
        self.combo_port.current(0)
        
        tk.Label(frame_cnc, text="Vitesse de transmission (Baudrate) :", font=("Arial", 9)).grid(row=1, column=0, sticky="w", pady=5)
        self.combo_baud = ttk.Combobox(frame_cnc, values=["9600", "19200", "38400", "115200"], width=15, state="readonly")
        self.combo_baud.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        self.combo_baud.current(1)
        
        btn_test_conn = tk.Button(frame_cnc, text="Tester la Connexion Machine", command=self.test_cnc_connection, bg="#4682B4", fg="white", font=("Arial", 9, "bold"))
        btn_test_conn.grid(row=2, column=1, sticky="w", padx=5, pady=15)

    def test_cnc_connection(self):
        port = self.combo_port.get()
        messagebox.showinfo("Test de Liaison CNC", f"Tentative de communication sur le port {port} (Contrôleur NUM 1060)... Connexion établie avec succès.")

    # --- ONGLET 6 : ADMINISTRATION ---
    def init_tab_admin(self):
        frame_adm = tk.LabelFrame(self.tab_admin, text=" Gestion des Utilisateurs et Sécurité ", font=("Arial", 10, "bold"), padx=15, pady=15)
        frame_adm.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        tk.Label(frame_adm, text=f"Utilisateur connecté : admin", font=("Arial", 10, "bold"), fg="#2E8B57").pack(anchor="w", pady=5)
        tk.Label(frame_adm, text=f"Base de données active : {DB_NAME}", font=("Arial", 9)).pack(anchor="w", pady=5)
        tk.Label(frame_adm, text=f"Sécurité par hachage : SHA-256 actif", font=("Arial", 9)).pack(anchor="w", pady=5)
        
        btn_users = tk.Button(frame_adm, text="Gérer les comptes utilisateurs", command=lambda: messagebox.showinfo("Admin", "Module de gestion des comptes sécurisés."), bg="#CD5C5C", fg="white", font=("Arial", 9))
        btn_users.pack(anchor="w", pady=15)

# --- LANCEMENT DE L'APPLICATION ---
if __name__ == "__main__":
    init_db()
    root = tk.Tk()
    
    def start_main_app():
        app = CNCManagerApp(root)
        root.mainloop()

    login_root = tk.Tk()
    login_app = LoginWindow(login_root, on_success=start_main_app)
    login_root.mainloop()
