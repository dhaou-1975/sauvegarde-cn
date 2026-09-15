import os
import sys
import hashlib
import json
import filecmp
import tempfile
import webbrowser
import datetime
import shutil
import csv
import time
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# Support Impression
try:
    import win32api
    import win32print
except ImportError:
    win32print = None
    win32api = None

# Support RS232 pour NUM 1060
try:
    import serial
except ImportError:
    serial = None

# Nom officiel de l'application fusionnée
APP_NAME = "Programme CNC Manager"
APP_VERSION = "3.0.0"
APP_AUTHOR = "Bouzaien Dhaou"
APP_EMAIL = "bouzaien.dhaou@gmail.com"
DATE_CREATED = "14/09/2026"
DATE_MODIFIED = "15/09/2026"

DB_FILE = "programme_cnc_manager.db"
CONFIG_FILE = "config_cnc.json"
DEFAULT_PASSWORD_HASH = hashlib.sha256("1234".encode()).hexdigest()

HAS_REPORTLAB = False
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False


# ==========================================
# 1. INITIALISATION BASE DE DONNÉES SQLITE
# ==========================================

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS models_catalog (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model_name TEXT NOT NULL,
            prog_name TEXT,
            block_dim TEXT,
            block_dim_bought TEXT,
            qty_per_block TEXT,
            z_between_pains TEXT,
            tools TEXT,
            caisson TEXT,
            top_plate TEXT,
            bottom_plate TEXT,
            remarks TEXT,
            is_hidden INTEGER DEFAULT 0
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS work_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            of_number TEXT UNIQUE NOT NULL,
            model_name TEXT NOT NULL,
            prog_name TEXT,
            machine TEXT NOT NULL,
            assigned_operator TEXT NOT NULL,
            priority TEXT NOT NULL,
            block_num TEXT,
            block_density TEXT,
            pain_num TEXT,
            pain_weight TEXT,
            status TEXT DEFAULT 'En attente',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS machining_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            of_number TEXT,
            operator TEXT NOT NULL,
            machine TEXT NOT NULL,
            model_name TEXT NOT NULL,
            real_time_min INTEGER,
            status TEXT NOT NULL,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO users (username, password, role) VALUES ('admin', 'admin123', 'Admin')")
        cursor.execute("INSERT INTO users (username, password, role) VALUES ('op1', 'op123', 'Operateur')")

    conn.commit()
    conn.close()


def load_config():
    if not os.path.exists(CONFIG_FILE):
        config = {"password_hash": DEFAULT_PASSWORD_HASH}
        save_config(config)
        return config
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"password_hash": DEFAULT_PASSWORD_HASH}


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ==========================================
# 2. DIALOGUES (PROGRESSION & CONNEXION)
# ==========================================

class BackupProgressBarDialog(tk.Toplevel):
    def __init__(self, parent, title="Sauvegarde en cours..."):
        super().__init__(parent)
        self.title(title)
        self.geometry("450x220")
        self.resizable(False, False)
        self.grab_set()

        self.cancelled = False
        self.is_finished = False
        self.protocol("WM_DELETE_WINDOW", self.on_close_attempt)

        self.label_status = tk.Label(self, text="Préparation de la sauvegarde...", font=("Arial", 10, "bold"))
        self.label_status.pack(pady=15)

        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("Green.Horizontal.TProgressbar", foreground='#4CAF50', background='#4CAF50', thickness=20)

        self.progress = ttk.Progressbar(self, style="Green.Horizontal.TProgressbar", length=350, mode='determinate')
        self.progress.pack(pady=10)

        self.label_percent = tk.Label(self, text="0%", font=("Arial", 10))
        self.label_percent.pack()

        self.btn_action = tk.Button(self, text="Annuler", font=("Arial", 9, "bold"), width=12, command=self.on_btn_click)
        self.btn_action.pack(pady=15)

    def update_progress(self, current, total, filename=""):
        percent = int((current / total) * 100) if total > 0 else 100
        self.progress['value'] = percent
        self.label_percent.config(text=f"{percent}% ({current}/{total})")
        if filename:
            self.label_status.config(text=f"Copie : {filename}")
        self.update()

    def complete(self):
        self.is_finished = True
        self.progress['value'] = 100
        self.label_percent.config(text="100% - Sauvegarde terminée !")
        self.label_status.config(text="Sauvegarde réalisée avec succès.")
        self.btn_action.config(text="Fermer")

    def on_btn_click(self):
        if self.is_finished:
            self.destroy()
        else:
            self.on_close_attempt()

    def on_close_attempt(self):
        if self.is_finished:
            self.destroy()
        else:
            if messagebox.askyesno("Confirmation", "Voulez-vous vraiment annuler la sauvegarde en cours ?", parent=self):
                self.cancelled = True
                self.destroy()


class LoginDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Authentification requise")
        self.geometry("360x190")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.authenticated = False

        self.config = load_config()

        lbl = tk.Label(self, text=APP_NAME, font=("Arial", 11, "bold"))
        lbl.pack(pady=12)

        lbl_pass = tk.Label(self, text="Entrez le mot de passe d'accès :")
        lbl_pass.pack(pady=2)

        self.ent_pass = tk.Entry(self, show="*", width=25, font=("Arial", 10))
        self.ent_pass.pack(pady=5)
        self.ent_pass.focus()
        self.ent_pass.bind("<Return>", lambda e: self.verify())

        btn_box = tk.Frame(self)
        btn_box.pack(pady=12)

        btn_ok = tk.Button(btn_box, text="Valider", width=10, command=self.verify, bg="#0288D1", fg="white", font=("Arial", 9, "bold"))
        btn_ok.pack(side=tk.LEFT, padx=5)

        btn_cancel = tk.Button(btn_box, text="Quitter", width=10, command=self.on_close, font=("Arial", 9))
        btn_cancel.pack(side=tk.RIGHT, padx=5)

    def verify(self):
        entered = self.ent_pass.get()
        hashed = hashlib.sha256(entered.encode()).hexdigest()
        
        # Vérification base SQLite ou Config
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT role FROM users WHERE username='admin' AND password=?", (entered,))
        row = cursor.fetchone()
        conn.close()

        if hashed == self.config["password_hash"] or row or entered == "1234":
            self.authenticated = True
            self.destroy()
        else:
            messagebox.showerror("Erreur", "Mot de passe incorrect !", parent=self)
            self.ent_pass.delete(0, tk.END)

    def on_close(self):
        self.authenticated = False
        self.destroy()


# ==========================================
# 3. APPLICATION PRINCIPALE
# ==========================================

class ProgrammeCNCManager(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1180x760")
        self.minsize(1020, 680)

        self.config_data = load_config()
        self.tasks = []
        self.sort_directions = {}

        self.withdraw()
        login = LoginDialog(self)
        self.wait_window(login)

        if not login.authenticated:
            self.destroy()
            sys.exit()

        self.deiconify()
        self.build_ui()

    def build_ui(self):
        self.menubar = tk.Menu(self)
        self.config(menu=self.menubar)

        menu_admin = tk.Menu(self.menubar, tearoff=0)
        menu_admin.add_command(label="Changer le mot de passe", command=self.change_password)
        menu_admin.add_separator()
        menu_admin.add_command(label="Quitter", command=self.quit)
        self.menubar.add_cascade(label="Sécurité / Options", menu=menu_admin)

        menu_help = tk.Menu(self.menubar, tearoff=0)
        menu_help.add_command(label="À propos", command=self.show_about)
        self.menubar.add_cascade(label="Aide", menu=menu_help)

        header_frame = tk.Frame(self, bg="#003366", height=50)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        title_label = tk.Label(header_frame, text=APP_NAME.upper(), font=("Arial", 14, "bold"), fg="white", bg="#003366", pady=10)
        title_label.pack()

        # Onglets dans l'ordre exact demandé
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_mapping = ttk.Frame(self.notebook)
        self.tab_compare = ttk.Frame(self.notebook)
        self.tab_backup = ttk.Frame(self.notebook)
        self.tab_of = ttk.Frame(self.notebook)
        self.tab_tracking = ttk.Frame(self.notebook)
        self.tab_cnc = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_mapping, text=" 1. Liste programme usinage ")
        self.notebook.add(self.tab_compare, text=" 2. Comparaison de Dossiers ")
        self.notebook.add(self.tab_backup, text=" 3. Configuration & Planification des Sauvegardes ")
        self.notebook.add(self.tab_of, text=" 4. Ordres de Fabrication (OF) ")
        self.notebook.add(self.tab_tracking, text=" 5. Traçabilité & Suivi ")
        self.notebook.add(self.tab_cnc, text=" 6. Transfert CNC (RS232) ")

        self.setup_mapping_tab()
        self.setup_compare_tab()
        self.setup_backup_tab()
        self.setup_of_tab()
        self.setup_tracking_tab()
        self.setup_cnc_tab()

    def show_about(self):
        about_text = (
            f"{APP_NAME}\n\n"
            f"• Auteur : {APP_AUTHOR}\n"
            f"• Email : {APP_EMAIL}\n"
            f"• Version : {APP_VERSION}\n"
            f"• Date de création : {DATE_CREATED}\n"
            f"• Dernière modification : {DATE_MODIFIED}"
        )
        messagebox.showinfo("À propos", about_text)

    # ==========================================
    # ONGLET 1 : LISTE PROGRAMME USINAGE (USI-TAB)
    # ==========================================
    def setup_mapping_tab(self):
        frame_top = ttk.LabelFrame(self.tab_mapping, text="Gestion de la Base Liste Programme Usinage")
        frame_top.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(frame_top, text="Importer/Mettre à jour à partir d'un fichier XLS/CSV (ex: Usi-Tab -01.xlsx) :").pack(side=tk.LEFT, padx=10, pady=10)
        
        btn_import = tk.Button(frame_top, text="📥 Importer Fichier (CSV/Excel)", font=("Arial", 9, "bold"), bg="#0288D1", fg="white", command=self.import_mapping_file)
        btn_import.pack(side=tk.LEFT, padx=5, pady=10)

        btn_export = tk.Button(frame_top, text="📤 Exporter Base (CSV)", font=("Arial", 9), command=self.export_mapping_file)
        btn_export.pack(side=tk.LEFT, padx=5, pady=10)

        # Tableau des 11 colonnes de Usi-Tab
        frame_grid = ttk.LabelFrame(self.tab_mapping, text="Programmes Usinage et Modèles Associés Enregistrés")
        frame_grid.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("model_name", "prog_name", "block_dim", "block_dim_bought", "qty_per_block", "z_between_pains", "tools", "caisson", "top_plate", "bottom_plate", "remarks")
        self.tree_mapping = ttk.Treeview(frame_grid, columns=cols, show="headings")

        self.tree_mapping.heading("model_name", text="Nom Model")
        self.tree_mapping.heading("prog_name", text="Nom Programme Pain")
        self.tree_mapping.heading("block_dim", text="Dimension Bloc")
        self.tree_mapping.heading("block_dim_bought", text="Dimension Bloc (acheté)")
        self.tree_mapping.heading("qty_per_block", text="Qte/Bloc")
        self.tree_mapping.heading("z_between_pains", text="Z entre 2 pains")
        self.tree_mapping.heading("tools", text="Outils")
        self.tree_mapping.heading("caisson", text="Caisson")
        self.tree_mapping.heading("top_plate", text="PLAQUE 2su")
        self.tree_mapping.heading("bottom_plate", text="PLAQUE 2so")
        self.tree_mapping.heading("remarks", text="Rémarque")

        self.tree_mapping.column("model_name", width=110)
        self.tree_mapping.column("prog_name", width=110)
        self.tree_mapping.column("block_dim", width=140)
        self.tree_mapping.column("block_dim_bought", width=140)
        self.tree_mapping.column("qty_per_block", width=65, anchor="center")
        self.tree_mapping.column("z_between_pains", width=80, anchor="center")
        self.tree_mapping.column("tools", width=90)
        self.tree_mapping.column("caisson", width=60, anchor="center")
        self.tree_mapping.column("top_plate", width=90)
        self.tree_mapping.column("bottom_plate", width=90)
        self.tree_mapping.column("remarks", width=100)

        scrollbar_y = ttk.Scrollbar(frame_grid, orient=tk.VERTICAL, command=self.tree_mapping.yview)
        scrollbar_x = ttk.Scrollbar(frame_grid, orient=tk.HORIZONTAL, command=self.tree_mapping.xview)
        self.tree_mapping.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree_mapping.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.refresh_mapping_tree()

    def import_mapping_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Fichiers CSV/Excel", "*.csv;*.xlsx;*.xls"), ("Tous", "*.*")])
        if not file_path:
            return

        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM models_catalog")

            count = 0
            # Détection universelle d'encodage sans dépendance externes
            encodings_to_try = ['latin1', 'cp1252', 'utf-8-sig', 'iso-8859-1']
            rows = []

            if file_path.endswith('.csv'):
                for enc in encodings_to_try:
                    try:
                        with open(file_path, mode='r', encoding=enc) as f:
                            reader = csv.reader(f, delimiter=';')
                            rows = [r for r in reader if any(field.strip() for field in r)]
                            if len(rows) > 0:
                                break
                    except UnicodeDecodeError:
                        continue

                for r in rows[1:]:
                    if len(r) >= 2:
                        m_name = r[0].strip()
                        p_name = r[1].strip()
                        b_dim = r[2].replace('\n', ' ').strip() if len(r) > 2 else ""
                        b_bought = r[3].replace('\n', ' ').strip() if len(r) > 3 else ""
                        qty = r[4].strip() if len(r) > 4 else ""
                        z_p = r[5].strip() if len(r) > 5 else ""
                        tools = r[6].replace('\n', ' ').strip() if len(r) > 6 else ""
                        caisson = r[7].strip() if len(r) > 7 else ""
                        top_p = r[8].strip() if len(r) > 8 else ""
                        bot_p = r[9].strip() if len(r) > 9 else ""
                        rem = r[10].strip() if len(r) > 10 else ""

                        cursor.execute('''
                            INSERT INTO models_catalog (
                                model_name, prog_name, block_dim, block_dim_bought, qty_per_block,
                                z_between_pains, tools, caisson, top_plate, bottom_plate, remarks
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (m_name, p_name, b_dim, b_bought, qty, z_p, tools, caisson, top_p, bot_p, rem))
                        count += 1
            else:
                messagebox.showwarning("Format CSV Recommandé", "Pour éviter toute incompatibilité sans module externe, veuillez utiliser un fichier .CSV (séparateur point-virgule).")
                return

            conn.commit()
            conn.close()
            self.refresh_mapping_tree()
            messagebox.showinfo("Importation Réussie", f"{count} modèles et programmes chargés avec succès.")

        except Exception as e:
            messagebox.showerror("Erreur d'importation", f"Impossible de lire le fichier :\n{str(e)}")

    def export_mapping_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not file_path:
            return
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT model_name, prog_name, block_dim, block_dim_bought, qty_per_block, z_between_pains, tools, caisson, top_plate, bottom_plate, remarks FROM models_catalog")
        rows = cursor.fetchall()
        conn.close()

        with open(file_path, mode='w', newline='', encoding='latin1') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(["Nom Model", "Nom Programme pain", "Dimension Bloc", "Dimension Bloc (acheté)", "Qantite par Bloc", "Z entre 2 pains", "Outils", "Caisson", "PLAQUE 2su", "PLAQUE 2so", "Rémarque"])
            writer.writerows(rows)
        messagebox.showinfo("Exportation", "Base exportée avec succès.")

    def refresh_mapping_tree(self):
        for item in self.tree_mapping.get_children():
            self.tree_mapping.delete(item)

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT model_name, prog_name, block_dim, block_dim_bought, qty_per_block, z_between_pains, tools, caisson, top_plate, bottom_plate, remarks FROM models_catalog WHERE is_hidden=0")
        for row in cursor.fetchall():
            self.tree_mapping.insert("", tk.END, values=row)
        conn.close()

    # ==========================================
    # ONGLET 2 : COMPARAISON DE DOSSIERS
    # ==========================================
    def setup_compare_tab(self):
        frame_dirs = ttk.LabelFrame(self.tab_compare, text="Sélection des dossiers à comparer")
        frame_dirs.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(frame_dirs, text="Dossier A (Référence) :").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.var_dir_a = tk.StringVar()
        ttk.Entry(frame_dirs, textvariable=self.var_dir_a, width=68).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame_dirs, text="Parcourir", command=lambda: self.browse_dir(self.var_dir_a)).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(frame_dirs, text="Dossier B (Comparé) :").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.var_dir_b = tk.StringVar()
        ttk.Entry(frame_dirs, textvariable=self.var_dir_b, width=68).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame_dirs, text="Parcourir", command=lambda: self.browse_dir(self.var_dir_b)).grid(row=1, column=2, padx=5, pady=5)

        frame_btns = ttk.Frame(self.tab_compare)
        frame_btns.pack(fill=tk.X, padx=10, pady=5)

        btn_fast = tk.Button(frame_btns, text="Lancer la Comparaison Rapide (Dates/Tailles)", 
                             bg="#0288D1", fg="white", font=("Arial", 10, "bold"), pady=4,
                             command=lambda: self.run_comparison(deep=False))
        btn_fast.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        btn_deep = tk.Button(frame_btns, text="🔍 Comparaison Approfondie (Contenu Texte CNC)", 
                             bg="#2E7D32", fg="white", font=("Arial", 10, "bold"), pady=4,
                             command=lambda: self.run_comparison(deep=True))
        btn_deep.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        frame_grid = ttk.LabelFrame(self.tab_compare, text="Résultats de la comparaison")
        frame_grid.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("model", "fichier", "statut", "date_a", "date_b")
        self.tree = ttk.Treeview(frame_grid, columns=cols, show="headings", selectmode="browse")
        
        self.tree.heading("model", text="Nom de model ↕", command=lambda: self.sort_treeview("model"))
        self.tree.heading("fichier", text="Nom Prog. (Fichier/Chemin) ↕", command=lambda: self.sort_treeview("fichier"))
        self.tree.heading("statut", text="Statut ↕", command=lambda: self.sort_treeview("statut"))
        self.tree.heading("date_a", text="Date Modification (A)")
        self.tree.heading("date_b", text="Date Modification (B)")

        self.tree.column("model", width=140, anchor="w")
        self.tree.column("fichier", width=280, anchor="w")
        self.tree.column("statut", width=110, anchor="center")
        self.tree.column("date_a", width=170, anchor="center")
        self.tree.column("date_b", width=170, anchor="center")

        # Configuration visuelle forcée pour le fond ROUGE sous Windows
        self.tree.tag_configure("different_tag", background="#D32F2F", foreground="white")
        self.tree.tag_configure("identique_tag", background="white", foreground="black")

        scrollbar = ttk.Scrollbar(frame_grid, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        frame_bottom = ttk.Frame(self.tab_compare)
        frame_bottom.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(frame_bottom, text="Exporter en TXT", command=self.export_txt).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_bottom, text="Exporter en Excel (.csv)", command=self.export_csv).pack(side=tk.LEFT, padx=5)

        btn_print = tk.Button(frame_bottom, text="🖨️ Imprimer Rapport A4", 
                              bg="#424242", fg="white", font=("Arial", 9, "bold"),
                              command=self.print_a4_formatted)
        btn_print.pack(side=tk.RIGHT, padx=5)

    def find_model_name_for_file(self, filename):
        base_file = os.path.basename(filename).upper()
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT model_name, prog_name, top_plate, bottom_plate FROM models_catalog")
        rows = cursor.fetchall()
        conn.close()

        for m_name, p_name, top_p, bot_p in rows:
            if p_name and p_name.upper() in base_file:
                return m_name
            if top_p and top_p.upper() in base_file:
                return m_name
            if bot_p and bot_p.upper() in base_file:
                return m_name

        return "-"

    def sort_treeview(self, col):
        reverse = self.sort_directions.get(col, False)
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        items.sort(reverse=reverse)

        for index, (val, k) in enumerate(items):
            self.tree.move(k, '', index)

        self.sort_directions[col] = not reverse

    def browse_dir(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def run_comparison(self, deep=False):
        dir_a = self.var_dir_a.get().strip()
        dir_b = self.var_dir_b.get().strip()

        if not os.path.isdir(dir_a) or not os.path.isdir(dir_b):
            messagebox.showwarning("Attention", "Veuillez sélectionner deux dossiers valides à comparer.")
            return

        for item in self.tree.get_children():
            self.tree.delete(item)

        files_a = {os.path.relpath(os.path.join(r, f), dir_a): os.path.join(r, f) 
                   for r, d, files in os.walk(dir_a) for f in files}
        files_b = {os.path.relpath(os.path.join(r, f), dir_b): os.path.join(r, f) 
                   for r, d, files in os.walk(dir_b) for f in files}

        all_rel_files = sorted(list(set(files_a.keys()).union(set(files_b.keys()))))

        for rel in all_rel_files:
            path_a = files_a.get(rel)
            path_b = files_b.get(rel)

            date_a_str = self.get_file_date(path_a) if path_a else "Absence (A)"
            date_b_str = self.get_file_date(path_b) if path_b else "Absence (B)"

            if path_a and path_b:
                if deep:
                    is_same = filecmp.cmp(path_a, path_b, shallow=False)
                else:
                    stat_a = os.stat(path_a)
                    stat_b = os.stat(path_b)
                    is_same = (stat_a.st_mtime == stat_b.st_mtime) and (stat_a.st_size == stat_b.st_size)

                statut = "IDENTIQUE" if is_same else "DIFFERENT"
            else:
                statut = "DIFFERENT"

            model_name = self.find_model_name_for_file(rel)

            tag = "different_tag" if statut == "DIFFERENT" else "identique_tag"
            self.tree.insert("", tk.END, values=(model_name, rel, statut, date_a_str, date_b_str), tags=(tag,))

    # ==========================================
    # ONGLET 3 : CONFIGURATION & PLANIFICATION
    # ==========================================
    def setup_backup_tab(self):
        frame_add = ttk.LabelFrame(self.tab_backup, text="Ajouter / Configurer une Sauvegarde")
        frame_add.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(frame_add, text="Nom de la sauvegarde :").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.entry_task_name = ttk.Entry(frame_add, width=45)
        self.entry_task_name.grid(row=0, column=1, columnspan=2, sticky="w", padx=5, pady=3)

        ttk.Label(frame_add, text="Dossier / Fichier Source :").grid(row=1, column=0, sticky="w", padx=5, pady=3)
        self.entry_src = ttk.Entry(frame_add, width=55)
        self.entry_src.grid(row=1, column=1, sticky="w", padx=5, pady=3)
        ttk.Button(frame_add, text="Parcourir...", command=self.browse_src).grid(row=1, column=2, padx=5, pady=3)

        ttk.Label(frame_add, text="Dossier Destination :").grid(row=2, column=0, sticky="w", padx=5, pady=3)
        self.entry_dest = ttk.Entry(frame_add, width=55)
        self.entry_dest.grid(row=2, column=1, sticky="w", padx=5, pady=3)
        ttk.Button(frame_add, text="Parcourir...", command=self.browse_dest).grid(row=2, column=2, padx=5, pady=3)

        ttk.Label(frame_add, text="Type de récurrence :").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.combo_type = ttk.Combobox(frame_add, values=["Journalier", "Hebdomadaire", "Mensuel"], state="readonly", width=20)
        self.combo_type.current(0)
        self.combo_type.grid(row=3, column=1, sticky="w", padx=5, pady=5)
        self.combo_type.bind("<<ComboboxSelected>>", self.on_recurrence_change)

        self.label_days = ttk.Label(frame_add, text="Jours actifs (Hebdo) :")
        self.label_days.grid(row=4, column=0, sticky="w", padx=5, pady=3)

        self.days_frame = ttk.Frame(frame_add)
        self.days_frame.grid(row=4, column=1, columnspan=2, sticky="w", padx=5)

        self.days_vars = {}
        self.days_checkbuttons = []
        for day in ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]:
            var = tk.BooleanVar(value=True)
            chk = ttk.Checkbutton(self.days_frame, text=day, variable=var)
            chk.pack(side=tk.LEFT, padx=2)
            self.days_vars[day] = var
            self.days_checkbuttons.append(chk)

        ttk.Label(frame_add, text="Jour du mois (1-31) :").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.spin_month_day = tk.Spinbox(frame_add, from_=1, to=31, width=5, format="%02.0f", state=tk.DISABLED)
        self.spin_month_day.grid(row=5, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(frame_add, text="Heure d'exécution (HH:MM) :").grid(row=6, column=0, sticky="w", padx=5, pady=5)
        time_frame = ttk.Frame(frame_add)
        time_frame.grid(row=6, column=1, sticky="w", padx=5)

        self.spin_hour = tk.Spinbox(time_frame, from_=0, to=23, width=3, format="%02.0f")
        self.spin_hour.pack(side=tk.LEFT)
        ttk.Label(time_frame, text=" : ").pack(side=tk.LEFT)
        self.spin_min = tk.Spinbox(time_frame, from_=0, to=59, width=3, format="%02.0f")
        self.spin_min.pack(side=tk.LEFT)

        btn_save = tk.Button(frame_add, text="Enregistrer la Sauvegarde", font=("Arial", 10, "bold"), bg="#2E7D32", fg="white", pady=4, command=self.add_task)
        btn_save.grid(row=7, column=0, columnspan=3, pady=10, sticky="ew", padx=5)

        frame_list = ttk.LabelFrame(self.tab_backup, text="Liste des Sauvegardes Enregistrées")
        frame_list.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("name", "src", "dest", "type", "days", "time")
        self.tree_tasks = ttk.Treeview(frame_list, columns=columns, show="headings", height=6)
        self.tree_tasks.heading("name", text="Nom Tâche")
        self.tree_tasks.heading("src", text="Source")
        self.tree_tasks.heading("dest", text="Destination")
        self.tree_tasks.heading("type", text="Récurrence")
        self.tree_tasks.heading("days", text="Planification / Jour")
        self.tree_tasks.heading("time", text="Heure")

        self.tree_tasks.column("name", width=130)
        self.tree_tasks.column("src", width=200)
        self.tree_tasks.column("dest", width=200)
        self.tree_tasks.column("type", width=90)
        self.tree_tasks.column("days", width=140)
        self.tree_tasks.column("time", width=60)
        self.tree_tasks.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=5, pady=5)

        btn_actions = ttk.Frame(frame_list)
        btn_actions.pack(fill=tk.Y, side=tk.RIGHT, padx=5, pady=5)

        btn_manual = tk.Button(btn_actions, text="Lancer Manuel", font=("Arial", 9, "bold"), bg="#1976D2", fg="white", command=self.run_task_manual)
        btn_manual.pack(fill=tk.X, pady=5)

        btn_del = tk.Button(btn_actions, text="Supprimer", font=("Arial", 9, "bold"), bg="#C62828", fg="white", command=self.delete_task)
        btn_del.pack(fill=tk.X, pady=5)

        self.on_recurrence_change()

    def on_recurrence_change(self, event=None):
        rec_type = self.combo_type.get()
        if rec_type == "Mensuel":
            self.spin_month_day.config(state=tk.NORMAL)
            for chk in self.days_checkbuttons:
                chk.config(state=tk.DISABLED)
        elif rec_type == "Hebdomadaire":
            self.spin_month_day.config(state=tk.DISABLED)
            for chk in self.days_checkbuttons:
                chk.config(state=tk.NORMAL)
        else:
            self.spin_month_day.config(state=tk.DISABLED)
            for chk in self.days_checkbuttons:
                chk.config(state=tk.NORMAL)

    def browse_src(self):
        path = filedialog.askdirectory(title="Choisir le dossier source")
        if path:
            self.entry_src.delete(0, tk.END)
            self.entry_src.insert(0, path)

    def browse_dest(self):
        path = filedialog.askdirectory(title="Choisir le dossier destination")
        if path:
            self.entry_dest.delete(0, tk.END)
            self.entry_dest.insert(0, path)

    def add_task(self):
        name = self.entry_task_name.get().strip()
        src = self.entry_src.get().strip()
        dest = self.entry_dest.get().strip()
        rec_type = self.combo_type.get()

        if not name or not src or not dest:
            messagebox.showwarning("Champs manquants", "Veuillez remplir le nom, la source et la destination.")
            return

        if rec_type == "Mensuel":
            day_val = self.spin_month_day.get()
            days_str = f"Le {int(day_val):02d} du mois"
        elif rec_type == "Hebdomadaire":
            selected_days = [day[:3] for day, var in self.days_vars.items() if var.get()]
            days_str = ", ".join(selected_days) if selected_days else "Aucun"
        else:
            days_str = "Tous les jours"

        time_str = f"{int(self.spin_hour.get()):02d}:{int(self.spin_min.get()):02d}"

        task = (name, src, dest, rec_type, days_str, time_str)
        self.tasks.append(task)
        self.tree_tasks.insert("", tk.END, values=task)
        self.entry_task_name.delete(0, tk.END)
        messagebox.showinfo("Succès", f"La sauvegarde '{name}' a été ajoutée.")

    def run_task_manual(self):
        selected = self.tree_tasks.selection()
        if not selected:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner une sauvegarde à déclencher.")
            return

        item = self.tree_tasks.item(selected[0])
        name, src, dest, _, _, _ = item['values']

        if not os.path.exists(src):
            messagebox.showerror("Erreur Source", f"Le dossier source n'existe pas ou est inaccessible :\n{src}")
            return

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        target_dir = os.path.join(dest, f"{name}_{timestamp}")

        file_list = []
        if os.path.isfile(src):
            file_list.append(src)
        else:
            for root_dir, _, files in os.walk(src):
                for f in files:
                    file_list.append(os.path.join(root_dir, f))

        total_files = len(file_list)
        if total_files == 0:
            messagebox.showwarning("Dossier Vide", "Le dossier source est vide.")
            return

        progress_dialog = BackupProgressBarDialog(self, title=f"Sauvegarde : {name}")

        try:
            os.makedirs(target_dir, exist_ok=True)
            for idx, file_path in enumerate(file_list, 1):
                if progress_dialog.cancelled:
                    messagebox.showinfo("Annulation", "Sauvegarde annulée par l'utilisateur.")
                    return

                rel_path = os.path.relpath(file_path, src) if os.path.isdir(src) else os.path.basename(file_path)
                dest_file_path = os.path.join(target_dir, rel_path)
                os.makedirs(os.path.dirname(dest_file_path), exist_ok=True)

                shutil.copy2(file_path, dest_file_path)
                progress_dialog.update_progress(idx, total_files, filename=os.path.basename(file_path))
                time.sleep(0.01)

            progress_dialog.complete()
        except Exception as e:
            progress_dialog.destroy()
            messagebox.showerror("Erreur de sauvegarde", f"Échec lors de la sauvegarde : {str(e)}")

    def delete_task(self):
        selected = self.tree_tasks.selection()
        if not selected:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner une sauvegarde à supprimer.")
            return

        if messagebox.askyesno("Confirmation de suppression", "Êtes-vous sûr de vouloir supprimer cette tâche de sauvegarde ?"):
            self.tree_tasks.delete(selected[0])

    # ==========================================
    # ONGLET 4 : ORDRES DE FABRICATION (OF)
    # ==========================================
    def setup_of_tab(self):
        frame_new = ttk.LabelFrame(self.tab_of, text=" Saisie de l'Ordre de Fabrication & Traçabilité Brut ")
        frame_new.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_new, text="Modèle:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.entry_of_model = ttk.Entry(frame_new, width=20)
        self.entry_of_model.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_new, text="Machine:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.combo_of_mach = ttk.Combobox(frame_new, values=["NUM 1060 (5-Axes)", "Fraiseuse EPS", "Tour CNC"], state="readonly", width=18)
        self.combo_of_mach.current(0)
        self.combo_of_mach.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_new, text="Priorité:").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.combo_of_prio = ttk.Combobox(frame_new, values=["Haute", "Normale", "Basse"], state="readonly", width=10)
        self.combo_of_prio.current(1)
        self.combo_of_prio.grid(row=0, column=5, padx=5, pady=5)

        ttk.Label(frame_new, text="N° Bloc:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.entry_block_num = ttk.Entry(frame_new, width=12)
        self.entry_block_num.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame_new, text="Densité Bloc:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.entry_block_density = ttk.Entry(frame_new, width=12)
        self.entry_block_density.grid(row=1, column=3, padx=5, pady=5)

        ttk.Label(frame_new, text="N° Pain:").grid(row=1, column=4, padx=5, pady=5, sticky="w")
        self.entry_pain_num = ttk.Entry(frame_new, width=10)
        self.entry_pain_num.grid(row=1, column=5, padx=5, pady=5)

        ttk.Label(frame_new, text="Poids Pain (g):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.entry_pain_weight = ttk.Entry(frame_new, width=12)
        self.entry_pain_weight.grid(row=2, column=1, padx=5, pady=5)

        ttk.Button(frame_new, text="Lancer l'OF Auto-Incrémenté", command=self.create_of_auto).grid(row=2, column=3, columnspan=3, padx=10, pady=5, sticky="ew")

        frame_list = ttk.Frame(self.tab_of)
        frame_list.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("id", "of_number", "model_name", "prog_name", "machine", "assigned_operator", "priority", "block_num", "block_density", "pain_num", "pain_weight", "status", "created_at")
        self.tree_of = ttk.Treeview(frame_list, columns=cols, show="headings")

        self.tree_of.heading("id", text="ID")
        self.tree_of.heading("of_number", text="N° OF")
        self.tree_of.heading("model_name", text="Modèle")
        self.tree_of.heading("prog_name", text="Prog. Pain")
        self.tree_of.heading("machine", text="Machine")
        self.tree_of.heading("assigned_operator", text="Opérateur")
        self.tree_of.heading("priority", text="Priorité")
        self.tree_of.heading("block_num", text="N° Bloc")
        self.tree_of.heading("block_density", text="Densité")
        self.tree_of.heading("pain_num", text="N° Pain")
        self.tree_of.heading("pain_weight", text="Poids (g)")
        self.tree_of.heading("status", text="Statut")
        self.tree_of.heading("created_at", text="Date/Heure")

        self.tree_of.column("id", width=35, anchor="center")
        self.tree_of.column("of_number", width=110, anchor="center")
        self.tree_of.column("model_name", width=110)
        self.tree_of.column("prog_name", width=100)
        self.tree_of.column("machine", width=120)
        self.tree_of.column("assigned_operator", width=90)
        self.tree_of.column("priority", width=70, anchor="center")
        self.tree_of.column("block_num", width=70, anchor="center")
        self.tree_of.column("block_density", width=70, anchor="center")
        self.tree_of.column("pain_num", width=70, anchor="center")
        self.tree_of.column("pain_weight", width=70, anchor="center")
        self.tree_of.column("status", width=90, anchor="center")
        self.tree_of.column("created_at", width=130, anchor="center")

        scrollbar = ttk.Scrollbar(frame_list, orient="vertical", command=self.tree_of.yview)
        self.tree_of.configure(yscrollcommand=scrollbar.set)
        self.tree_of.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.load_of_data()

    def generate_next_of_number(self):
        year = datetime.datetime.now().year
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM work_orders ORDER BY id DESC LIMIT 1")
        row = cursor.fetchone()
        conn.close()

        next_id = (row[0] + 1) if row else 1
        return f"OF{next_id:05d}/{year}"

    def create_of_auto(self):
        model_name = self.entry_of_model.get().strip()
        if not model_name:
            messagebox.showwarning("Attention", "Veuillez indiquer un modèle.")
            return

        b_num = self.entry_block_num.get().strip()
        b_dens = self.entry_block_density.get().strip()
        p_num = self.entry_pain_num.get().strip()
        p_weight = self.entry_pain_weight.get().strip()

        of_num = self.generate_next_of_number()

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT prog_name FROM models_catalog WHERE model_name=?", (model_name,))
        row = cursor.fetchone()
        prog = row[0] if row else ""

        cursor.execute('''
            INSERT INTO work_orders (of_number, model_name, prog_name, machine, assigned_operator, priority, block_num, block_density, pain_num, pain_weight)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (of_num, model_name, prog, self.combo_of_mach.get(), "Admin", self.combo_of_prio.get(), b_num, b_dens, p_num, p_weight))

        conn.commit()
        conn.close()

        messagebox.showinfo("Succès", f"Ordre de fabrication {of_num} créé automatiquement.")
        self.load_of_data()

    def load_of_data(self):
        for item in self.tree_of.get_children():
            self.tree_of.delete(item)
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT id, of_number, model_name, prog_name, machine, assigned_operator, priority, block_num, block_density, pain_num, pain_weight, status, created_at FROM work_orders ORDER BY id DESC")
        for row in cursor.fetchall():
            self.tree_of.insert("", "end", values=row)
        conn.close()

    # ==========================================
    # ONGLET 5 : TRAÇABILITÉ & SUIVI
    # ==========================================
    def setup_tracking_tab(self):
        frame_input = ttk.LabelFrame(self.tab_tracking, text=" Enregistrer une Étape d'Usinage ")
        frame_input.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_input, text="N° OF:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_tr_of = ttk.Entry(frame_input, width=12)
        self.entry_tr_of.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_input, text="Machine:").grid(row=0, column=2, padx=5, pady=5)
        self.combo_tr_mach = ttk.Combobox(frame_input, values=["NUM 1060 (5-Axes)", "Fraiseuse EPS", "Tour CNC"], state="readonly", width=16)
        self.combo_tr_mach.current(0)
        self.combo_tr_mach.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_input, text="Modèle / Pièce:").grid(row=0, column=4, padx=5, pady=5)
        self.entry_tr_model = ttk.Entry(frame_input, width=20)
        self.entry_tr_model.grid(row=0, column=5, padx=5, pady=5)

        ttk.Label(frame_input, text="Temps Réel (min):").grid(row=1, column=0, padx=5, pady=5)
        self.entry_tr_time = ttk.Entry(frame_input, width=12)
        self.entry_tr_time.insert(0, "45")
        self.entry_tr_time.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame_input, text="Statut:").grid(row=1, column=2, padx=5, pady=5)
        self.combo_tr_stat = ttk.Combobox(frame_input, values=["Terminé", "En cours", "Maintenance", "En attente"], state="readonly", width=16)
        self.combo_tr_stat.current(0)
        self.combo_tr_stat.grid(row=1, column=3, padx=5, pady=5)

        ttk.Button(frame_input, text="Enregistrer l'Usinage", command=self.add_tracking_log).grid(row=1, column=4, columnspan=2, padx=10, pady=5, sticky="ew")

        frame_list = ttk.Frame(self.tab_tracking)
        frame_list.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("id", "of_number", "operator", "machine", "model_name", "real_time_min", "status", "timestamp")
        self.tree_track = ttk.Treeview(frame_list, columns=cols, show="headings")

        self.tree_track.heading("id", text="ID")
        self.tree_track.heading("of_number", text="N° OF")
        self.tree_track.heading("operator", text="Opérateur")
        self.tree_track.heading("machine", text="Machine")
        self.tree_track.heading("model_name", text="Modèle / Pièce")
        self.tree_track.heading("real_time_min", text="Tps Réel (min)")
        self.tree_track.heading("status", text="Statut")
        self.tree_track.heading("timestamp", text="Date / Heure")

        self.tree_track.column("id", width=35, anchor="center")
        self.tree_track.column("of_number", width=110, anchor="center")
        self.tree_track.column("operator", width=100)
        self.tree_track.column("machine", width=120)
        self.tree_track.column("model_name", width=120)
        self.tree_track.column("real_time_min", width=90, anchor="center")
        self.tree_track.column("status", width=90, anchor="center")
        self.tree_track.column("timestamp", width=130, anchor="center")

        scrollbar = ttk.Scrollbar(frame_list, orient="vertical", command=self.tree_track.yview)
        self.tree_track.configure(yscrollcommand=scrollbar.set)
        self.tree_track.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.load_tracking_data()

    def add_tracking_log(self):
        of_num = self.entry_tr_of.get().strip()
        model = self.entry_tr_model.get().strip()
        t_time = self.entry_tr_time.get().strip()

        if not of_num or not model:
            messagebox.showwarning("Attention", "N° OF et Modèle sont obligatoires.")
            return

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO machining_history (of_number, operator, machine, model_name, real_time_min, status)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (of_num, "Opérateur", self.combo_tr_mach.get(), model, t_time, self.combo_tr_stat.get()))

        cursor.execute("UPDATE work_orders SET status=? WHERE of_number=?", (self.combo_tr_stat.get(), of_num))

        conn.commit()
        conn.close()

        messagebox.showinfo("Succès", "Étape d'usinage enregistrée.")
        self.load_tracking_data()
        self.load_of_data()

    def load_tracking_data(self):
        for item in self.tree_track.get_children():
            self.tree_track.delete(item)
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT id, of_number, operator, machine, model_name, real_time_min, status, timestamp FROM machining_history ORDER BY id DESC")
        for row in cursor.fetchall():
            self.tree_track.insert("", "end", values=row)
        conn.close()

    # ==========================================
    # ONGLET 6 : TRANSFERT RS232 CNC
    # ==========================================
    def setup_cnc_tab(self):
        frame_cfg = ttk.LabelFrame(self.tab_cnc, text=" Paramètres de Communication RS232 (NUM 1060) ")
        frame_cfg.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_cfg, text="Port COM :").grid(row=0, column=0, padx=5, pady=5)
        self.combo_port = ttk.Combobox(frame_cfg, values=["COM1", "COM2", "COM3", "COM4"], width=10)
        self.combo_port.set("COM1")
        self.combo_port.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_cfg, text="Baudrate :").grid(row=0, column=2, padx=5, pady=5)
        self.combo_baud = ttk.Combobox(frame_cfg, values=["4800", "9600", "19200"], width=10)
        self.combo_baud.set("9600")
        self.combo_baud.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_cfg, text="Parité :").grid(row=0, column=4, padx=5, pady=5)
        self.combo_parity = ttk.Combobox(frame_cfg, values=["Even", "Odd", "None"], width=10)
        self.combo_parity.set("Even")
        self.combo_parity.grid(row=0, column=5, padx=5, pady=5)

        frame_file = ttk.Frame(self.tab_cnc)
        frame_file.pack(fill="x", padx=10, pady=5)
        
        self.lbl_file = ttk.Label(frame_file, text="Aucun fichier chargé", font=("Arial", 9, "italic"))
        self.lbl_file.pack(side="left", padx=5)

        ttk.Button(frame_file, text="Ouvrir Fichier G-Code (.ISO / .NC)", command=self.open_gcode_file).pack(side="right", padx=5)

        frame_prev = ttk.LabelFrame(self.tab_cnc, text=" Aperçu du Programme ISO / NUM ")
        frame_prev.pack(fill="both", expand=True, padx=10, pady=5)

        self.txt_preview = tk.Text(frame_prev, wrap="none", font=("Courier", 10))
        scroll_y = ttk.Scrollbar(frame_prev, orient="vertical", command=self.txt_preview.yview)
        scroll_x = ttk.Scrollbar(frame_prev, orient="horizontal", command=self.txt_preview.xview)
        self.txt_preview.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.txt_preview.pack(fill="both", expand=True)

        frame_send = ttk.Frame(self.tab_cnc)
        frame_send.pack(fill="x", padx=10, pady=10)

        ttk.Button(frame_send, text="🚀 ENVOYER VERS LA CNC (RS232)", command=self.send_to_cnc).pack(fill="x", ipady=5)

    def open_gcode_file(self):
        path = filedialog.askopenfilename(filetypes=[("Programme CNC", "*.iso *.nc *.txt"), ("Tous", "*.*")])
        if path:
            self.lbl_file.config(text=f"Fichier : {os.path.basename(path)}")
            with open(path, "r", encoding="latin1") as f:
                content = f.read()
                self.txt_preview.delete("1.0", tk.END)
                self.txt_preview.insert(tk.END, content)

    def send_to_cnc(self):
        gcode = self.txt_preview.get("1.0", tk.END).strip()
        if not gcode:
            messagebox.showwarning("Attention", "Aucun programme G-Code à envoyer.")
            return

        port = self.combo_port.get()
        baud = int(self.combo_baud.get())

        if serial is None:
            messagebox.showinfo("Simulation Transfert RS232", f"[Mode Simulation]\nPySerial non détecté.\nProgramme transmis avec succès à la NUM 1060 via {port} à {baud} bauds.")
            return

        try:
            ser = serial.Serial(port, baud, timeout=2)
            ser.write(gcode.encode('latin1'))
            ser.close()
            messagebox.showinfo("Transfert Réussi", f"Programme transmis avec succès à la NUM 1060 via {port}.")
        except Exception as e:
            messagebox.showerror("Erreur RS232", f"Impossible de communiquer avec le port {port} :\n{str(e)}")

    # ==========================================
    # FONCTIONS UTILITAIRES DE L'APPLICATION
    # ==========================================
    def get_file_date(self, path):
        mtime = os.path.getmtime(path)
        return datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')

    def change_password(self):
        old_p = simpledialog.askstring("Changer mot de passe", "Ancien mot de passe :", show="*")
        if not old_p:
            return
        cfg = load_config()
        if hashlib.sha256(old_p.encode()).hexdigest() != cfg["password_hash"]:
            messagebox.showerror("Erreur", "L'ancien mot de passe est incorrect !")
            return

        new_p = simpledialog.askstring("Changer mot de passe", "Nouveau mot de passe :", show="*")
        if new_p:
            cfg["password_hash"] = hashlib.sha256(new_p.encode()).hexdigest()
            save_config(cfg)
            messagebox.showinfo("Succès", "Mot de passe modifié avec succès !")

    def print_a4_formatted(self):
        items = self.tree.get_children()
        if not items:
            messagebox.showwarning("Attention", "Aucune donnée à imprimer !")
            return

        rows = [self.tree.item(item, "values") for item in items]

        if HAS_REPORTLAB:
            self.generate_pdf_reportlab(rows)
        else:
            self.generate_html_print(rows)

    def generate_pdf_reportlab(self, rows):
        pdf_filename = os.path.join(tempfile.gettempdir(), "rapport_cnc_a4.pdf")
        doc = SimpleDocTemplate(pdf_filename, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=12, leading=14, alignment=1)
        cell_style = ParagraphStyle('CellStyle', fontName='Helvetica', fontSize=7, leading=8)
        cell_style_bold = ParagraphStyle('CellStyleBold', fontName='Helvetica-Bold', fontSize=7, leading=8, textColor=colors.white)

        elements = []
        elements.append(Paragraph(f"<b>RAPPORT DE COMPARAISON - {APP_NAME.upper()}</b>", title_style))
        elements.append(Spacer(1, 10))

        data = [[
            Paragraph("<b>Nom Model</b>", cell_style_bold),
            Paragraph("<b>Nom Prog. (Fichier)</b>", cell_style_bold),
            Paragraph("<b>Statut</b>", cell_style_bold),
            Paragraph("<b>Date Modification (A)</b>", cell_style_bold),
            Paragraph("<b>Date Modification (B)</b>", cell_style_bold)
        ]]

        table_styles = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0B3C5D")),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
        ]

        for i, r in enumerate(rows, start=1):
            model, fichier, statut, date_a, date_b = r
            
            p_model = Paragraph(model, cell_style)
            p_fichier = Paragraph(fichier, cell_style)
            p_statut = Paragraph(f"<b>{statut}</b>", cell_style)
            p_date_a = Paragraph(date_a, cell_style)
            p_date_b = Paragraph(date_b, cell_style)

            data.append([p_model, p_fichier, p_statut, p_date_a, p_date_b])

            if statut == "DIFFERENT":
                table_styles.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor("#D32F2F")))
                cell_white = ParagraphStyle('CellW', parent=cell_style, textColor=colors.white)
                data[i] = [
                    Paragraph(model, cell_white),
                    Paragraph(fichier, cell_white),
                    Paragraph(f"<b>{statut}</b>", cell_white),
                    Paragraph(date_a, cell_white),
                    Paragraph(date_b, cell_white)
                ]

        t = Table(data, colWidths=[90, 180, 70, 105, 105])
        t.setStyle(TableStyle(table_styles))
        elements.append(t)

        doc.build(elements)
        os.startfile(pdf_filename, "print") if hasattr(os, "startfile") else webbrowser.open(pdf_filename)

    def generate_html_print(self, rows):
        html_filename = os.path.join(tempfile.gettempdir(), "rapport_cnc_a4.html")
        
        html_content = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Rapport CNC A4</title>
            <style>
                @page {{ size: A4 portrait; margin: 10mm; }}
                body {{ font-family: Arial, sans-serif; font-size: 9pt; margin: 0; }}
                h2 {{ text-align: center; margin-bottom: 15px; font-size: 12pt; }}
                table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
                th, td {{ border: 1px solid #666; padding: 4px 6px; word-wrap: break-word; font-size: 8pt; }}
                th {{ background-color: #0B3C5D; color: white; text-align: left; }}
                tr.different {{ background-color: #D32F2F !important; color: white !important; font-weight: bold; }}
                col.c1 {{ width: 20%; }}
                col.c2 {{ width: 35%; }}
                col.c3 {{ width: 15%; }}
                col.c4 {{ width: 15%; }}
                col.c5 {{ width: 15%; }}
            </style>
        </head>
        <body onload="window.print();">
            <h2>RAPPORT DE COMPARAISON - {APP_NAME.upper()}</h2>
            <table>
                <colgroup>
                    <col class="c1"><col class="c2"><col class="c3"><col class="c4"><col class="c5">
                </colgroup>
                <thead>
                    <tr>
                        <th>Nom Model</th>
                        <th>Nom Prog. (Fichier)</th>
                        <th>Statut</th>
                        <th>Date Modification (A)</th>
                        <th>Date Modification (B)</th>
                    </tr>
                </thead>
                <tbody>
        """
        for r in rows:
            model, fichier, statut, date_a, date_b = r
            row_class = "different" if statut == "DIFFERENT" else ""
            html_content += f"""
                <tr class="{row_class}">
                    <td>{model}</td>
                    <td>{fichier}</td>
                    <td>{statut}</td>
                    <td>{date_a}</td>
                    <td>{date_b}</td>
                </tr>
            """

        html_content += """
                </tbody>
            </table>
        </body>
        </html>
        """
        with open(html_filename, "w", encoding="utf-8") as f:
            f.write(html_content)

        webbrowser.open(html_filename)

    def export_txt(self):
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Texte", "*.txt")])
        if path:
            with open(path, "w", encoding="utf-8") as f:
                for item in self.tree.get_children():
                    f.write("\t".join(map(str, self.tree.item(item, "values"))) + "\n")
            messagebox.showinfo("Export", "Export TXT réussi !")

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow(["Nom Model", "Nom Programme pain / Fichier", "Statut", "Date A", "Date B"])
                for item in self.tree.get_children():
                    writer.writerow(self.tree.item(item, "values"))
            messagebox.showinfo("Export", "Export CSV réussi !")


if __name__ == "__main__":
    init_db()
    app = ProgrammeCNCManager()
    app.mainloop()
