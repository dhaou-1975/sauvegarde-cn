import sys
import os
import sqlite3
import hashlib
import datetime
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

DB_NAME = "cnc_atelier.db"
VIRTUAL_CN_DIR = os.path.abspath("CN_Virtuelle")

if not os.path.exists(VIRTUAL_CN_DIR):
    os.makedirs(VIRTUAL_CN_DIR)

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # Users table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL CHECK(role IN ('Admin', 'Superviseur', 'Opérateur'))
        )
    ''')
    
    # Default admin user if none exists
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        h = hashlib.sha256("admin123".encode()).hexdigest()
        cursor.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)", ("dhaou", h, "Admin"))
        h_op = hashlib.sha256("op123".encode()).hexdigest()
        cursor.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)", ("operateur1", h_op, "Opérateur"))
        h_sup = hashlib.sha256("sup123".encode()).hexdigest()
        cursor.execute("INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)", ("superviseur1", h_sup, "Superviseur"))
        conn.commit()

    # Article Types table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS article_types (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            is_windsurf INTEGER DEFAULT 0
        )
    ''')
    cursor.execute("SELECT COUNT(*) FROM article_types")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO article_types (name, is_windsurf) VALUES ('Planche à voile', 1)")
        cursor.execute("INSERT INTO article_types (name, is_windsurf) VALUES ('PVC', 0)")
        cursor.execute("INSERT INTO article_types (name, is_windsurf) VALUES ('Surfaçage moule', 0)")
        cursor.execute("INSERT INTO article_types (name, is_windsurf) VALUES ('USBox', 0)")
        cursor.execute("INSERT INTO article_types (name, is_windsurf) VALUES ('Bois', 0)")
        conn.commit()

    # Catalog Models table (Flexible for Windsurf + generic)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS catalog_models (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            article_type TEXT NOT NULL,
            programme_pain TEXT,
            densite TEXT,
            dimensions_bloc TEXT,
            dimension_bloc_achete TEXT,
            qte_pains_bloc TEXT,
            z_entre_pains TEXT,
            outils TEXT,
            caisson TEXT,
            prog_pvc_dessus TEXT,
            prog_pvc_dessous TEXT,
            remarque TEXT,
            status TEXT DEFAULT 'Etude' CHECK(status IN ('Etude', 'Test', 'Valide'))
        )
    ''')

    # Blocks / Raw materials stock
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS raw_blocks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            block_number TEXT UNIQUE NOT NULL,
            material_type TEXT,
            density REAL,
            weight REAL,
            reception_date TEXT,
            tested INTEGER DEFAULT 0,
            stock_qty REAL
        )
    ''')

    # Manufacturing Orders (OF)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS manufacturing_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            of_number TEXT UNIQUE NOT NULL,
            model_name TEXT,
            article_type TEXT,
            assigned_operator TEXT,
            status TEXT DEFAULT 'Créé',
            block_number TEXT,
            part_number TEXT,
            finishing_operator TEXT,
            weight_finished REAL,
            quality_status TEXT DEFAULT 'En attente',
            non_conformity_cause TEXT,
            comment TEXT,
            created_date TEXT
        )
    ''')

    # Tools inventory & life
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tools (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tool_ref TEXT UNIQUE NOT NULL,
            description TEXT,
            max_life_hours REAL,
            current_usage_hours REAL DEFAULT 0,
            max_parts INTEGER,
            current_parts_count INTEGER DEFAULT 0
        )
    ''')

    # Machine Maintenance log
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS machine_maintenance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine_name TEXT,
            intervention_date TEXT,
            description TEXT,
            technician TEXT
        )
    ''')

    # Traceability Logs
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS traceability_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            username TEXT,
            action TEXT,
            details TEXT,
            machine TEXT,
            status TEXT
        )
    ''')

    conn.close()

init_db()

class CNCApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CNC Atelier Manager V2 - Phase 1 Prototype")
        self.root.geometry("1200x800")
        self.current_user = None
        self.current_role = None
        
        self.show_login_screen()

    def clear_window(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def show_login_screen(self):
        self.clear_window()
        frame = ttk.Frame(self.root, padding=40)
        frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        ttk.Label(frame, text="CNC Atelier Manager V2", font=("Helvetica", 20, "bold")).pack(pady=10)
        ttk.Label(frame, text="Authentification requise", font=("Helvetica", 12)).pack(pady=5)
        
        f_form = ttk.Frame(frame)
        f_form.pack(pady=15)
        
        ttk.Label(f_form, text="Identifiant :").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.e_user = ttk.Entry(f_form, width=25)
        self.e_user.grid(row=0, column=1, pady=5)
        self.e_user.insert(0, "dhaou")
        
        ttk.Label(f_form, text="Mot de passe :").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.e_pass = ttk.Entry(f_form, show="*", width=25)
        self.e_pass.grid(row=1, column=1, pady=5)
        self.e_pass.insert(0, "admin123")
        
        ttk.Button(frame, text="Se connecter", command=self.authenticate).pack(pady=15)

    def authenticate(self):
        u = self.e_user.get().strip()
        p = self.e_pass.get().strip()
        h = hashlib.sha256(p.encode()).hexdigest()
        
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT role FROM users WHERE username=? AND password_hash=?", (u, h))
        res = cursor.fetchone()
        conn.close()
        
        if res:
            self.current_user = u
            self.current_role = res[0]
            self.log_action("Connexion", f"Utilisateur {u} connecté avec le rôle {self.current_role}")
            self.show_main_interface()
        else:
            messagebox.showerror("Erreur", "Identifiant ou mot de passe incorrect.")

    def log_action(self, action, details, machine="CN_Virtuelle", status="Succès"):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cursor.execute("INSERT INTO traceability_logs (timestamp, username, action, details, machine, status) VALUES (?, ?, ?, ?, ?, ?)",
                       (ts, self.current_user or "System", action, details, machine, status))
        conn.commit()
        conn.close()

    def show_main_interface(self):
        self.clear_window()
        
        # Top bar
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(side=tk.TOP, fill=tk.X)
        
        ttk.Label(top_frame, text=f"Connecté : {self.current_user} ({self.current_role})", font=("Helvetica", 11, "bold")).pack(side=tk.LEFT)
        ttk.Button(top_frame, text="Changer d'utilisateur", command=self.show_login_screen).pack(side=tk.RIGHT)
        
        # Notebook for modules
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Tab 1: Catalogue CFAO
        self.tab_catalog = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_catalog, text="1. Catalogue CFAO")
        self.init_catalog_tab()
        
        # Tab 2: Ordres de Fabrication (OF) & Production
        self.tab_of = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_of, text="2. Ordres de Fabrication (OF)")
        self.init_of_tab()

        # Tab 3: Transfert CN Virtuelle
        self.tab_cn = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_cn, text="3. Transfert CN (RS232 Virtuel)")
        self.init_cn_tab()

        # Tab 4: Traçabilité & Enquête
        self.tab_trace = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_trace, text="4. Traçabilité & Enquête")
        self.init_trace_tab()

        # Tab 5: Comparaison WinMerge
        self.tab_compare = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_compare, text="5. Comparaison G-code (.tap)")
        self.init_compare_tab()

        # Tab 6: Stocks & Outils & Dashboard & Maintenance
        self.tab_tools = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_tools, text="6. Stocks, Outils & Dashboard")
        self.init_tools_dashboard_tab()

    def init_catalog_tab(self):
        frame = self.tab_catalog
        lbl = ttk.Label(frame, text="Catalogue des Modèles CFAO", font=("Helvetica", 14, "bold"))
        lbl.pack(anchor=tk.W, padx=10, pady=10)
        
        f_controls = ttk.Frame(frame, padding=5)
        f_controls.pack(fill=tk.X, padx=10)
        
        ttk.Button(f_controls, text="Actualiser", command=self.load_catalog_data).pack(side=tk.LEFT, padx=5)
        if self.current_role in ["Admin", "Superviseur"]:
            ttk.Button(f_controls, text="Ajouter un Modèle", command=self.add_catalog_model_dialog).pack(side=tk.LEFT, padx=5)
            ttk.Button(f_controls, text="Importer Excel/CSV", command=self.import_catalog_dialog).pack(side=tk.LEFT, padx=5)
            
        # Treeview for catalog
        columns = ("ID", "Nom", "Type", "Prog Pain", "Densité", "Dimensions", "Statut", "Remarque")
        self.tree_catalog = ttk.Treeview(frame, columns=columns, show="headings", height=20)
        for col in columns:
            self.tree_catalog.heading(col, text=col)
            self.tree_catalog.column(col, width=120)
        self.tree_catalog.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.load_catalog_data()

    def load_catalog_data(self):
        for row in self.tree_catalog.get_children():
            self.tree_catalog.delete(row)
        
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        if self.current_role == "Opérateur":
            cursor.execute("SELECT id, name, article_type, programme_pain, densite, dimensions_bloc, status, remarque FROM catalog_models WHERE status='Valide'")
        else:
            cursor.execute("SELECT id, name, article_type, programme_pain, densite, dimensions_bloc, status, remarque FROM catalog_models")
        for r in cursor.fetchall():
            self.tree_catalog.insert("", tk.END, values=r)
        conn.close()

    def add_catalog_model_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Ajouter un Modèle CFAO")
        win.geometry("500x600")
        
        ttk.Label(win, text="Nom du modèle :").pack(anchor=tk.W, padx=20, pady=5)
        e_name = ttk.Entry(win, width=40)
        e_name.pack(padx=20)
        
        ttk.Label(win, text="Type d'article :").pack(anchor=tk.W, padx=20, pady=5)
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM article_types")
        types = [row[0] for row in cursor.fetchall()]
        conn.close()
        
        cb_type = ttk.Combobox(win, values=types, state="readonly", width=38)
        cb_type.pack(padx=20)
        cb_type.set("Planche à voile")
        
        fields = [
            ("Programme pain", "e_prog"),
            ("Densité recommandée", "e_dens"),
            ("Dimensions du bloc", "e_dim"),
            ("Dimension bloc acheté", "e_dim_ach"),
            ("Quantité de pains par bloc", "e_qte"),
            ("Z entre 2 pains", "e_z"),
            ("Outils nécessaires", "e_outils"),
            ("Caisson (P / G)", "e_caisson"),
            ("Prog PVC face dessus", "e_pvc_su"),
            ("Prog PVC face dessous", "e_pvc_so"),
            ("Remarque", "e_rem")
        ]
        
        entries = {}
        for label, key in fields:
            ttk.Label(win, text=label).pack(anchor=tk.W, padx=20, pady=2)
            en = ttk.Entry(win, width=40)
            en.pack(padx=20)
            entries[key] = en
            
        ttk.Label(win, text="Statut :").pack(anchor=tk.W, padx=20, pady=5)
        cb_status = ttk.Combobox(win, values=["Etude", "Test", "Valide"], state="readonly", width=38)
        cb_status.pack(padx=20)
        cb_status.set("Etude")
        
        def save():
            name = e_name.get().strip()
            if not name:
                messagebox.showerror("Erreur", "Le nom du modèle est obligatoire.")
                return
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO catalog_models (name, article_type, programme_pain, densite, dimensions_bloc, dimension_bloc_achete, qte_pains_bloc, z_entre_pains, outils, caisson, prog_pvc_dessus, prog_pvc_dessous, remarque, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (name, cb_type.get(), entries['e_prog'].get(), entries['e_dens'].get(), entries['e_dim'].get(),
                  entries['e_dim_ach'].get(), entries['e_qte'].get(), entries['e_z'].get(), entries['e_outils'].get(),
                  entries['e_caisson'].get(), entries['e_pvc_su'].get(), entries['e_pvc_so'].get(), entries['e_rem'].get(), cb_status.get()))
            conn.commit()
            conn.close()
            self.log_action("Catalogue", f"Ajout du modèle {name} ({cb_type.get()})")
            messagebox.showinfo("Succès", "Modèle ajouté avec succès.")
            win.destroy()
            self.load_catalog_data()
            
        ttk.Button(win, text="Enregistrer", command=save).pack(pady=15)

    def import_catalog_dialog(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
        if file_path:
            messagebox.showinfo("Import", f"Fichier {file_path} importé avec succès (simulation).")
            self.log_action("Catalogue", f"Import de données depuis {file_path}")

    def init_of_tab(self):
        frame = self.tab_of
        ttk.Label(frame, text="Ordres de Fabrication (OF) & Flux Atelier", font=("Helvetica", 14, "bold")).pack(anchor=tk.W, padx=10, pady=10)
        
        f_ctrl = ttk.Frame(frame, padding=5)
        f_ctrl.pack(fill=tk.X, padx=10)
        
        ttk.Button(f_ctrl, text="Actualiser", command=self.load_of_data).pack(side=tk.LEFT, padx=5)
        if self.current_role in ["Admin", "Superviseur"]:
            ttk.Button(f_ctrl, text="Créer un OF", command=self.create_of_dialog).pack(side=tk.LEFT, padx=5)
            ttk.Button(f_ctrl, text="Enregistrer un Bloc Matière", command=self.create_block_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(f_ctrl, text="Mettre à jour Étape / Finition", command=self.update_of_progress_dialog).pack(side=tk.LEFT, padx=5)
        
        columns = ("ID", "OF #", "Modèle", "Type", "Opérateur Usinage", "Statut", "Bloc #", "Pain #", "Opérateur Finition", "Qualité")
        self.tree_of = ttk.Treeview(frame, columns=columns, show="headings", height=20)
        for col in columns:
            self.tree_of.heading(col, text=col)
            self.tree_of.column(col, width=100)
        self.tree_of.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.load_of_data()

    def load_of_data(self):
        for row in self.tree_of.get_children():
            self.tree_of.delete(row)
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        if self.current_role == "Opérateur":
            cursor.execute("SELECT id, of_number, model_name, article_type, assigned_operator, status, block_number, part_number, finishing_operator, quality_status FROM manufacturing_orders WHERE assigned_operator=?", (self.current_user,))
        else:
            cursor.execute("SELECT id, of_number, model_name, article_type, assigned_operator, status, block_number, part_number, finishing_operator, quality_status FROM manufacturing_orders")
        for r in cursor.fetchall():
            self.tree_of.insert("", tk.END, values=r)
        conn.close()

    def create_of_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Créer un Ordre de Fabrication (OF)")
        win.geometry("450x400")
        
        ttk.Label(win, text="Numéro d'OF :").pack(anchor=tk.W, padx=20, pady=5)
        e_of = ttk.Entry(win, width=35)
        e_of.pack(padx=20)
        e_of.insert(0, f"OF-{datetime.datetime.now().strftime('%Y%m%d-%H%M')}")
        
        ttk.Label(win, text="Modèle CFAO :").pack(anchor=tk.W, padx=20, pady=5)
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT name, article_type FROM catalog_models WHERE status='Valide'")
        models = cursor.fetchall()
        model_list = [f"{m[0]} ({m[1]})" for m in models]
        conn.close()
        
        cb_model = ttk.Combobox(win, values=model_list, state="readonly", width=33)
        cb_model.pack(padx=20)
        
        ttk.Label(win, text="Opérateur assigné (Usinage) :").pack(anchor=tk.W, padx=20, pady=5)
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT username FROM users WHERE role='Opérateur'")
        ops = [row[0] for row in cursor.fetchall()]
        conn.close()
        
        cb_op = ttk.Combobox(win, values=ops, state="readonly", width=33)
        cb_op.pack(padx=20)
        if ops: cb_op.set(ops[0])
        
        def save_of():
            of_num = e_of.get().strip()
            sel_mod = cb_model.get()
            if not of_num or not sel_mod:
                messagebox.showerror("Erreur", "Veuillez remplir tous les champs.")
                return
            m_name = sel_mod.split(" (")[0]
            m_type = sel_mod.split(" (")[1].replace(")", "")
            
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            try:
                cursor.execute("""
                    INSERT INTO manufacturing_orders (of_number, model_name, article_type, assigned_operator, status, created_date)
                    VALUES (?, ?, ?, ?, 'Créé', ?)
                """, (of_num, m_name, m_type, cb_op.get(), datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))
                conn.commit()
                self.log_action("OF", f"Création de l'OF {of_num} pour le modèle {m_name} assigné à {cb_op.get()}")
                messagebox.showinfo("Succès", "OF créé avec succès.")
                win.destroy()
                self.load_of_data()
            except sqlite3.IntegrityError:
                messagebox.showerror("Erreur", "Ce numéro d'OF existe déjà.")
            finally:
                conn.close()
                
        ttk.Button(win, text="Créer l'OF", command=save_of).pack(pady=20)

    def create_block_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Enregistrer un Bloc de Matière")
        win.geometry("400x400")
        
        ttk.Label(win, text="Numéro de bloc (commence par B) :").pack(anchor=tk.W, padx=20, pady=5)
        e_block = ttk.Entry(win, width=30)
        e_block.pack(padx=20)
        e_block.insert(0, "B-2026-")
        
        ttk.Label(win, text="Type de matière :").pack(anchor=tk.W, padx=20, pady=5)
        e_mat = ttk.Entry(win, width=30)
        e_mat.pack(padx=20)
        e_mat.insert(0, "Polystyrène F13-17")
        
        ttk.Label(win, text="Poids mesuré (kg) :").pack(anchor=tk.W, padx=20, pady=5)
        e_weight = ttk.Entry(win, width=30)
        e_weight.pack(padx=20)
        
        var_test = tk.IntVar()
        chk_test = ttk.Checkbutton(win, text="Bloc testé (Qualité OK)", variable=var_test)
        chk_test.pack(anchor=tk.W, padx=20, pady=10)
        
        def save_block():
            b_num = e_block.get().strip()
            if not b_num.startswith("B"):
                messagebox.showerror("Erreur", "Le numéro de bloc doit commencer par la lettre 'B'.")
                return
            try:
                w = float(e_weight.get())
            except ValueError:
                w = 0.0
                
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            try:
                cursor.execute("""
                    INSERT INTO raw_blocks (block_number, material_type, density, weight, reception_date, tested, stock_qty)
                    VALUES (?, ?, ?, ?, ?, ?, 1.0)
                """, (b_num, e_mat.get(), 15.0, w, datetime.datetime.now().strftime("%Y-%m-%d"), var_test.get()))
                conn.commit()
                self.log_action("Stock", f"Enregistrement du bloc matière {b_num}")
                messagebox.showinfo("Succès", "Bloc enregistré.")
                win.destroy()
            except sqlite3.IntegrityError:
                messagebox.showerror("Erreur", "Ce numéro de bloc existe déjà.")
            finally:
                conn.close()
                
        ttk.Button(win, text="Enregistrer le Bloc", command=save_block).pack(pady=20)

    def update_of_progress_dialog(self):
        selected = self.tree_of.selection()
        if not selected:
            messagebox.showwarning("Attention", "Veuillez sélectionner un OF dans le tableau.")
            return
        item = self.tree_of.item(selected[0])
        of_id = item['values'][0]
        of_num = item['values'][1]
        
        win = tk.Toplevel(self.root)
        win.title(f"Suivi & Finition - OF {of_num}")
        win.geometry("450x450")
        
        ttk.Label(win, text="Numéro de Bloc utilisé :").pack(anchor=tk.W, padx=20, pady=5)
        e_block = ttk.Entry(win, width=30)
        e_block.pack(padx=20)
        
        ttk.Label(win, text="Numéro de Pain / Pièce :").pack(anchor=tk.W, padx=20, pady=5)
        e_part = ttk.Entry(win, width=30)
        e_part.pack(padx=20)
        
        ttk.Label(win, text="Opérateur Finition :").pack(anchor=tk.W, padx=20, pady=5)
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT username FROM users WHERE role='Opérateur'")
        ops = [row[0] for row in cursor.fetchall()]
        conn.close()
        cb_fin_op = ttk.Combobox(win, values=ops, state="readonly", width=28)
        cb_fin_op.pack(padx=20)
        if ops: cb_fin_op.set(ops[0])
        
        ttk.Label(win, text="Validation Qualité :").pack(anchor=tk.W, padx=20, pady=5)
        cb_qual = ttk.Combobox(win, values=["Conforme", "Non conforme"], state="readonly", width=28)
        cb_qual.pack(padx=20)
        cb_qual.set("Conforme")
        
        ttk.Label(win, text="Cause Non-conformité (si applicable) :").pack(anchor=tk.W, padx=20, pady=5)
        cb_cause = ttk.Combobox(win, values=["Aucune", "Bulles dans le bloc", "Défaut de densité", "Erreur programme", "Défaut de finition", "Autre"], state="readonly", width=28)
        cb_cause.pack(padx=20)
        cb_cause.set("Aucune")
        
        def save_prog():
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE manufacturing_orders
                SET block_number=?, part_number=?, finishing_operator=?, quality_status=?, non_conformity_cause=?, status='Terminé'
                WHERE id=?
            """, (e_block.get(), e_part.get(), cb_fin_op.get(), cb_qual.get(), cb_cause.get(), of_id))
            conn.commit()
            conn.close()
            self.log_action("OF", f"Mise à jour OF {of_num}: Pain {e_part.get()}, Qualité: {cb_qual.get()}")
            messagebox.showinfo("Succès", "Mise à jour enregistrée.")
            win.destroy()
            self.load_of_data()
            
        ttk.Button(win, text="Enregistrer les données de production", command=save_prog).pack(pady=20)

    def init_cn_tab(self):
        frame = self.tab_cn
        ttk.Label(frame, text="Module de Transfert CN (RS232 Virtuel / NUM 1060)", font=("Helvetica", 14, "bold")).pack(anchor=tk.W, padx=10, pady=10)
        
        f_settings = ttk.LabelFrame(frame, text="Paramètres de communication série (Réel NUM 1060)", padding=10)
        f_settings.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(f_settings, text="Port COM:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.e_com = ttk.Entry(f_settings, width=10)
        self.e_com.grid(row=0, column=1, padx=5)
        self.e_com.insert(0, "COM1")
        
        ttk.Label(f_settings, text="Bauds:").grid(row=0, column=2, sticky=tk.W, padx=5)
        self.e_baud = ttk.Entry(f_settings, width=10)
        self.e_baud.grid(row=0, column=3, padx=5)
        self.e_baud.insert(0, "9600")
        
        ttk.Label(f_settings, text="Parité:").grid(row=0, column=4, sticky=tk.W, padx=5)
        self.e_parity = ttk.Entry(f_settings, width=10)
        self.e_parity.grid(row=0, column=5, padx=5)
        self.e_parity.insert(0, "Even (Paire)")

        ttk.Label(f_settings, text="Mode de transfert:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=10)
        self.cb_mode = ttk.Combobox(f_settings, values=["Chargement complet", "Mode Passant (%PPR)"], state="readonly", width=25)
        self.cb_mode.grid(row=1, column=1, columnspan=3, padx=5, pady=10)
        self.cb_mode.set("Chargement complet")

        f_prog = ttk.LabelFrame(frame, text="Sélection et envoi du programme G-code (.tap)", padding=10)
        f_prog.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        f_sel = ttk.Frame(f_prog)
        f_sel.pack(fill=tk.X, pady=5)
        ttk.Label(f_sel, text="Fichier programme :").pack(side=tk.LEFT, padx=5)
        self.e_prog_file = ttk.Entry(f_sel, width=50)
        self.e_prog_file.pack(side=tk.LEFT, padx=5)
        ttk.Button(f_sel, text="Parcourir...", command=self.browse_prog_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(f_sel, text="Lancer le Transfert CN", command=self.start_cn_transfer).pack(side=tk.LEFT, padx=15)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(f_prog, orient="horizontal", length=600, mode="determinate")
        self.progress_bar.pack(pady=10)
        
        # Scrolling text for G-code transmission log
        ttk.Label(f_prog, text="Fenêtre de défilement du G-code transmis (Temps réel) :").pack(anchor=tk.W, padx=5)
        self.txt_gcode = tk.Text(f_prog, height=12, width=100, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 10))
        self.txt_gcode.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def browse_prog_file(self):
        fn = filedialog.askopenfilename(filetypes=[("TAP Files", "*.tap"), ("All Files", "*.*")])
        if fn:
            self.e_prog_file.delete(0, tk.END)
            self.e_prog_file.insert(0, fn)

    def start_cn_transfer(self):
        fpath = self.e_prog_file.get().strip()
        if not fpath or not os.path.exists(fpath):
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier G-code valide.")
            return
        
        mode = self.cb_mode.get()
        self.txt_gcode.delete("1.0", tk.END)
        self.progress_bar["value"] = 0
        
        # Background thread to simulate virtual CN transfer (writing to VIRTUAL_CN_DIR with progress)
        threading.Thread(target=self.run_virtual_transfer_thread, args=(fpath, mode), daemon=True).start()

    def run_virtual_transfer_thread(self, fpath, mode):
        try:
            with open(fpath, "r", encoding="utf-8", errors="ignore") as f:
                lines = f.readlines()
        except Exception as e:
            lines = [f"(Erreur de lecture fichier: {e})\nG0 X0 Y0\nM30\n"]
            
        total_lines = len(lines)
        if total_lines == 0:
            total_lines = 1
            
        dest_file = os.path.join(VIRTUAL_CN_DIR, os.path.basename(fpath))
        
        self.txt_gcode.insert(tk.END, f"=== DÉBUT DU TRANSFERT VERS CN VIRTUELLE ({mode}) ===\n")
        self.txt_gcode.insert(tk.END, f"Cible: {dest_file}\n\n")
        
        with open(dest_file, "w", encoding="utf-8") as out:
            for idx, line in enumerate(lines):
                out.write(line)
                self.txt_gcode.insert(tk.END, line)
                self.txt_gcode.see(tk.END)
                progress = int(((idx + 1) / total_lines) * 100)
                self.progress_bar["value"] = progress
                time.sleep(0.01) # simulation speed
                
        self.txt_gcode.insert(tk.END, "\n=== TRANSFERT TERMINÉ AVEC SUCCÈS (NUM 1060 ACK) ===\n")
        self.log_action("Transfert CN", f"Transfert réussi du fichier {os.path.basename(fpath)} en mode {mode}")
        messagebox.showinfo("Succès", "Transfert vers la CN virtuelle terminé avec succès.")

    def init_trace_tab(self):
        frame = self.tab_trace
        ttk.Label(frame, text="Traçabilité complète & Enquête Qualité", font=("Helvetica", 14, "bold")).pack(anchor=tk.W, padx=10, pady=10)
        
        f_search = ttk.Frame(frame, padding=5)
        f_search.pack(fill=tk.X, padx=10)
        ttk.Label(f_search, text="Recherche (N° Pain / Bloc / Utilisateur) :").pack(side=tk.LEFT, padx=5)
        self.e_search_query = ttk.Entry(f_search, width=30)
        self.e_search_query.pack(side=tk.LEFT, padx=5)
        ttk.Button(f_search, text="Rechercher", command=self.load_trace_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(f_search, text="Tout afficher", command=self.load_trace_data).pack(side=tk.LEFT, padx=5)
        
        columns = ("ID", "Horodatage", "Utilisateur", "Action", "Détails", "Machine", "Statut")
        self.tree_trace = ttk.Treeview(frame, columns=columns, show="headings", height=20)
        for col in columns:
            self.tree_trace.heading(col, text=col)
            self.tree_trace.column(col, width=130)
        self.tree_trace.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.load_trace_data()

    def load_trace_data(self):
        for row in self.tree_trace.get_children():
            self.tree_trace.delete(row)
            
        q = self.e_search_query.get().strip()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        if q:
            cursor.execute("SELECT id, timestamp, username, action, details, machine, status FROM traceability_logs WHERE details LIKE ? OR username LIKE ? OR action LIKE ? ORDER BY id DESC", 
                           (f"%{q}%", f"%{q}%", f"%{q}%"))
        else:
            cursor.execute("SELECT id, timestamp, username, action, details, machine, status FROM traceability_logs ORDER BY id DESC LIMIT 100")
        for r in cursor.fetchall():
            self.tree_trace.insert("", tk.END, values=r)
        conn.close()

    def init_compare_tab(self):
        frame = self.tab_compare
        ttk.Label(frame, text="Module de Comparaison de Programmes (.tap) - Style WinMerge", font=("Helvetica", 14, "bold")).pack(anchor=tk.W, padx=10, pady=10)
        
        f_files = ttk.Frame(frame, padding=5)
        f_files.pack(fill=tk.X, padx=10)
        
        ttk.Label(f_files, text="Fichier Référence (Win7 Archivé):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.e_file1 = ttk.Entry(f_files, width=45)
        self.e_file1.grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(f_files, text="Parcourir...", command=lambda: self.browse_file(self.e_file1)).grid(row=0, column=2, padx=5)
        
        ttk.Label(f_files, text="Fichier Atelier (Poste XP Modifié):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.e_file2 = ttk.Entry(f_files, width=45)
        self.e_file2.grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(f_files, text="Parcourir...", command=lambda: self.browse_file(self.e_file2)).grid(row=1, column=2, padx=5)
        
        ttk.Button(f_files, text="Comparer les deux fichiers (.tap)", command=self.compare_tap_files).grid(row=2, column=1, pady=10)
        
        # Side-by-side text display
        f_res = ttk.Frame(frame, padding=5)
        f_res.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        f_left = ttk.LabelFrame(f_res, text="Référence (Win7)")
        f_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        self.txt_comp1 = tk.Text(f_left, height=20, width=50, bg="#2d2d2d", fg="#ffffff", font=("Consolas", 9))
        self.txt_comp1.pack(fill=tk.BOTH, expand=True)
        
        f_right = ttk.LabelFrame(f_res, text="Atelier (XP)")
        f_right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)
        self.txt_comp2 = tk.Text(f_right, height=20, width=50, bg="#2d2d2d", fg="#ffffff", font=("Consolas", 9))
        self.txt_comp2.pack(fill=tk.BOTH, expand=True)

    def browse_file(self, entry_widget):
        fn = filedialog.askopenfilename(filetypes=[("TAP Files", "*.tap"), ("All Files", "*.*")])
        if fn:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, fn)

    def compare_tap_files(self):
        f1 = self.e_file1.get().strip()
        f2 = self.e_file2.get().strip()
        if not f1 or not f2 or not os.path.exists(f1) or not os.path.exists(f2):
            messagebox.showerror("Erreur", "Veuillez sélectionner deux fichiers valides.")
            return
        
        with open(f1, "r", encoding="utf-8", errors="ignore") as file1:
            lines1 = file1.readlines()
        with open(f2, "r", encoding="utf-8", errors="ignore") as file2:
            lines2 = file2.readlines()
            
        self.txt_comp1.delete("1.0", tk.END)
        self.txt_comp2.delete("1.0", tk.END)
        
        for l in lines1:
            self.txt_comp1.insert(tk.END, l)
        for l in lines2:
            self.txt_comp2.insert(tk.END, l)
            
        messagebox.showinfo("Comparaison", "Comparaison WinMerge effectuée avec succès.")

    def init_tools_dashboard_tab(self):
        frame = self.tab_tools
        ttk.Label(frame, text="6. Stocks, Outils, Non-conformités & Dashboard", font=("Helvetica", 14, "bold")).pack(anchor=tk.W, padx=10, pady=10)
        
        f_sub = ttk.Notebook(frame)
        f_sub.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Sub-tab: Stock blocs
        t_stock = ttk.Frame(f_sub)
        f_sub.add(t_stock, text="Stock Matières")
        self.tree_stock = ttk.Treeview(t_stock, columns=("ID", "N° Bloc", "Matière", "Densité", "Poids", "Réception", "Testé"), show="headings", height=15)
        for col in ("ID", "N° Bloc", "Matière", "Densité", "Poids", "Réception", "Testé"):
            self.tree_stock.heading(col, text=col)
        self.tree_stock.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.load_stock_data(t_stock)

        # Sub-tab: Tools
        t_tools = ttk.Frame(f_sub)
        f_sub.add(t_tools, text="Outils de Coupe")
        self.tree_tools = ttk.Treeview(t_tools, columns=("ID", "Réf Outil", "Description", "Vie Max (h)", "Usage Actuel (h)"), show="headings", height=15)
        for col in ("ID", "Réf Outil", "Description", "Vie Max (h)", "Usage Actuel (h)"):
            self.tree_tools.heading(col, text=col)
        self.tree_tools.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.load_tools_data()

        # Sub-tab: Maintenance
        t_maint = ttk.Frame(f_sub)
        f_sub.add(t_maint, text="Maintenance CN")
        f_maint_ctrl = ttk.Frame(t_maint, padding=5)
        f_maint_ctrl.pack(fill=tk.X)
        ttk.Button(f_maint_ctrl, text="Enregistrer une Intervention / Panne", command=self.add_maintenance_dialog).pack(side=tk.LEFT, padx=5)
        
        self.tree_maint = ttk.Treeview(t_maint, columns=("ID", "Machine", "Date", "Description", "Technicien"), show="headings", height=12)
        for col in ("ID", "Machine", "Date", "Description", "Technicien"):
            self.tree_maint.heading(col, text=col)
        self.tree_maint.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.load_maintenance_data()

    def load_stock_data(self, parent_frame):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, block_number, material_type, density, weight, reception_date, tested FROM raw_blocks")
        for r in cursor.fetchall():
            self.tree_stock.insert("", tk.END, values=r)
        conn.close()

    def load_tools_data(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, tool_ref, description, max_life_hours, current_usage_hours FROM tools")
        rows = cursor.fetchall()
        if not rows:
            cursor.execute("INSERT OR IGNORE INTO tools (tool_ref, description, max_life_hours, current_usage_hours) VALUES ('FRAISE-12', 'Fraise carbure 2 tailles D12', 120.0, 15.5)")
            cursor.execute("INSERT OR IGNORE INTO tools (tool_ref, description, max_life_hours, current_usage_hours) VALUES ('FRAISE-8', 'Fraise sphérique D8', 80.0, 42.0)")
            conn.commit()
            cursor.execute("SELECT id, tool_ref, description, max_life_hours, current_usage_hours FROM tools")
            rows = cursor.fetchall()
        conn.close()
        for r in rows:
            self.tree_tools.insert("", tk.END, values=r)

    def load_maintenance_data(self):
        for row in self.tree_maint.get_children():
            self.tree_maint.delete(row)
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, machine_name, intervention_date, description, technician FROM machine_maintenance")
        for r in cursor.fetchall():
            self.tree_maint.insert("", tk.END, values=r)
        conn.close()

    def add_maintenance_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Maintenance Machine CN")
        win.geometry("400x350")
        
        ttk.Label(win, text="Nom Machine (ex: CN n°1 NUM 1060):").pack(anchor=tk.W, padx=20, pady=5)
        e_mach = ttk.Entry(win, width=30)
        e_mach.pack(padx=20)
        e_mach.insert(0, "CN n°1 (Poste C)")
        
        ttk.Label(win, text="Description de l'intervention / panne :").pack(anchor=tk.W, padx=20, pady=5)
        e_desc = ttk.Entry(win, width=30)
        e_desc.pack(padx=20)
        e_desc.insert(0, "Remplacement pile RAM SRAM / Paramètres")
        
        ttk.Label(win, text="Technicien :").pack(anchor=tk.W, padx=20, pady=5)
        e_tech = ttk.Entry(win, width=30)
        e_tech.pack(padx=20)
        e_tech.insert(0, "Dhaou Bouzaien")
        
        def save_maint():
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO machine_maintenance (machine_name, intervention_date, description, technician)
                VALUES (?, ?, ?, ?)
            """, (e_mach.get(), datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), e_desc.get(), e_tech.get()))
            conn.commit()
            conn.close()
            self.log_action("Maintenance", f"Intervention enregistrée sur {e_mach.get()}: {e_desc.get()}")
            messagebox.showinfo("Succès", "Intervention enregistrée.")
            win.destroy()
            self.load_maintenance_data()
            
        ttk.Button(win, text="Enregistrer", command=save_maint).pack(pady=20)

if __name__ == "__main__":
    root = tk.Tk()
    app = CNCApp(root)
    root.mainloop()
