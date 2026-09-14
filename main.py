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
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# Informations de l'application (Section "À propos")
APP_NAME = "Gestionnaire de Sauvegarde des Programmes CNC"
APP_VERSION = "2.1.0"
APP_AUTHOR = "Bouzaien Dhaou"
APP_EMAIL = "bouzaien.dhaou@gmail.com"
DATE_CREATED = "14/09/2026"
DATE_MODIFIED = "14/09/2026"

# Configuration et mot de passe (Crypté SHA-256)
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


def load_config():
    if not os.path.exists(CONFIG_FILE):
        config = {"password_hash": DEFAULT_PASSWORD_HASH}
        save_config(config)
        return config
    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {"password_hash": DEFAULT_PASSWORD_HASH}


def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)


# ==========================================
# BOÎTE DE PROGRESSION VERTE POUR SAUVEGARDE
# ==========================================
class BackupProgressBarDialog(tk.Toplevel):
    def __init__(self, parent, title="Sauvegarde en cours..."):
        super().__init__(parent)
        self.title(title)
        self.geometry("450x200")
        self.resizable(False, False)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", lambda: None)

        self.label_status = tk.Label(self, text="Préparation de la sauvegarde...", font=("Arial", 10, "bold"))
        self.label_status.pack(pady=15)

        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("Green.Horizontal.TProgressbar", foreground='#4CAF50', background='#4CAF50', thickness=20)

        self.progress = ttk.Progressbar(self, style="Green.Horizontal.TProgressbar", length=350, mode='determinate')
        self.progress.pack(pady=10)

        self.label_percent = tk.Label(self, text="0%", font=("Arial", 10))
        self.label_percent.pack()

        self.btn_close = tk.Button(self, text="Fermer", font=("Arial", 10, "bold"), bg="#2E7D32", fg="white", state=tk.DISABLED, command=self.destroy)
        self.btn_close.pack(pady=15)

    def update_progress(self, current, total, filename=""):
        percent = int((current / total) * 100) if total > 0 else 100
        self.progress['value'] = percent
        self.label_percent.config(text=f"{percent}% ({current}/{total})")
        if filename:
            self.label_status.config(text=f"Copie : {filename}")
        self.update()

    def complete(self):
        self.progress['value'] = 100
        self.label_percent.config(text="100% - Sauvegarde terminée !")
        self.label_status.config(text="Sauvegarde réalisée avec succès.")
        self.btn_close.config(state=tk.NORMAL)
        self.protocol("WM_DELETE_WINDOW", self.destroy)


# ==========================================
# BOÎTE DE DIALOGUE D'AUTHENTIFICATION
# ==========================================
class LoginDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Authentification requise")
        self.geometry("360x190")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.authenticated = False

        self.config = load_config()

        lbl = tk.Label(self, text="Gestionnaire de Sauvegarde CNC", font=("Arial", 11, "bold"))
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
        if hashed == self.config["password_hash"]:
            self.authenticated = True
            self.destroy()
        else:
            messagebox.showerror("Erreur", "Mot de passe incorrect !", parent=self)
            self.ent_pass.delete(0, tk.END)

    def on_close(self):
        self.authenticated = False
        self.destroy()


# ==========================================
# APPLICATION PRINCIPALE
# ==========================================
class CNCBackupManagerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1020x720")
        self.minsize(950, 680)

        self.tasks = []

        # Authentification au démarrage
        self.withdraw()
        login = LoginDialog(self)
        self.wait_window(login)

        if not login.authenticated:
            self.destroy()
            sys.exit()

        self.deiconify()
        self.build_ui()

    def build_ui(self):
        # Barre de menus
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

        # En-tête principal
        header_frame = tk.Frame(self, bg="#003366", height=50)
        header_frame.pack(fill=tk.X, side=tk.TOP)
        title_label = tk.Label(header_frame, text=APP_NAME.upper(), font=("Arial", 13, "bold"), fg="white", bg="#003366", pady=10)
        title_label.pack()

        # Système d'onglets
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tab_backup = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_backup, text=" Configuration & Planification des Sauvegardes ")

        self.tab_compare = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_compare, text=" Comparaison de Dossiers ")

        self.setup_backup_tab()
        self.setup_compare_tab()

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
    # ONGLET 1 : SAUVEGARDE ET PLANIFICATION
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

        # Jours d'exécution hebdomadaire
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

        # Jour du mois pour le mode Mensuel
        ttk.Label(frame_add, text="Jour du mois (1-31) :").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.spin_month_day = tk.Spinbox(frame_add, from_=1, to=31, width=5, format="%02.0f", state=tk.DISABLED)
        self.spin_month_day.grid(row=5, column=1, sticky="w", padx=5, pady=5)

        # Heure d'exécution
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

        # Liste des sauvegardes
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
        else:  # Journalier
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
        if selected:
            self.tree_tasks.delete(selected[0])

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

        cols = ("statut", "fichier", "date_a", "date_b")
        self.tree = ttk.Treeview(frame_grid, columns=cols, show="headings", selectmode="browse")
        
        self.tree.heading("statut", text="Statut")
        self.tree.heading("fichier", text="Fichier / Chemin Relatif")
        self.tree.heading("date_a", text="Date Modification (A)")
        self.tree.heading("date_b", text="Date Modification (B)")

        self.tree.column("statut", width=120, anchor="center")
        self.tree.column("fichier", width=400, anchor="w")
        self.tree.column("date_a", width=180, anchor="center")
        self.tree.column("date_b", width=180, anchor="center")

        # Coloration : ROUGE avec texte BLANC pour statut DIFFERENT
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

            tag = "different_tag" if statut == "DIFFERENT" else "identique_tag"
            self.tree.insert("", tk.END, values=(statut, rel, date_a_str, date_b_str), tags=(tag,))

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

    # --- IMPRESSION FORMAT A4 PORTRAIT AUTO-AJUSTABLE ---
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
        elements.append(Paragraph("<b>RAPPORT DE COMPARAISON DES PROGRAMMES CNC</b>", title_style))
        elements.append(Spacer(1, 10))

        data = [[
            Paragraph("<b>Statut</b>", cell_style_bold),
            Paragraph("<b>Fichier / Chemin Relatif</b>", cell_style_bold),
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
            statut, fichier, date_a, date_b = r
            
            p_statut = Paragraph(f"<b>{statut}</b>", cell_style)
            p_fichier = Paragraph(fichier, cell_style)
            p_date_a = Paragraph(date_a, cell_style)
            p_date_b = Paragraph(date_b, cell_style)

            data.append([p_statut, p_fichier, p_date_a, p_date_b])

            if statut == "DIFFERENT":
                table_styles.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor("#D32F2F")))
                cell_white = ParagraphStyle('CellW', parent=cell_style, textColor=colors.white)
                data[i] = [
                    Paragraph(f"<b>{statut}</b>", cell_white),
                    Paragraph(fichier, cell_white),
                    Paragraph(date_a, cell_white),
                    Paragraph(date_b, cell_white)
                ]

        t = Table(data, colWidths=[70, 260, 110, 110])
        t.setStyle(TableStyle(table_styles))
        elements.append(t)

        doc.build(elements)
        os.startfile(pdf_filename, "print") if hasattr(os, "startfile") else webbrowser.open(pdf_filename)

    def generate_html_print(self, rows):
        html_filename = os.path.join(tempfile.gettempdir(), "rapport_cnc_a4.html")
        
        html_content = """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Rapport CNC A4</title>
            <style>
                @page { size: A4 portrait; margin: 10mm; }
                body { font-family: Arial, sans-serif; font-size: 9pt; margin: 0; }
                h2 { text-align: center; margin-bottom: 15px; font-size: 12pt; }
                table { width: 100%; border-collapse: collapse; table-layout: fixed; }
                th, td { border: 1px solid #666; padding: 4px 6px; word-wrap: break-word; font-size: 8pt; }
                th { background-color: #0B3C5D; color: white; text-align: left; }
                tr.different { background-color: #D32F2F !important; color: white !important; font-weight: bold; }
                col.c1 { width: 15%; }
                col.c2 { width: 45%; }
                col.c3 { width: 20%; }
                col.c4 { width: 20%; }
            </style>
        </head>
        <body onload="window.print();">
            <h2>RAPPORT DE COMPARAISON DES PROGRAMMES CNC</h2>
            <table>
                <colgroup>
                    <col class="c1"><col class="c2"><col class="c3"><col class="c4">
                </colgroup>
                <thead>
                    <tr>
                        <th>Statut</th>
                        <th>Fichier / Chemin Relatif</th>
                        <th>Date Modification (A)</th>
                        <th>Date Modification (B)</th>
                    </tr>
                </thead>
                <tbody>
        """
        for r in rows:
            statut, fichier, date_a, date_b = r
            row_class = "different" if statut == "DIFFERENT" else ""
            html_content += f"""
                <tr class="{row_class}">
                    <td>{statut}</td>
                    <td>{fichier}</td>
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
                writer.writerow(["Statut", "Fichier", "Date A", "Date B"])
                for item in self.tree.get_children():
                    writer.writerow(self.tree.item(item, "values"))
            messagebox.showinfo("Export", "Export CSV/Excel réussi !")


if __name__ == "__main__":
    app = CNCBackupManagerApp()
    app.mainloop()
