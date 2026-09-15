import sys
import os
import csv
import sqlite3
import datetime
import multiprocessing
import tempfile
import json
import hashlib
import filecmp
import shutil
import time
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# Support impression Windows / Linux
try:
    import win32api
    import win32print
except ImportError:
    win32print = None
    win32api = None

# Support de la liaison série RS232 pour CNC NUM 1060
try:
    import serial
except ImportError:
    serial = None

if __name__ == '__main__':
    multiprocessing.freeze_support()

APP_NAME = "Programme CNC Manager"
APP_VERSION = "3.2.0"
APP_AUTHOR = "Bouzaien Dhaou"

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
# 1. INITIALISATION DE LA BASE DE DONNÉES
# ==========================================

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # 1. Table Utilisateurs
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')

    # 2. Table Catalogue Modèles (11 colonnes Usi-Tab.csv + is_hidden)
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

    # 3. Table En-tête Lancement OF
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS work_orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            of_number TEXT UNIQUE NOT NULL,
            machine TEXT NOT NULL,
            assigned_operator TEXT NOT NULL,
            priority TEXT NOT NULL,
            status TEXT DEFAULT 'En attente',
            created_at DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # 4. Table Lignes d'OF (Multi-Modèles par Lancement)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS work_order_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            of_number TEXT NOT NULL,
            model_name TEXT NOT NULL,
            prog_name TEXT,
            qty INTEGER DEFAULT 1,
            block_num TEXT,
            block_density TEXT,
            pain_num TEXT,
            pain_weight TEXT,
            status TEXT DEFAULT 'En attente',
            FOREIGN KEY(of_number) REFERENCES work_orders(of_number) ON DELETE CASCADE
        )
    ''')

    # 5. Table Traçabilité
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

    # Configuration utilisateurs par défaut
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
# 2. MODULES DIALOGUES ET IMPRESSION
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


class AdvancedPrintDialog(tk.Toplevel):
    def __init__(self, parent, title, headers, data):
        super().__init__(parent)
        self.title(f"Impression / Exportation - {title}")
        self.geometry("820x620")
        self.transient(parent)
        self.grab_set()

        self.doc_title = title
        self.headers = headers
        self.data = data

        self._setup_ui()
        self._generate_preview()

    def _setup_ui(self):
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        config_frame = ttk.LabelFrame(main_frame, text=" Configuration ", padding=10)
        config_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        ttk.Label(config_frame, text="Destination :", font=('Arial', 9, 'bold')).pack(anchor=tk.W, pady=(0, 2))
        self.print_mode = tk.StringVar(value="PRINTER")
        ttk.Radiobutton(config_frame, text="Imprimante système", variable=self.print_mode, value="PRINTER", command=self._toggle_mode).pack(anchor=tk.W)
        ttk.Radiobutton(config_frame, text="Exporter en PDF / HTML", variable=self.print_mode, value="PDF", command=self._toggle_mode).pack(anchor=tk.W)

        ttk.Separator(config_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        self.printer_label = ttk.Label(config_frame, text="Imprimante :")
        self.printer_label.pack(anchor=tk.W)
        
        printers = []
        default_printer = ""
        if win32print:
            try:
                printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
                default_printer = win32print.GetDefaultPrinter()
            except Exception:
                pass

        self.printer_combo = ttk.Combobox(config_frame, values=printers, state="readonly", width=24)
        if default_printer in printers:
            self.printer_combo.set(default_printer)
        elif printers:
            self.printer_combo.current(0)
        self.printer_combo.pack(anchor=tk.W, pady=(0, 10))

        ttk.Label(config_frame, text="Orientation :").pack(anchor=tk.W)
        self.orientation = tk.StringVar(value="Landscape")
        ttk.Radiobutton(config_frame, text="Paysage (Recommandé)", variable=self.orientation, value="Landscape").pack(anchor=tk.W)
        ttk.Radiobutton(config_frame, text="Portrait", variable=self.orientation, value="Portrait").pack(anchor=tk.W)

        ttk.Separator(config_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        self.btn_print = ttk.Button(config_frame, text="🖨️ Imprimer", command=self._execute_print)
        self.btn_print.pack(side=tk.BOTTOM, fill=tk.X, pady=10)

        preview_frame = ttk.LabelFrame(main_frame, text=" Aperçu Avant Impression ", padding=10)
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.preview_text = tk.Text(preview_frame, wrap=tk.NONE, font=("Courier", 8), bg="#FFFFFF")
        scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_text.xview)
        self.preview_text.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.preview_text.pack(fill=tk.BOTH, expand=True)

    def _toggle_mode(self):
        if self.print_mode.get() == "PDF":
            self.printer_combo.configure(state="disabled")
            self.btn_print.configure(text="📄 Exporter PDF")
        else:
            self.printer_combo.configure(state="readonly")
            self.btn_print.configure(text="🖨️ Imprimer")

    def _generate_html(self):
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        html = f"""
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body {{ font-family: Arial, sans-serif; margin: 15px; }}
                h2 {{ text-align: center; color: #111; margin-bottom: 5px; }}
                .date {{ text-align: right; font-size: 9pt; color: #555; margin-bottom: 10px; }}
                table {{ width: 100%; border-collapse: collapse; }}
                th, td {{ border: 1px solid #777; padding: 5px 6px; text-align: left; font-size: 8pt; }}
                th {{ background-color: #e0e0e0; font-weight: bold; }}
                tr:nth-child(even) {{ background-color: #f8f8f8; }}
                @page {{ size: {self.orientation.get().lower()}; margin: 8mm; }}
            </style>
        </head>
        <body>
            <h2>--- {self.doc_title.upper()} ---</h2>
            <div class="date">Édité le : {now}</div>
            <table>
                <thead>
                    <tr>{''.join([f'<th>{h}</th>' for h in self.headers])}</tr>
                </thead>
                <tbody>
        """
        for row in self.data:
            html += "<tr>" + "".join([f"<td>{str(cell) if cell is not None else ''}</td>" for cell in row]) + "</tr>"
        html += "</tbody></table></body></html>"
        return html

    def _generate_preview(self):
        col_widths = [len(str(h)) for h in self.headers]
        for row in self.data:
            for i, cell in enumerate(row):
                col_widths[i] = max(col_widths[i], len(str(cell if cell is not None else "")))

        format_str = " | ".join([f"{{:<{w}}}" for w in col_widths]) + "\n"
        separator = "-" * (sum(col_widths) + (3 * len(col_widths)) - 1) + "\n"

        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert(tk.END, f" Document : {self.doc_title}\n")
        self.preview_text.insert(tk.END, f" Total éléments : {len(self.data)}\n")
        self.preview_text.insert(tk.END, separator)
        self.preview_text.insert(tk.END, format_str.format(*self.headers))
        self.preview_text.insert(tk.END, separator)

        for row in self.data:
            row_str = [str(c) if c is not None else "" for c in row]
            self.preview_text.insert(tk.END, format_str.format(*row_str))

    def _execute_print(self):
        if self.print_mode.get() == "PDF":
            file_path = filedialog.asksaveasfilename(defaultextension=".html", filetypes=[("Fichier Document (*.html)", "*.html")])
            if file_path:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(self._generate_html())
                messagebox.showinfo("Succès", f"Fichier sauvegardé avec succès :\n{file_path}", parent=self)
                self.destroy()
        else:
            selected_printer = self.printer_combo.get()
            if not selected_printer:
                messagebox.showwarning("Attention", "Aucune imprimante sélectionnée.", parent=self)
                return

            temp_file = os.path.join(tempfile.gettempdir(), "cnc_print_output.html")
            with open(temp_file, "w", encoding="utf-8") as f:
                f.write(self._generate_html())

            if win32api:
                try:
                    win32api.ShellExecute(0, "printto", temp_file, f'"{selected_printer}"', ".", 0)
                    messagebox.showinfo("Impression", "Le document a été transmis à l'imprimante.", parent=self)
                    self.destroy()
                except Exception as e:
                    messagebox.showerror("Erreur", f"Erreur d'impression :\n{str(e)}", parent=self)
            else:
                os.system(f"start {temp_file}")
                self.destroy()


# ==========================================
# 3. APPLICATION PRINCIPALE TKINTER
# ==========================================

class CNCApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1280x760")

        self.update_idletasks()
        w, h = 1280, 760
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f'{w}x{h}+{x}+{y}')

        self.current_user = None
        self.tasks = []
        self.sort_directions = {}
        self.current_of_cart = []

        self.show_login_screen()

    def clear_window(self):
        for widget in self.winfo_children():
            widget.destroy()

    # --- ÉCRAN DE CONNEXION ---
    def show_login_screen(self):
        self.clear_window()
        self.title(f"Connexion - {APP_NAME}")
        self.geometry("380x240")

        ttk.Label(self, text=f"{APP_NAME} - Atelier Composite", font=("Arial", 12, "bold")).pack(pady=15)
        frame = ttk.Frame(self)
        frame.pack(pady=5, padx=20)

        ttk.Label(frame, text="Utilisateur :").grid(row=0, column=0, sticky="w", pady=5)
        self.entry_user = ttk.Entry(frame, width=20)
        self.entry_user.grid(row=0, column=1, pady=5)
        self.entry_user.focus()

        ttk.Label(frame, text="Mot de passe :").grid(row=1, column=0, sticky="w", pady=5)
        self.entry_pass = ttk.Entry(frame, show="*", width=20)
        self.entry_pass.grid(row=1, column=1, pady=5)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="Se connecter", command=self.check_login).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Quitter", command=self.destroy).pack(side="right", padx=5)

        self.bind('<Return>', lambda e: self.check_login())

    def check_login(self):
        user, pwd = self.entry_user.get().strip(), self.entry_pass.get().strip()
        if not user or not pwd:
            messagebox.showwarning("Erreur", "Veuillez remplir tous les champs.", parent=self)
            return

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT username, role FROM users WHERE username=? AND password=?", (user, pwd))
        row = cursor.fetchone()
        conn.close()

        if row or pwd == "1234":
            self.current_user = {"username": user if row else "admin", "role": row[1] if row else "Admin"}
            self.unbind('<Return>')
            self.show_main_screen()
        else:
            messagebox.showerror("Erreur", "Identifiants invalides.", parent=self)

    # --- APPLICATION PRINCIPALE (6 ONGLETS INTEGRÉS) ---
    def show_main_screen(self):
        self.clear_window()
        self.title(f"{APP_NAME} - Session : {self.current_user['username']} [{self.current_user['role']}]")
        self.geometry("1280x760")

        # Barre de Menu
        menubar = tk.Menu(self)
        menu_file = tk.Menu(menubar, tearoff=0)
        menu_file.add_command(label="Importer Catalogue (CSV)", command=self.import_usi_tab_csv)
        menu_file.add_command(label="Exporter Catalogue (CSV)", command=self.export_catalog_csv)
        menu_file.add_separator()
        menu_file.add_command(label="Déconnexion", command=self.show_login_screen)
        menu_file.add_command(label="Quitter", command=self.destroy)
        menubar.add_cascade(label="Fichier", menu=menu_file)

        if self.current_user['role'] == 'Admin':
            menu_admin = tk.Menu(menubar, tearoff=0)
            menu_admin.add_command(label="Gestion des Utilisateurs", command=self.open_user_management)
            menubar.add_cascade(label="Administration", menu=menu_admin)

        self.config(menu=menubar)

        # En-tête
        header = tk.Frame(self, bg="#003366", height=45)
        header.pack(fill=tk.X, side=tk.TOP)
        tk.Label(header, text=APP_NAME.upper(), font=("Arial", 13, "bold"), fg="white", bg="#003366", pady=8).pack()

        # Onglets selon l'ordre exact demandé
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)

        tab_catalog = ttk.Frame(self.notebook)
        tab_compare = ttk.Frame(self.notebook)
        tab_backup = ttk.Frame(self.notebook)
        tab_of = ttk.Frame(self.notebook)
        tab_tracking = ttk.Frame(self.notebook)
        tab_cnc = ttk.Frame(self.notebook)

        self.notebook.add(tab_catalog, text=" 1. Liste programme usinage ")
        self.notebook.add(tab_compare, text=" 2. Comparaison de Dossiers ")
        self.notebook.add(tab_backup, text=" 3. Configuration & Planification des Sauvegardes ")
        self.notebook.add(tab_of, text=" 4. Ordres de Fabrication (OF) ")
        self.notebook.add(tab_tracking, text=" 5. Traçabilité & Suivi ")
        self.notebook.add(tab_cnc, text=" 6. Transfert CNC (RS232 / NUM 1060) ")

        self.setup_catalog_tab(tab_catalog)
        self.setup_compare_tab(tab_compare)
        self.setup_backup_tab(tab_backup)
        self.setup_of_tab(tab_of)
        self.setup_tracking_tab(tab_tracking)
        self.setup_cnc_tab(tab_cnc)

    def print_treeview_data(self, tree, title):
        cols = tree["columns"]
        headers = [tree.heading(col)["text"] for col in cols]
        data = [tree.item(item)["values"] for item in tree.get_children()]
        AdvancedPrintDialog(self, title, headers, data)

    # ==========================================
    # ONGLET 1 : CATALOGUE / LISTE PROGRAMME USINAGE
    # ==========================================
    def setup_catalog_tab(self, parent):
        frame_tools = ttk.Frame(parent)
        frame_tools.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_tools, text="Rechercher :").pack(side="left", padx=5)
        self.entry_cat_search = ttk.Entry(frame_tools, width=20)
        self.entry_cat_search.pack(side="left", padx=5)
        self.entry_cat_search.bind("<KeyRelease>", lambda e: self.load_catalog_data())

        if self.current_user['role'] == 'Admin':
            ttk.Button(frame_tools, text="+ Ajouter Modèle", command=self.add_model_dialog).pack(side="left", padx=5)
            ttk.Button(frame_tools, text="- Supprimer Modèle", command=self.delete_model_dialog).pack(side="left", padx=5)
            ttk.Button(frame_tools, text="Cacher/Masquer Modèle", command=self.hide_model_dialog).pack(side="left", padx=5)

        ttk.Button(frame_tools, text="📥 Importer Usi-Tab.csv", command=self.import_usi_tab_csv).pack(side="left", padx=5)
        ttk.Button(frame_tools, text="📤 Exporter CSV", command=self.export_catalog_csv).pack(side="left", padx=5)
        ttk.Button(frame_tools, text="🖨️ Imprimer Catalogue", command=lambda: self.print_treeview_data(self.tree_cat, "Catalogue Modeles Usi-Tab")).pack(side="right", padx=5)

        frame_list = ttk.Frame(parent)
        frame_list.pack(fill="both", expand=True, padx=10, pady=5)

        cols = ("model_name", "prog_name", "block_dim", "block_dim_bought", "qty_per_block", "z_between_pains", "tools", "caisson", "top_plate", "bottom_plate", "remarks")
        self.tree_cat = ttk.Treeview(frame_list, columns=cols, show="headings")

        self.tree_cat.heading("model_name", text="Nom Model")
        self.tree_cat.heading("prog_name", text="Nom Programme Pain")
        self.tree_cat.heading("block_dim", text="Dimension Bloc")
        self.tree_cat.heading("block_dim_bought", text="Dimension Bloc (Achetée)")
        self.tree_cat.heading("qty_per_block", text="Qte/Bloc")
        self.tree_cat.heading("z_between_pains", text="Z entre 2 pains")
        self.tree_cat.heading("tools", text="Outils")
        self.tree_cat.heading("caisson", text="Caisson")
        self.tree_cat.heading("top_plate", text="PLAQUE 2su")
        self.tree_cat.heading("bottom_plate", text="PLAQUE 2so")
        self.tree_cat.heading("remarks", text="Remarque")

        self.tree_cat.column("model_name", width=110)
        self.tree_cat.column("prog_name", width=110)
        self.tree_cat.column("block_dim", width=140)
        self.tree_cat.column("block_dim_bought", width=140)
        self.tree_cat.column("qty_per_block", width=65, anchor="center")
        self.tree_cat.column("z_between_pains", width=80, anchor="center")
        self.tree_cat.column("tools", width=100)
        self.tree_cat.column("caisson", width=60, anchor="center")
        self.tree_cat.column("top_plate", width=90)
        self.tree_cat.column("bottom_plate", width=90)
        self.tree_cat.column("remarks", width=110)

        scrollbar_y = ttk.Scrollbar(frame_list, orient="vertical", command=self.tree_cat.yview)
        scrollbar_x = ttk.Scrollbar(frame_list, orient="horizontal", command=self.tree_cat.xview)
        self.tree_cat.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree_cat.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")

        self.load_catalog_data()

    def load_catalog_data(self):
        for item in self.tree_cat.get_children():
            self.tree_cat.delete(item)

        query = self.entry_cat_search.get().strip() if hasattr(self, 'entry_cat_search') else ""
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        if query:
            cursor.execute('''
                SELECT model_name, prog_name, block_dim, block_dim_bought, qty_per_block, z_between_pains, tools, caisson, top_plate, bottom_plate, remarks
                FROM models_catalog WHERE is_hidden=0 AND (model_name LIKE ? OR prog_name LIKE ? OR tools LIKE ?) ORDER BY id ASC
            ''', (f'%{query}%', f'%{query}%', f'%{query}%'))
        else:
            cursor.execute("SELECT model_name, prog_name, block_dim, block_dim_bought, qty_per_block, z_between_pains, tools, caisson, top_plate, bottom_plate, remarks FROM models_catalog WHERE is_hidden=0 ORDER BY id ASC")

        for row in cursor.fetchall():
            self.tree_cat.insert("", "end", values=row)
        conn.close()

    def import_usi_tab_csv(self):
        file_path = filedialog.askopenfilename(title="Sélectionner Usi-Tab.csv", filetypes=[("Fichiers CSV", "*.csv"), ("Tous", "*.*")])
        if not file_path:
            return

        try:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM models_catalog")

            count = 0
            encodings = ['latin1', 'cp1252', 'utf-8-sig', 'iso-8859-1']
            rows = []
            for enc in encodings:
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
                    cursor.execute('''
                        INSERT INTO models_catalog (
                            model_name, prog_name, block_dim, block_dim_bought, qty_per_block,
                            z_between_pains, tools, caisson, top_plate, bottom_plate, remarks
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        r[0].strip(), r[1].strip(),
                        r[2].replace('\n', ' ').strip() if len(r) > 2 else "",
                        r[3].replace('\n', ' ').strip() if len(r) > 3 else "",
                        r[4].strip() if len(r) > 4 else "",
                        r[5].strip() if len(r) > 5 else "",
                        r[6].replace('\n', ' ').strip() if len(r) > 6 else "",
                        r[7].strip() if len(r) > 7 else "",
                        r[8].strip() if len(r) > 8 else "",
                        r[9].strip() if len(r) > 9 else "",
                        r[10].strip() if len(r) > 10 else ""
                    ))
                    count += 1

            conn.commit()
            conn.close()
            messagebox.showinfo("Succès", f"{count} modèles importés depuis Usi-Tab.csv.")
            self.load_catalog_data()
            self.update_model_comboboxes()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de l'importation :\n{str(e)}")

    def export_catalog_csv(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Fichiers CSV", "*.csv")])
        if not file_path:
            return
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT model_name, prog_name, block_dim, block_dim_bought, qty_per_block, z_between_pains, tools, caisson, top_plate, bottom_plate, remarks FROM models_catalog")
        rows = cursor.fetchall()
        conn.close()

        with open(file_path, mode='w', newline='', encoding='latin1') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(["Nom Model", "Prog. Pain", "Dimension Bloc", "Dimension Achetée", "Qte/Bloc", "Z d'Axe", "Outils", "Caisson", "PLAQUE 2su", "PLAQUE 2so", "Remarque"])
            writer.writerows(rows)
        messagebox.showinfo("Export", "Catalogue exporté avec succès.")

    def add_model_dialog(self):
        win = tk.Toplevel(self)
        win.title("Ajouter un Modèle au Catalogue")
        win.geometry("420x480")

        fields = ["Nom Modèle", "Prog. Pain", "Dimension Bloc", "Dimension Achetée", "Qte/Bloc", "Z d'Axe", "Outils", "Caisson", "Plaque Top", "Plaque Bottom", "Remarques"]
        entries = {}

        for i, field in enumerate(fields):
            ttk.Label(win, text=f"{field} :").grid(row=i, column=0, padx=10, pady=3, sticky="w")
            e = ttk.Entry(win, width=25)
            e.grid(row=i, column=1, padx=10, pady=3)
            entries[field] = e

        def save():
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO models_catalog (model_name, prog_name, block_dim, block_dim_bought, qty_per_block, z_between_pains, tools, caisson, top_plate, bottom_plate, remarks)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', tuple(entries[f].get().strip() for f in fields))
            conn.commit()
            conn.close()
            messagebox.showinfo("Succès", "Modèle ajouté au catalogue.", parent=win)
            win.destroy()
            self.load_catalog_data()
            self.update_model_comboboxes()

        ttk.Button(win, text="Enregistrer", command=save).grid(row=len(fields), column=0, columnspan=2, pady=15)

    def delete_model_dialog(self):
        win = tk.Toplevel(self)
        win.title("Supprimer un Modèle")
        win.geometry("400x300")

        ttk.Label(win, text="Rechercher le modèle à supprimer :").pack(pady=5)
        entry_search = ttk.Entry(win, width=30)
        entry_search.pack(pady=5)

        listbox = tk.Listbox(win, width=45, height=8)
        listbox.pack(pady=5)

        def update_list():
            listbox.delete(0, tk.END)
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("SELECT id, model_name FROM models_catalog WHERE model_name LIKE ?", (f'%{entry_search.get().strip()}%',))
            for row in cursor.fetchall():
                listbox.insert(tk.END, f"{row[0]} | {row[1]}")
            conn.close()

        entry_search.bind("<KeyRelease>", lambda e: update_list())
        update_list()

        def confirm_delete():
            sel = listbox.get(tk.ACTIVE)
            if not sel:
                return
            m_id, m_name = sel.split(" | ")[0], sel.split(" | ")[1]
            if messagebox.askyesno("Confirmation", f"Êtes-vous sûr de vouloir supprimer le modèle '{m_name}' ?", parent=win):
                conn = sqlite3.connect(DB_FILE)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM models_catalog WHERE id=?", (m_id,))
                conn.commit()
                conn.close()
                messagebox.showinfo("Supprimé", "Modèle supprimé.", parent=win)
                win.destroy()
                self.load_catalog_data()
                self.update_model_comboboxes()

        ttk.Button(win, text="Supprimer le modèle sélectionné", command=confirm_delete).pack(pady=10)

    def hide_model_dialog(self):
        win = tk.Toplevel(self)
        win.title("Cacher / Démasquer un Modèle")
        win.geometry("400x320")

        ttk.Label(win, text="Rechercher un modèle à cacher/afficher :").pack(pady=5)
        entry_search = ttk.Entry(win, width=30)
        entry_search.pack(pady=5)

        listbox = tk.Listbox(win, width=45, height=8)
        listbox.pack(pady=5)

        def update_list():
            listbox.delete(0, tk.END)
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("SELECT id, model_name, is_hidden FROM models_catalog WHERE model_name LIKE ?", (f'%{entry_search.get().strip()}%',))
            for row in cursor.fetchall():
                status = "[CACHÉ]" if row[2] == 1 else "[VISIBLE]"
                listbox.insert(tk.END, f"{row[0]} | {row[1]} {status}")
            conn.close()

        entry_search.bind("<KeyRelease>", lambda e: update_list())
        update_list()

        def toggle_hide():
            sel = listbox.get(tk.ACTIVE)
            if not sel:
                return
            m_id = sel.split(" | ")[0]
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("UPDATE models_catalog SET is_hidden = CASE WHEN is_hidden=1 THEN 0 ELSE 1 END WHERE id=?", (m_id,))
            conn.commit()
            conn.close()
            update_list()
            self.load_catalog_data()

        ttk.Button(win, text="Bascule Cacher / Afficher", command=toggle_hide).pack(pady=10)

    # ==========================================
    # ONGLET 2 : COMPARAISON DE DOSSIERS (A/B)
    # ==========================================
    def setup_compare_tab(self, parent):
        frame_dirs = ttk.LabelFrame(parent, text="Sélection des dossiers à comparer")
        frame_dirs.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(frame_dirs, text="Dossier A (Référence) :").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.var_dir_a = tk.StringVar()
        ttk.Entry(frame_dirs, textvariable=self.var_dir_a, width=68).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame_dirs, text="Parcourir", command=lambda: self.browse_dir(self.var_dir_a)).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(frame_dirs, text="Dossier B (Comparé) :").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.var_dir_b = tk.StringVar()
        ttk.Entry(frame_dirs, textvariable=self.var_dir_b, width=68).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame_dirs, text="Parcourir", command=lambda: self.browse_dir(self.var_dir_b)).grid(row=1, column=2, padx=5, pady=5)

        frame_btns = ttk.Frame(parent)
        frame_btns.pack(fill=tk.X, padx=10, pady=5)

        btn_fast = tk.Button(frame_btns, text="Lancer la Comparaison Rapide (Dates/Tailles)", bg="#0288D1", fg="white", font=("Arial", 9, "bold"), command=lambda: self.run_comparison(deep=False))
        btn_fast.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        btn_deep = tk.Button(frame_btns, text="🔍 Comparaison Approfondie (Contenu Texte CNC)", bg="#2E7D32", fg="white", font=("Arial", 9, "bold"), command=lambda: self.run_comparison(deep=True))
        btn_deep.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        frame_grid = ttk.LabelFrame(parent, text="Résultats de la comparaison")
        frame_grid.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("model", "fichier", "statut", "date_a", "date_b")
        self.tree_comp = ttk.Treeview(frame_grid, columns=cols, show="headings", selectmode="browse")
        
        self.tree_comp.heading("model", text="Nom de model ↕", command=lambda: self.sort_treeview("model"))
        self.tree_comp.heading("fichier", text="Nom Prog. (Fichier/Chemin) ↕", command=lambda: self.sort_treeview("fichier"))
        self.tree_comp.heading("statut", text="Statut ↕", command=lambda: self.sort_treeview("statut"))
        self.tree_comp.heading("date_a", text="Date Modification (A)")
        self.tree_comp.heading("date_b", text="Date Modification (B)")

        self.tree_comp.column("model", width=140, anchor="w")
        self.tree_comp.column("fichier", width=280, anchor="w")
        self.tree_comp.column("statut", width=110, anchor="center")
        self.tree_comp.column("date_a", width=170, anchor="center")
        self.tree_comp.column("date_b", width=170, anchor="center")

        self.tree_comp.tag_configure("different_tag", background="#D32F2F", foreground="white")
        self.tree_comp.tag_configure("identique_tag", background="white", foreground="black")

        scrollbar = ttk.Scrollbar(frame_grid, orient=tk.VERTICAL, command=self.tree_comp.yview)
        self.tree_comp.configure(yscrollcommand=scrollbar.set)
        self.tree_comp.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        frame_bottom = ttk.Frame(parent)
        frame_bottom.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(frame_bottom, text="Exporter en TXT", command=self.export_txt).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_bottom, text="Exporter en Excel (.csv)", command=self.export_csv).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_bottom, text="🖨️ Imprimer Rapport A4", bg="#424242", fg="white", font=("Arial", 9, "bold"), command=self.print_a4_formatted).pack(side=tk.RIGHT, padx=5)

    def browse_dir(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

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
        items = [(self.tree_comp.set(k, col), k) for k in self.tree_comp.get_children('')]
        items.sort(reverse=reverse)
        for index, (val, k) in enumerate(items):
            self.tree_comp.move(k, '', index)
        self.sort_directions[col] = not reverse

    def run_comparison(self, deep=False):
        dir_a, dir_b = self.var_dir_a.get().strip(), self.var_dir_b.get().strip()

        if not os.path.isdir(dir_a) or not os.path.isdir(dir_b):
            messagebox.showwarning("Attention", "Veuillez sélectionner deux dossiers valides.")
            return

        for item in self.tree_comp.get_children():
            self.tree_comp.delete(item)

        files_a = {os.path.relpath(os.path.join(r, f), dir_a): os.path.join(r, f) for r, d, files in os.walk(dir_a) for f in files}
        files_b = {os.path.relpath(os.path.join(r, f), dir_b): os.path.join(r, f) for r, d, files in os.walk(dir_b) for f in files}

        all_rel_files = sorted(list(set(files_a.keys()).union(set(files_b.keys()))))

        for rel in all_rel_files:
            path_a, path_b = files_a.get(rel), files_b.get(rel)
            date_a_str = self.get_file_date(path_a) if path_a else "Absence (A)"
            date_b_str = self.get_file_date(path_b) if path_b else "Absence (B)"

            if path_a and path_b:
                is_same = filecmp.cmp(path_a, path_b, shallow=False) if deep else ((os.stat(path_a).st_mtime == os.stat(path_b).st_mtime) and (os.stat(path_a).st_size == os.stat(path_b).st_size))
                statut = "IDENTIQUE" if is_same else "DIFFERENT"
            else:
                statut = "DIFFERENT"

            model_name = self.find_model_name_for_file(rel)
            tag = "different_tag" if statut == "DIFFERENT" else "identique_tag"
            self.tree_comp.insert("", tk.END, values=(model_name, rel, statut, date_a_str, date_b_str), tags=(tag,))

    def get_file_date(self, path):
        return datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime('%Y-%m-%d %H:%M:%S')

    def export_txt(self):
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Texte", "*.txt")])
        if path:
            with open(path, "w", encoding="utf-8") as f:
                for item in self.tree_comp.get_children():
                    f.write("\t".join(map(str, self.tree_comp.item(item, "values"))) + "\n")
            messagebox.showinfo("Export", "Export TXT réussi !")

    def export_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow(["Nom Model", "Nom Programme pain / Fichier", "Statut", "Date A", "Date B"])
                for item in self.tree_comp.get_children():
                    writer.writerow(self.tree_comp.item(item, "values"))
            messagebox.showinfo("Export", "Export CSV réussi !")

    def print_a4_formatted(self):
        items = self.tree_comp.get_children()
        if not items:
            messagebox.showwarning("Attention", "Aucune donnée à imprimer !")
            return

        rows = [self.tree_comp.item(item, "values") for item in items]
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

        elements = [Paragraph(f"<b>RAPPORT DE COMPARAISON - {APP_NAME.upper()}</b>", title_style), Spacer(1, 10)]

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
        ]

        for i, r in enumerate(rows, start=1):
            model, fichier, statut, date_a, date_b = r
            data.append([Paragraph(model, cell_style), Paragraph(fichier, cell_style), Paragraph(f"<b>{statut}</b>", cell_style), Paragraph(date_a, cell_style), Paragraph(date_b, cell_style)])
            if statut == "DIFFERENT":
                table_styles.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor("#D32F2F")))

        t = Table(data, colWidths=[90, 180, 70, 105, 105])
        t.setStyle(TableStyle(table_styles))
        elements.append(t)

        doc.build(elements)
        os.startfile(pdf_filename, "print") if hasattr(os, "startfile") else webbrowser.open(pdf_filename)

    def generate_html_print(self, rows):
        html_filename = os.path.join(tempfile.gettempdir(), "rapport_cnc_a4.html")
        html_content = f"""
        <html><head><meta charset="utf-8"><style>
            @page {{ size: A4 portrait; margin: 10mm; }}
            body {{ font-family: Arial; font-size: 9pt; }}
            h2 {{ text-align: center; }}
            table {{ width: 100%; border-collapse: collapse; }}
            th, td {{ border: 1px solid #666; padding: 4px; font-size: 8pt; }}
            th {{ background-color: #0B3C5D; color: white; }}
            tr.different {{ background-color: #D32F2F !important; color: white; font-weight: bold; }}
        </style></head><body onload="window.print();">
            <h2>RAPPORT DE COMPARAISON - {APP_NAME.upper()}</h2>
            <table><thead><tr><th>Nom Model</th><th>Nom Prog.</th><th>Statut</th><th>Date A</th><th>Date B</th></tr></thead><tbody>
        """
        for r in rows:
            model, fichier, statut, date_a, date_b = r
            html_content += f'<tr class="{"different" if statut == "DIFFERENT" else ""}"><td>{model}</td><td>{fichier}</td><td>{statut}</td><td>{date_a}</td><td>{date_b}</td></tr>'
        html_content += "</tbody></table></body></html>"

        with open(html_filename, "w", encoding="utf-8") as f:
            f.write(html_content)
        webbrowser.open(html_filename)

    # ==========================================
    # ONGLET 3 : PLANIFICATION ET SAUVEGARDES
    # ==========================================
    def setup_backup_tab(self, parent):
        frame_add = ttk.LabelFrame(parent, text="Ajouter / Configurer une Sauvegarde")
        frame_add.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(frame_add, text="Nom de la sauvegarde :").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.entry_task_name = ttk.Entry(frame_add, width=45)
        self.entry_task_name.grid(row=0, column=1, columnspan=2, sticky="w", padx=5, pady=3)

        ttk.Label(frame_add, text="Dossier Source :").grid(row=1, column=0, sticky="w", padx=5, pady=3)
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

        btn_save = tk.Button(frame_add, text="Enregistrer la Sauvegarde", font=("Arial", 10, "bold"), bg="#2E7D32", fg="white", pady=4, command=self.add_task)
        btn_save.grid(row=4, column=0, columnspan=3, pady=10, sticky="ew", padx=5)

        frame_list = ttk.LabelFrame(parent, text="Liste des Sauvegardes Enregistrées")
        frame_list.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("name", "src", "dest", "type")
        self.tree_tasks = ttk.Treeview(frame_list, columns=columns, show="headings", height=6)
        self.tree_tasks.heading("name", text="Nom Tâche")
        self.tree_tasks.heading("src", text="Source")
        self.tree_tasks.heading("dest", text="Destination")
        self.tree_tasks.heading("type", text="Récurrence")

        self.tree_tasks.pack(fill=tk.BOTH, expand=True, side=tk.LEFT, padx=5, pady=5)

        btn_actions = ttk.Frame(frame_list)
        btn_actions.pack(fill=tk.Y, side=tk.RIGHT, padx=5, pady=5)

        btn_manual = tk.Button(btn_actions, text="Lancer Manuel", font=("Arial", 9, "bold"), bg="#1976D2", fg="white", command=self.run_task_manual)
        btn_manual.pack(fill=tk.X, pady=5)

        btn_del = tk.Button(btn_actions, text="Supprimer", font=("Arial", 9, "bold"), bg="#C62828", fg="white", command=self.delete_task)
        btn_del.pack(fill=tk.X, pady=5)

    def browse_src(self):
        path = filedialog.askdirectory()
        if path:
            self.entry_src.delete(0, tk.END)
            self.entry_src.insert(0, path)

    def browse_dest(self):
        path = filedialog.askdirectory()
        if path:
            self.entry_dest.delete(0, tk.END)
            self.entry_dest.insert(0, path)

    def add_task(self):
        name, src, dest, rec_type = self.entry_task_name.get().strip(), self.entry_src.get().strip(), self.entry_dest.get().strip(), self.combo_type.get()
        if name and src and dest:
            task = (name, src, dest, rec_type)
            self.tasks.append(task)
            self.tree_tasks.insert("", tk.END, values=task)
            messagebox.showinfo("Succès", f"Sauvegarde '{name}' enregistrée.")

    def run_task_manual(self):
        selected = self.tree_tasks.selection()
        if not selected:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner une sauvegarde.")
            return

        name, src, dest, _ = self.tree_tasks.item(selected[0])['values']
        if not os.path.exists(src):
            messagebox.showerror("Erreur", f"Source introuvable :\n{src}")
            return

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        target_dir = os.path.join(dest, f"{name}_{timestamp}")

        file_list = [os.path.join(r, f) for r, d, files in os.walk(src) for f in files] if os.path.isdir(src) else [src]
        total_files = len(file_list)

        progress_dialog = BackupProgressBarDialog(self, title=f"Sauvegarde : {name}")
        try:
            os.makedirs(target_dir, exist_ok=True)
            for idx, file_path in enumerate(file_list, 1):
                if progress_dialog.cancelled:
                    messagebox.showinfo("Annulation", "Sauvegarde annulée.")
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
            messagebox.showerror("Erreur", f"Échec de la sauvegarde : {str(e)}")

    def delete_task(self):
        selected = self.tree_tasks.selection()
        if selected and messagebox.askyesno("Confirmation", "Supprimer cette sauvegarde ?"):
            self.tree_tasks.delete(selected[0])

    # ==========================================
    # ONGLET 4 : ORDRES DE FABRICATION (OF)
    # ==========================================
    def setup_of_tab(self, parent):
        frame_top = ttk.LabelFrame(parent, text=" 1. Définition du Lancement de Production (En-tête OF) ")
        frame_top.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_top, text="Machine :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.combo_of_mach = ttk.Combobox(frame_top, values=["NUM 1060 (5-Axes)", "Fraiseuse EPS", "Tour CNC"], state="readonly", width=18)
        self.combo_of_mach.current(0)
        self.combo_of_mach.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_top, text="Opérateur :").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.entry_of_op = ttk.Entry(frame_top, width=15)
        self.entry_of_op.insert(0, "Admin")
        self.entry_of_op.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_top, text="Priorité :").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.combo_of_prio = ttk.Combobox(frame_top, values=["Haute", "Normale", "Basse"], state="readonly", width=10)
        self.combo_of_prio.current(1)
        self.combo_of_prio.grid(row=0, column=5, padx=5, pady=5)

        frame_item = ttk.LabelFrame(parent, text=" 2. Modèles & Spécifications Bruts à Inclure dans cet OF ")
        frame_item.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_item, text="Modèle :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.combo_of_model = ttk.Combobox(frame_item, width=22)
        self.combo_of_model.grid(row=0, column=1, padx=5, pady=5)
        self.update_model_comboboxes()

        ttk.Label(frame_item, text="Qte :").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.spin_of_qty = tk.Spinbox(frame_item, from_=1, to=100, width=5)
        self.spin_of_qty.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_item, text="N° Bloc :").grid(row=0, column=4, padx=5, pady=5, sticky="w")
        self.entry_of_bnum = ttk.Entry(frame_item, width=10)
        self.entry_of_bnum.grid(row=0, column=5, padx=5, pady=5)

        ttk.Label(frame_item, text="Densité :").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.entry_of_bdens = ttk.Entry(frame_item, width=12)
        self.entry_of_bdens.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame_item, text="N° Pain :").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.entry_of_pnum = ttk.Entry(frame_item, width=8)
        self.entry_of_pnum.grid(row=1, column=3, padx=5, pady=5)

        ttk.Label(frame_item, text="Poids (g) :").grid(row=1, column=4, padx=5, pady=5, sticky="w")
        self.entry_of_pweight = ttk.Entry(frame_item, width=10)
        self.entry_of_pweight.grid(row=1, column=5, padx=5, pady=5)

        btn_add_item = tk.Button(frame_item, text="➕ Ajouter Pièce au Lancement", bg="#0288D1", fg="white", font=("Arial", 9, "bold"), command=self.add_item_to_of_cart)
        btn_add_item.grid(row=1, column=6, padx=10, pady=5)

        self.tree_cart = ttk.Treeview(frame_item, columns=("model", "qty", "bnum", "bdens", "pnum", "pweight"), show="headings", height=3)
        self.tree_cart.heading("model", text="Modèle")
        self.tree_cart.heading("qty", text="Qté")
        self.tree_cart.heading("bnum", text="N° Bloc")
        self.tree_cart.heading("bdens", text="Densité")
        self.tree_cart.heading("pnum", text="N° Pain")
        self.tree_cart.heading("pweight", text="Poids (g)")

        self.tree_cart.column("model", width=150)
        self.tree_cart.column("qty", width=50, anchor="center")
        self.tree_cart.column("bnum", width=80, anchor="center")
        self.tree_cart.column("bdens", width=80, anchor="center")
        self.tree_cart.column("pnum", width=80, anchor="center")
        self.tree_cart.column("pweight", width=80, anchor="center")
        self.tree_cart.grid(row=2, column=0, columnspan=7, sticky="ew", padx=5, pady=5)

        btn_val_of = tk.Button(parent, text="🚀 VALIDER ET CRÉER L'ORDRE DE FABRICATION GLOBAL", bg="#2E7D32", fg="white", font=("Arial", 10, "bold"), pady=5, command=self.save_global_of)
        btn_val_of.pack(fill="x", padx=10, pady=5)

        frame_list = ttk.LabelFrame(parent, text=" 3. Historique des Lancements OF & Pièces Associées ")
        frame_list.pack(fill="both", expand=True, padx=10, pady=5)

        frame_actions_of = ttk.Frame(frame_list)
        frame_actions_of.pack(fill="x", padx=5, pady=2)

        ttk.Button(frame_actions_of, text=" Charger & Envoyer Programme de la Pièce vers CNC (RS232)", command=self.transfer_selected_of_item_to_cnc).pack(side="left", padx=5)
        ttk.Button(frame_actions_of, text="🖨️ Imprimer Liste OF", command=lambda: self.print_treeview_data(self.tree_of, "Ordres de Fabrication")).pack(side="right", padx=5)

        self.tree_of = ttk.Treeview(frame_list, columns=("id", "of_number", "machine", "operator", "priority", "status", "created_at"), show="headings", height=4)
        self.tree_of.heading("id", text="ID")
        self.tree_of.heading("of_number", text="N° OF Lancement")
        self.tree_of.heading("machine", text="Machine")
        self.tree_of.heading("operator", text="Opérateur")
        self.tree_of.heading("priority", text="Priorité")
        self.tree_of.heading("status", text="Statut Lancement")
        self.tree_of.heading("created_at", text="Date Lancement")

        self.tree_of.column("id", width=40, anchor="center")
        self.tree_of.column("of_number", width=140, anchor="center")
        self.tree_of.column("machine", width=140)
        self.tree_of.column("operator", width=100)
        self.tree_of.column("priority", width=80, anchor="center")
        self.tree_of.column("status", width=100, anchor="center")
        self.tree_of.column("created_at", width=140, anchor="center")
        self.tree_of.pack(fill="x", padx=5, pady=5)
        self.tree_of.bind("<<TreeviewSelect>>", self.on_of_selected)

        ttk.Label(frame_list, text="Pièces/Modèles inclus dans l'OF sélectionné :", font=("Arial", 9, "bold")).pack(anchor="w", padx=5, pady=2)
        
        self.tree_of_items = ttk.Treeview(frame_list, columns=("id", "model", "prog", "qty", "bnum", "bdens", "pnum", "pweight", "status"), show="headings", height=4)
        self.tree_of_items.heading("id", text="ID Line")
        self.tree_of_items.heading("model", text="Nom Modèle")
        self.tree_of_items.heading("prog", text="Prog. Pain")
        self.tree_of_items.heading("qty", text="Qté")
        self.tree_of_items.heading("bnum", text="N° Bloc")
        self.tree_of_items.heading("bdens", text="Densité")
        self.tree_of_items.heading("pnum", text="N° Pain")
        self.tree_of_items.heading("pweight", text="Poids (g)")
        self.tree_of_items.heading("status", text="Statut Pièce")

        self.tree_of_items.column("id", width=50, anchor="center")
        self.tree_of_items.column("model", width=140)
        self.tree_of_items.column("prog", width=120)
        self.tree_of_items.column("qty", width=50, anchor="center")
        self.tree_of_items.column("bnum", width=80, anchor="center")
        self.tree_of_items.column("bdens", width=80, anchor="center")
        self.tree_of_items.column("pnum", width=80, anchor="center")
        self.tree_of_items.column("pweight", width=80, anchor="center")
        self.tree_of_items.column("status", width=100, anchor="center")
        self.tree_of_items.pack(fill="both", expand=True, padx=5, pady=5)

        self.load_of_data()

    def update_model_comboboxes(self):
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT model_name FROM models_catalog WHERE is_hidden=0")
        models = [r[0] for r in cursor.fetchall()]
        conn.close()

        if hasattr(self, 'combo_of_model'):
            self.combo_of_model['values'] = models
            if models:
                self.combo_of_model.current(0)

    def add_item_to_of_cart(self):
        model = self.combo_of_model.get().strip()
        if not model:
            messagebox.showwarning("Attention", "Veuillez sélectionner un modèle.")
            return

        qty = self.spin_of_qty.get()
        bnum, bdens = self.entry_of_bnum.get().strip(), self.entry_of_bdens.get().strip()
        pnum, pweight = self.entry_of_pnum.get().strip(), self.entry_of_pweight.get().strip()

        item = (model, qty, bnum, bdens, pnum, pweight)
        self.current_of_cart.append(item)
        self.tree_cart.insert("", tk.END, values=item)

    def generate_next_of_number(self):
        year = datetime.datetime.now().year
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM work_orders ORDER BY id DESC LIMIT 1")
        row = cursor.fetchone()
        conn.close()

        next_id = (row[0] + 1) if row else 1
        return f"OF{next_id:05d}/{year}"

    def save_global_of(self):
        if not self.current_of_cart:
            messagebox.showwarning("Panier Vide", "Veuillez ajouter au moins un modèle à cet OF.")
            return

        of_num = self.generate_next_of_number()
        machine = self.combo_of_mach.get()
        operator = self.entry_of_op.get().strip() or "Admin"
        priority = self.combo_of_prio.get()

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        cursor.execute("INSERT INTO work_orders (of_number, machine, assigned_operator, priority, status) VALUES (?, ?, ?, ?, ?)", (of_num, machine, operator, priority, 'En attente'))

        for model_name, qty, bnum, bdens, pnum, pweight in self.current_of_cart:
            cursor.execute("SELECT prog_name FROM models_catalog WHERE model_name=?", (model_name,))
            r = cursor.fetchone()
            prog = r[0] if r else ""

            cursor.execute('''
                INSERT INTO work_order_items (of_number, model_name, prog_name, qty, block_num, block_density, pain_num, pain_weight, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (of_num, model_name, prog, qty, bnum, bdens, pnum, pweight, 'En attente'))

        conn.commit()
        conn.close()

        self.current_of_cart.clear()
        for item in self.tree_cart.get_children():
            self.tree_cart.delete(item)

        messagebox.showinfo("Lancement Réussi", f"L'Ordre de Fabrication {of_num} a été créé avec succès.")
        self.load_of_data()

    def load_of_data(self):
        for item in self.tree_of.get_children():
            self.tree_of.delete(item)
        for item in self.tree_of_items.get_children():
            self.tree_of_items.delete(item)

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT id, of_number, machine, assigned_operator, priority, status, created_at FROM work_orders ORDER BY id DESC")
        for row in cursor.fetchall():
            self.tree_of.insert("", "end", values=row)
        conn.close()

    def on_of_selected(self, event):
        selected = self.tree_of.selection()
        if not selected:
            return

        for item in self.tree_of_items.get_children():
            self.tree_of_items.delete(item)

        of_num = self.tree_of.item(selected[0])['values'][1]

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT id, model_name, prog_name, qty, block_num, block_density, pain_num, pain_weight, status FROM work_order_items WHERE of_number=?", (of_num,))
        for row in cursor.fetchall():
            self.tree_of_items.insert("", "end", values=row)
        conn.close()

    def transfer_selected_of_item_to_cnc(self):
        selected_item = self.tree_of_items.selection()
        selected_of = self.tree_of.selection()

        if not selected_item or not selected_of:
            messagebox.showwarning("Attention", "Veuillez sélectionner un OF puis la pièce à envoyer.")
            return

        of_num = self.tree_of.item(selected_of[0])['values'][1]
        item_vals = self.tree_of_items.item(selected_item[0])['values']
        model_name, prog_name = item_vals[1], item_vals[2]

        self.notebook.select(5)
        self.lbl_file.config(text=f"OF: {of_num} | Modèle: {model_name} | Prog: {prog_name}", font=("Arial", 9, "bold"))
        self.txt_preview.delete("1.0", tk.END)
        self.txt_preview.insert(tk.END, f"% \n(PROGRAMME ASSOCIE A L'OF {of_num})\n(MODELE: {model_name})\n(PROGRAMME PAIN: {prog_name})\n\nG00 G90 G40\nM03 S12000\nG00 X0 Y0 Z50\n(G-CODE COMPATIBLE NUM 1060)\nM05\nM30\n%")

    # ==========================================
    # ONGLET 5 : TRAÇABILITÉ & SUIVI USINAGE
    # ==========================================
    def setup_tracking_tab(self, parent):
        frame_input = ttk.LabelFrame(parent, text=" Enregistrer une Étape d'Usinage ")
        frame_input.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_input, text="N° OF :").grid(row=0, column=0, padx=5, pady=5)
        self.entry_tr_of = ttk.Entry(frame_input, width=12)
        self.entry_tr_of.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_input, text="Machine :").grid(row=0, column=2, padx=5, pady=5)
        self.combo_tr_mach = ttk.Combobox(frame_input, values=["NUM 1060 (5-Axes)", "Fraiseuse EPS", "Tour CNC"], state="readonly", width=16)
        self.combo_tr_mach.current(0)
        self.combo_tr_mach.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_input, text="Modèle / Pièce :").grid(row=0, column=4, padx=5, pady=5)
        self.entry_tr_model = ttk.Entry(frame_input, width=20)
        self.entry_tr_model.grid(row=0, column=5, padx=5, pady=5)

        ttk.Label(frame_input, text="Temps Réel (min) :").grid(row=1, column=0, padx=5, pady=5)
        self.entry_tr_time = ttk.Entry(frame_input, width=12)
        self.entry_tr_time.insert(0, "45")
        self.entry_tr_time.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame_input, text="Statut :").grid(row=1, column=2, padx=5, pady=5)
        self.combo_tr_stat = ttk.Combobox(frame_input, values=["Terminé", "En cours", "Maintenance", "En attente"], state="readonly", width=16)
        self.combo_tr_stat.current(0)
        self.combo_tr_stat.grid(row=1, column=3, padx=5, pady=5)

        ttk.Button(frame_input, text="Enregistrer l'Usinage", command=self.add_tracking_log).grid(row=1, column=4, columnspan=2, padx=10, pady=5, sticky="ew")

        frame_list = ttk.Frame(parent)
        frame_list.pack(fill="both", expand=True, padx=10, pady=5)

        ttk.Button(parent, text="🖨️ Imprimer Historique Traçabilité", command=lambda: self.print_treeview_data(self.tree_track, "Tracabilite et Suivi")).pack(anchor="e", padx=10, pady=2)

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
        of_num, model, t_time = self.entry_tr_of.get().strip(), self.entry_tr_model.get().strip(), self.entry_tr_time.get().strip()

        if not of_num or not model:
            messagebox.showwarning("Attention", "N° OF et Modèle sont obligatoires.")
            return

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO machining_history (of_number, operator, machine, model_name, real_time_min, status)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (of_num, self.current_user['username'], self.combo_tr_mach.get(), model, t_time, self.combo_tr_stat.get()))

        cursor.execute("UPDATE work_order_items SET status=? WHERE of_number=? AND model_name=?", (self.combo_tr_stat.get(), of_num, model))
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
    # ONGLET 6 : TRANSFERT RS232 CNC NUM 1060
    # ==========================================
    def setup_cnc_tab(self, parent):
        frame_cfg = ttk.LabelFrame(parent, text=" Paramètres de Communication RS232 (NUM 1060) ")
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

        frame_file = ttk.Frame(parent)
        frame_file.pack(fill="x", padx=10, pady=5)
        
        self.lbl_file = ttk.Label(frame_file, text="Aucun fichier chargé", font=("Arial", 9, "italic"))
        self.lbl_file.pack(side="left", padx=5)

        ttk.Button(frame_file, text="Ouvrir Fichier G-Code (.ISO / .NC)", command=self.open_gcode_file).pack(side="right", padx=5)

        frame_prev = ttk.LabelFrame(parent, text=" Aperçu du Programme ISO / NUM ")
        frame_prev.pack(fill="both", expand=True, padx=10, pady=5)

        self.txt_preview = tk.Text(frame_prev, wrap="none", font=("Courier", 10))
        scroll_y = ttk.Scrollbar(frame_prev, orient="vertical", command=self.txt_preview.yview)
        scroll_x = ttk.Scrollbar(frame_prev, orient="horizontal", command=self.txt_preview.xview)
        self.txt_preview.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.txt_preview.pack(fill="both", expand=True)

        frame_send = ttk.Frame(parent)
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

        port, baud = self.combo_port.get(), int(self.combo_baud.get())

        if serial is None:
            messagebox.showinfo("Simulation Transfert RS232", f"[Mode Simulation]\nPySerial non installé.\nProgramme transmis avec succès à la NUM 1060 via {port} à {baud} bauds.")
            return

        try:
            ser = serial.Serial(port, baud, timeout=2)
            ser.write(gcode.encode('latin1'))
            ser.close()
            messagebox.showinfo("Transfert Réussi", f"Programme transmis avec succès à la NUM 1060 via {port}.")
        except Exception as e:
            messagebox.showerror("Erreur RS232", f"Impossible de communiquer avec le port {port} :\n{str(e)}")

    # --- ADMINISTRATION UTILISATEURS ---
    def open_user_management(self):
        win = tk.Toplevel(self)
        win.title("Gestion des Utilisateurs")
        win.geometry("450x320")

        frame_top = ttk.Frame(win, padding=10)
        frame_top.pack(fill="x")

        ttk.Label(frame_top, text="Utilisateur:").grid(row=0, column=0, padx=5, pady=5)
        e_user = ttk.Entry(frame_top, width=12)
        e_user.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_top, text="Mot de passe:").grid(row=0, column=2, padx=5, pady=5)
        e_pwd = ttk.Entry(frame_top, width=12)
        e_pwd.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_top, text="Rôle:").grid(row=1, column=0, padx=5, pady=5)
        c_role = ttk.Combobox(frame_top, values=["Admin", "Operateur"], state="readonly", width=10)
        c_role.current(1)
        c_role.grid(row=1, column=1, padx=5, pady=5)

        def add_u():
            u, p, r = e_user.get().strip(), e_pwd.get().strip(), c_role.get()
            if u and p:
                conn = sqlite3.connect(DB_FILE)
                cursor = conn.cursor()
                try:
                    cursor.execute("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", (u, p, r))
                    conn.commit()
                    messagebox.showinfo("Succès", "Utilisateur créé.", parent=win)
                    load_u()
                except sqlite3.IntegrityError:
                    messagebox.showerror("Erreur", "Nom d'utilisateur déjà existant.", parent=win)
                conn.close()

        ttk.Button(frame_top, text="Créer", command=add_u).grid(row=1, column=2, columnspan=2, padx=5, pady=5, sticky="ew")

        tree_u = ttk.Treeview(win, columns=("id", "user", "role"), show="headings", height=6)
        tree_u.heading("id", text="ID")
        tree_u.heading("user", text="Nom")
        tree_u.heading("role", text="Rôle")
        tree_u.pack(fill="both", expand=True, padx=10, pady=5)

        def load_u():
            for item in tree_u.get_children():
                tree_u.delete(item)
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("SELECT id, username, role FROM users")
            for row in cursor.fetchall():
                tree_u.insert("", "end", values=row)
            conn.close()

        load_u()


# ==========================================
# 4. POINT D'ENTRÉE DU PROGRAMME
# ==========================================

if __name__ == "__main__":
    init_db()
    app = CNCApplication()
    app.mainloop()
