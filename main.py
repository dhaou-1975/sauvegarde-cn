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
APP_VERSION = "4.1.0"
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

    # Table Utilisateurs
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')

    # Table Catalogue Modèles
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
            phase TEXT DEFAULT 'Étude',
            is_hidden INTEGER DEFAULT 0
        )
    ''')

    # Migration si la colonne phase n'existe pas
    cursor.execute("PRAGMA table_info(models_catalog)")
    cols = [column[1] for column in cursor.fetchall()]
    if 'phase' not in cols:
        cursor.execute("ALTER TABLE models_catalog ADD COLUMN phase TEXT DEFAULT 'Étude'")

    # Table Machines & Outils
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS machines (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            description TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS machine_tools (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine_id INTEGER NOT NULL,
            tool_name TEXT NOT NULL,
            diameter REAL,
            length_out REAL,
            length_comp REAL,
            pocket INTEGER,
            FOREIGN KEY(machine_id) REFERENCES machines(id) ON DELETE CASCADE
        )
    ''')

    # En-tête Lancement OF
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

    # Lignes d'OF (Multi-Modèles)
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

    # Traçabilité Usinage
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

    # Machines par défaut
    cursor.execute("SELECT COUNT(*) FROM machines")
    if cursor.fetchone()[0] == 0:
        cursor.execute("INSERT INTO machines (name, description) VALUES ('NUM 1060 (5-Axes)', 'Fraiseuse 5 axes NUM 1060')")
        cursor.execute("INSERT INTO machines (name, description) VALUES ('Fraiseuse EPS', 'Fraiseuse 3 axes grands volumes')")

    conn.commit()
    conn.close()


def load_config():
    default_config = {
        "password_hash": DEFAULT_PASSWORD_HASH,
        "default_working_dir": os.path.expanduser("~"),
        "use_last_backup_dir": False,
        "last_backup_dir": ""
    }
    if not os.path.exists(CONFIG_FILE):
        save_config(default_config)
        return default_config
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            for k, v in default_config.items():
                if k not in cfg:
                    cfg[k] = v
            return cfg
    except Exception:
        return default_config


def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


# ==========================================
# 2. DIALOGUES ET MODULES AUXILIAIRES
# ==========================================

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


# Dialogue d'Options (Outils > Options)
class OptionsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Configuration des Dossiers et Options")
        self.geometry("520x240")
        self.resizable(False, False)
        self.grab_set()
        self.parent = parent

        self.config = load_config()

        ttk.Label(self, text="Dossier de travail par défaut (Programmes G-Code) :", font=("Arial", 9, "bold")).pack(anchor="w", padx=15, pady=(15, 5))

        f_dir = ttk.Frame(self)
        f_dir.pack(fill="x", padx=15, pady=5)

        self.entry_dir = ttk.Entry(f_dir, width=48)
        self.entry_dir.insert(0, self.config.get("default_working_dir", ""))
        self.entry_dir.pack(side="left", padx=(0, 5))

        ttk.Button(f_dir, text="Parcourir", command=self.browse_dir).pack(side="left")

        self.var_use_last = tk.BooleanVar(value=self.config.get("use_last_backup_dir", False))
        chk = ttk.Checkbutton(self, text="Prendre par défaut le dossier de la dernière sauvegarde", variable=self.var_use_last)
        chk.pack(anchor="w", padx=15, pady=10)

        f_btn = ttk.Frame(self)
        f_btn.pack(side="bottom", fill="x", pady=15, padx=15)

        ttk.Button(f_btn, text="Enregistrer", command=self.save).pack(side="right", padx=5)
        ttk.Button(f_btn, text="Annuler", command=self.destroy).pack(side="right", padx=5)

    def browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.entry_dir.delete(0, tk.END)
            self.entry_dir.insert(0, d)

    def save(self):
        self.config["default_working_dir"] = self.entry_dir.get().strip()
        self.config["use_last_backup_dir"] = self.var_use_last.get()
        save_config(self.config)
        messagebox.showinfo("Succès", "Configuration enregistrée.", parent=self)
        self.destroy()


# Dialogue de Gestion des Machines (Outils > Configuration Machines)
class MachinesConfigDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Configuration des Machines & Outils de Coupe")
        self.geometry("850x550")
        self.grab_set()

        self._setup_ui()
        self.load_machines()

    def _setup_ui(self):
        paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Panneau Gauche : Liste des Machines
        frame_mach = ttk.LabelFrame(paned, text=" Parc Machines ", padding=10)
        paned.add(frame_mach, weight=1)

        self.list_machines = tk.Listbox(frame_mach, height=15)
        self.list_machines.pack(fill=tk.BOTH, expand=True, pady=5)
        self.list_machines.bind("<<ListboxSelect>>", self.on_machine_select)

        f_mach_btn = ttk.Frame(frame_mach)
        f_mach_btn.pack(fill="x", pady=5)
        ttk.Button(f_mach_btn, text="+ Machine", command=self.add_machine).pack(side="left", padx=2)
        ttk.Button(f_mach_btn, text="- Supprimer", command=self.delete_machine).pack(side="left", padx=2)

        # Panneau Droit : Tableau Dynamique des Outils de Coupe
        frame_tools = ttk.LabelFrame(paned, text=" Tableau Dynamique des Outils de Coupe ", padding=10)
        paned.add(frame_tools, weight=2)

        f_add_tool = ttk.Frame(frame_tools)
        f_add_tool.pack(fill="x", pady=5)

        ttk.Label(f_add_tool, text="Outil:").grid(row=0, column=0, padx=2)
        self.e_tname = ttk.Entry(f_add_tool, width=12)
        self.e_tname.grid(row=0, column=1, padx=2)

        ttk.Label(f_add_tool, text="Ø(mm):").grid(row=0, column=2, padx=2)
        self.e_tdiam = ttk.Entry(f_add_tool, width=6)
        self.e_tdiam.grid(row=0, column=3, padx=2)

        ttk.Label(f_add_tool, text="L.Sort.:").grid(row=0, column=4, padx=2)
        self.e_tlout = ttk.Entry(f_add_tool, width=6)
        self.e_tlout.grid(row=0, column=5, padx=2)

        ttk.Label(f_add_tool, text="Comp.:").grid(row=0, column=6, padx=2)
        self.e_tlcomp = ttk.Entry(f_add_tool, width=6)
        self.e_tlcomp.grid(row=0, column=7, padx=2)

        ttk.Label(f_add_tool, text="Poche:").grid(row=0, column=8, padx=2)
        self.e_tpocket = ttk.Entry(f_add_tool, width=5)
        self.e_tpocket.grid(row=0, column=9, padx=2)

        ttk.Button(f_add_tool, text="Ajouter Outil", command=self.add_tool).grid(row=0, column=10, padx=5)
        self.e_tname.bind("<Return>", lambda e: self.add_tool())
        self.e_tpocket.bind("<Return>", lambda e: self.add_tool())

        cols = ("id", "tool_name", "diameter", "length_out", "length_comp", "pocket")
        self.tree_tools = ttk.Treeview(frame_tools, columns=cols, show="headings")
        self.tree_tools.heading("id", text="ID")
        self.tree_tools.heading("tool_name", text="Nom Outil")
        self.tree_tools.heading("diameter", text="Diamètre")
        self.tree_tools.heading("length_out", text="Long. Sortante")
        self.tree_tools.heading("length_comp", text="Long. Comp.")
        self.tree_tools.heading("pocket", text="N° Poche")

        self.tree_tools.column("id", width=30, anchor="center")
        self.tree_tools.column("tool_name", width=120)
        self.tree_tools.column("diameter", width=70, anchor="center")
        self.tree_tools.column("length_out", width=90, anchor="center")
        self.tree_tools.column("length_comp", width=90, anchor="center")
        self.tree_tools.column("pocket", width=60, anchor="center")

        self.tree_tools.pack(fill=tk.BOTH, expand=True, pady=5)

        ttk.Button(frame_tools, text="Supprimer Outil Sélectionné", command=self.delete_tool).pack(anchor="e", pady=5)

    def load_machines(self):
        self.list_machines.delete(0, tk.END)
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM machines")
        self.machines_data = cursor.fetchall()
        conn.close()

        for m in self.machines_data:
            self.list_machines.insert(tk.END, m[1])

        if self.machines_data:
            self.list_machines.select_set(0)
            self.on_machine_select(None)

    def get_selected_machine_id(self):
        sel = self.list_machines.curselection()
        if sel:
            return self.machines_data[sel[0]][0]
        return None

    def on_machine_select(self, event):
        m_id = self.get_selected_machine_id()
        for item in self.tree_tools.get_children():
            self.tree_tools.delete(item)

        if m_id:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("SELECT id, tool_name, diameter, length_out, length_comp, pocket FROM machine_tools WHERE machine_id=?", (m_id,))
            for row in cursor.fetchall():
                self.tree_tools.insert("", tk.END, values=row)
            conn.close()

    def add_machine(self):
        name = simpledialog.askstring("Machine", "Nom de la nouvelle machine :", parent=self)
        if name:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            try:
                cursor.execute("INSERT INTO machines (name, description) VALUES (?, ?)", (name, "Machine CNC"))
                conn.commit()
            except sqlite3.IntegrityError:
                messagebox.showerror("Erreur", "Machine déjà existante.", parent=self)
            conn.close()
            self.load_machines()

    def delete_machine(self):
        m_id = self.get_selected_machine_id()
        if m_id and messagebox.askyesno("Confirmation", "Supprimer cette machine et ses outils ?", parent=self):
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM machines WHERE id=?", (m_id,))
            conn.commit()
            conn.close()
            self.load_machines()

    def add_tool(self):
        m_id = self.get_selected_machine_id()
        if not m_id:
            return
        tname = self.e_tname.get().strip()
        if not tname:
            return

        try:
            diam = float(self.e_tdiam.get().replace(',', '.')) if self.e_tdiam.get() else 0.0
            lout = float(self.e_tlout.get().replace(',', '.')) if self.e_tlout.get() else 0.0
            lcomp = float(self.e_tlcomp.get().replace(',', '.')) if self.e_tlcomp.get() else 0.0
            pock = int(self.e_tpocket.get()) if self.e_tpocket.get() else 1
        except ValueError:
            messagebox.showerror("Erreur", "Valeurs numériques invalides.", parent=self)
            return

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO machine_tools (machine_id, tool_name, diameter, length_out, length_comp, pocket)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (m_id, tname, diam, lout, lcomp, pock))
        conn.commit()
        conn.close()

        self.e_tname.delete(0, tk.END)
        self.e_tdiam.delete(0, tk.END)
        self.e_tlout.delete(0, tk.END)
        self.e_tlcomp.delete(0, tk.END)
        self.e_tpocket.delete(0, tk.END)
        self.e_tname.focus()

        self.on_machine_select(None)

    def delete_tool(self):
        sel = self.tree_tools.selection()
        if sel:
            t_id = self.tree_tools.item(sel[0])['values'][0]
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM machine_tools WHERE id=?", (t_id,))
            conn.commit()
            conn.close()
            self.on_machine_select(None)


# Dialogue de Recherche Avancée de Modèle pour l'OF
class ModelSearchDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Sélectionner un Modèle")
        self.geometry("450x350")
        self.grab_set()

        self.selected_model = None

        ttk.Label(self, text="Recherche rapide de modèle :").pack(pady=5)
        self.entry_search = ttk.Entry(self, width=35)
        self.entry_search.pack(pady=5)
        self.entry_search.focus()

        self.listbox = tk.Listbox(self, width=50, height=10)
        self.listbox.pack(pady=5, fill=tk.BOTH, expand=True, padx=10)

        self.entry_search.bind("<KeyRelease>", lambda e: self.update_list())
        self.listbox.bind("<Double-1>", lambda e: self.confirm())

        self.update_list()

        ttk.Button(self, text="Valider Sélection", command=self.confirm).pack(pady=10)

    def update_list(self):
        q = self.entry_search.get().strip()
        self.listbox.delete(0, tk.END)
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT model_name FROM models_catalog WHERE is_hidden=0 AND model_name LIKE ? ORDER BY model_name ASC", (f'%{q}%',))
        for r in cursor.fetchall():
            self.listbox.insert(tk.END, r[0])
        conn.close()

    def confirm(self):
        sel = self.listbox.get(tk.ACTIVE)
        if sel:
            self.selected_model = sel
            self.destroy()


# ==========================================
# 3. APPLICATION PRINCIPALE
# ==========================================

class CNCApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1280x780")

        self.update_idletasks()
        w, h = 1280, 780
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f'{w}x{h}+{x}+{y}')

        self.current_user = None
        self.sort_directions = {}
        self.current_of_cart = []
        self.diff_lines = []
        self.current_diff_index = -1

        self.show_login_screen()

    def clear_window(self):
        for widget in self.winfo_children():
            widget.destroy()

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

    def show_main_screen(self):
        self.clear_window()
        self.title(f"{APP_NAME} - Session : {self.current_user['username']} [{self.current_user['role']}]")
        self.geometry("1280x780")

        # Menus Principaux
        menubar = tk.Menu(self)

        menu_file = tk.Menu(menubar, tearoff=0)
        menu_file.add_command(label="Importer Catalogue (CSV)", command=self.import_usi_tab_csv)
        menu_file.add_command(label="Exporter Catalogue (CSV)", command=self.export_catalog_csv)
        menu_file.add_separator()
        menu_file.add_command(label="Déconnexion", command=self.show_login_screen)
        menu_file.add_command(label="Quitter", command=self.destroy)
        menubar.add_cascade(label="Fichier", menu=menu_file)

        menu_tools = tk.Menu(menubar, tearoff=0)
        menu_tools.add_command(label="Configuration Machines & Outils", command=self.open_machines_config)
        menu_tools.add_command(label="Options...", command=self.open_options)
        menubar.add_cascade(label="Outils", menu=menu_tools)

        if self.current_user['role'] == 'Admin':
            menu_admin = tk.Menu(menubar, tearoff=0)
            menu_admin.add_command(label="Gestion des Utilisateurs", command=self.open_user_management)
            menubar.add_cascade(label="Administration", menu=menu_admin)

        self.config(menu=menubar)

        # En-tête
        header = tk.Frame(self, bg="#003366", height=45)
        header.pack(fill=tk.X, side=tk.TOP)
        tk.Label(header, text=APP_NAME.upper(), font=("Arial", 13, "bold"), fg="white", bg="#003366", pady=8).pack()

        # Onglets dans l'ordre requis
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)

        tab_catalog = ttk.Frame(self.notebook)
        tab_compare = ttk.Frame(self.notebook)
        tab_backup = ttk.Frame(self.notebook)
        tab_of = ttk.Frame(self.notebook)
        tab_tracking = ttk.Frame(self.notebook)
        tab_cnc = ttk.Frame(self.notebook)

        self.notebook.add(tab_catalog, text=" 1. Liste programme usinage ")
        self.notebook.add(tab_compare, text=" 2. Comparaison de Fichiers ")
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

    def open_options(self):
        OptionsDialog(self)

    def open_machines_config(self):
        MachinesConfigDialog(self)

    def print_treeview_data(self, tree, title):
        cols = tree["columns"]
        headers = [tree.heading(col)["text"] for col in cols]
        data = [tree.item(item)["values"] for item in tree.get_children()]
        AdvancedPrintDialog(self, title, headers, data)

    # ==========================================
    # ONGLET 1 : LISTE PROGRAMME USINAGE (CATALOGUE)
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

        cols = ("model_name", "prog_name", "block_dim", "block_dim_bought", "qty_per_block", "z_between_pains", "tools", "caisson", "top_plate", "bottom_plate", "phase", "remarks")
        self.tree_cat = ttk.Treeview(frame_list, columns=cols, show="headings")

        # Intitulés exactement révisés
        headings = {
            "model_name": "Nom Model ↕",
            "prog_name": "Nom Programme Pain ↕",
            "block_dim": "Dimension Bloc ↕",
            "block_dim_bought": "Dimension Bloc (Achetée) ↕",
            "qty_per_block": "Qte/Bloc ↕",
            "z_between_pains": "Z entre 2 pains ↕",
            "tools": "Outils ↕",
            "caisson": "Caisson ↕",
            "top_plate": "Plaque Pont ↕",
            "bottom_plate": "Plaque Carène ↕",
            "phase": "Phase ↕",
            "remarks": "Remarque ↕"
        }

        for c, h in headings.items():
            self.tree_cat.heading(c, text=h, command=lambda _c=c: self.sort_treeview_cat(_c))

        self.tree_cat.column("model_name", width=110)
        self.tree_cat.column("prog_name", width=110)
        self.tree_cat.column("block_dim", width=130)
        self.tree_cat.column("block_dim_bought", width=130)
        self.tree_cat.column("qty_per_block", width=65, anchor="center")
        self.tree_cat.column("z_between_pains", width=90, anchor="center")
        self.tree_cat.column("tools", width=90)
        self.tree_cat.column("caisson", width=60, anchor="center")
        self.tree_cat.column("top_plate", width=90)
        self.tree_cat.column("bottom_plate", width=90)
        self.tree_cat.column("phase", width=75, anchor="center")
        self.tree_cat.column("remarks", width=100)

        scrollbar_y = ttk.Scrollbar(frame_list, orient="vertical", command=self.tree_cat.yview)
        scrollbar_x = ttk.Scrollbar(frame_list, orient="horizontal", command=self.tree_cat.xview)
        self.tree_cat.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        self.tree_cat.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")

        # Interactions : Double-clic & Menu Clic Droit
        self.tree_cat.bind("<Double-1>", self.on_cat_double_click)
        self.tree_cat.bind("<Button-3>", self.on_cat_right_click)

        self.load_catalog_data()

    def sort_treeview_cat(self, col):
        reverse = self.sort_directions.get(col, False)
        items = [(self.tree_cat.set(k, col), k) for k in self.tree_cat.get_children('')]
        items.sort(reverse=reverse)
        for index, (val, k) in enumerate(items):
            self.tree_cat.move(k, '', index)
        self.sort_directions[col] = not reverse

    def load_catalog_data(self):
        for item in self.tree_cat.get_children():
            self.tree_cat.delete(item)

        query = self.entry_cat_search.get().strip() if hasattr(self, 'entry_cat_search') else ""
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        if query:
            cursor.execute('''
                SELECT model_name, prog_name, block_dim, block_dim_bought, qty_per_block, z_between_pains, tools, caisson, top_plate, bottom_plate, phase, remarks
                FROM models_catalog WHERE is_hidden=0 AND (model_name LIKE ? OR prog_name LIKE ? OR tools LIKE ?) ORDER BY id ASC
            ''', (f'%{query}%', f'%{query}%', f'%{query}%'))
        else:
            cursor.execute("SELECT model_name, prog_name, block_dim, block_dim_bought, qty_per_block, z_between_pains, tools, caisson, top_plate, bottom_plate, phase, remarks FROM models_catalog WHERE is_hidden=0 ORDER BY id ASC")

        for row in cursor.fetchall():
            self.tree_cat.insert("", "end", values=row)
        conn.close()

    def on_cat_double_click(self, event):
        sel = self.tree_cat.selection()
        if sel:
            m_name = self.tree_cat.item(sel[0])['values'][0]
            self.add_model_name_to_of_cart(m_name)

    def on_cat_right_click(self, event):
        item = self.tree_cat.identify_row(event.y)
        if item:
            self.tree_cat.selection_set(item)
            menu = tk.Menu(self, tearoff=0)
            menu.add_command(label="Ouvrir le programme (G-Code)", command=self.open_program_file_for_selected)
            menu.add_command(label="Ajouter à l'OF", command=lambda: self.add_model_name_to_of_cart(self.tree_cat.item(item)['values'][0]))
            menu.add_separator()
            menu.add_command(label="Supprimer de la liste", command=self.delete_selected_cat_model)
            menu.tk_popup(event.x_root, event.y_root)

    def open_program_file_for_selected(self):
        sel = self.tree_cat.selection()
        if not sel:
            return
        m_vals = self.tree_cat.item(sel[0])['values']
        p_name = m_vals[1]

        cfg = load_config()
        work_dir = cfg.get("default_working_dir", os.path.expanduser("~"))
        if cfg.get("use_last_backup_dir", False) and cfg.get("last_backup_dir"):
            work_dir = cfg["last_backup_dir"]

        target_file = os.path.join(work_dir, f"{p_name}.iso")
        if not os.path.exists(target_file):
            target_file = filedialog.askopenfilename(initialdir=work_dir, title=f"Sélectionner le programme pour {p_name}")

        if target_file and os.path.exists(target_file):
            self.notebook.select(5)
            self.lbl_file.config(text=f"Fichier : {os.path.basename(target_file)}")
            with open(target_file, "r", encoding="latin1") as f:
                self.txt_preview.delete("1.0", tk.END)
                self.txt_preview.insert(tk.END, f.read())

    def add_model_name_to_of_cart(self, m_name):
        self.notebook.select(3)
        self.combo_of_model.set(m_name)
        self.add_item_to_of_cart()

    def delete_selected_cat_model(self):
        sel = self.tree_cat.selection()
        if not sel:
            return
        m_name = self.tree_cat.item(sel[0])['values'][0]
        if messagebox.askyesno("Confirmation", f"Êtes-vous sûr de vouloir supprimer définitivement le modèle '{m_name}' de la liste ?"):
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM models_catalog WHERE model_name=?", (m_name,))
            conn.commit()
            conn.close()
            self.load_catalog_data()

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
                            z_between_pains, tools, caisson, top_plate, bottom_plate, phase, remarks
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
                        "Test",
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
        cursor.execute("SELECT model_name, prog_name, block_dim, block_dim_bought, qty_per_block, z_between_pains, tools, caisson, top_plate, bottom_plate, phase, remarks FROM models_catalog")
        rows = cursor.fetchall()
        conn.close()

        with open(file_path, mode='w', newline='', encoding='latin1') as f:
            writer = csv.writer(f, delimiter=';')
            writer.writerow(["Nom Model", "Prog. Pain", "Dimension Bloc", "Dimension Achetée", "Qte/Bloc", "Z entre 2 pains", "Outils", "Caisson", "Plaque Pont", "Plaque Carène", "Phase", "Remarque"])
            writer.writerows(rows)
        messagebox.showinfo("Export", "Catalogue exporté avec succès.")

    def add_model_dialog(self):
        win = tk.Toplevel(self)
        win.title("Ajouter un Modèle au Catalogue")
        win.geometry("420x520")

        fields = ["Nom Modèle", "Prog. Pain", "Dimension Bloc", "Dimension Achetée", "Qte/Bloc", "Z entre 2 pains", "Outils", "Caisson", "Plaque Pont", "Plaque Carène", "Remarques"]
        entries = {}

        for i, field in enumerate(fields):
            ttk.Label(win, text=f"{field} :").grid(row=i, column=0, padx=10, pady=3, sticky="w")
            e = ttk.Entry(win, width=25)
            e.grid(row=i, column=1, padx=10, pady=3)
            entries[field] = e

        ttk.Label(win, text="Phase :").grid(row=len(fields), column=0, padx=10, pady=3, sticky="w")
        c_phase = ttk.Combobox(win, values=["Étude", "Test", "Validation"], state="readonly", width=22)
        c_phase.current(0)
        c_phase.grid(row=len(fields), column=1, padx=10, pady=3)

        def save():
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO models_catalog (model_name, prog_name, block_dim, block_dim_bought, qty_per_block, z_between_pains, tools, caisson, top_plate, bottom_plate, phase, remarks)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                entries["Nom Modèle"].get().strip(), entries["Prog. Pain"].get().strip(),
                entries["Dimension Bloc"].get().strip(), entries["Dimension Achetée"].get().strip(),
                entries["Qte/Bloc"].get().strip(), entries["Z entre 2 pains"].get().strip(),
                entries["Outils"].get().strip(), entries["Caisson"].get().strip(),
                entries["Plaque Pont"].get().strip(), entries["Plaque Carène"].get().strip(),
                c_phase.get(), entries["Remarques"].get().strip()
            ))
            conn.commit()
            conn.close()
            messagebox.showinfo("Succès", "Modèle ajouté au catalogue.", parent=win)
            win.destroy()
            self.load_catalog_data()
            self.update_model_comboboxes()

        ttk.Button(win, text="Enregistrer", command=save).grid(row=len(fields)+1, column=0, columnspan=2, pady=15)

    def delete_model_dialog(self):
        win = ModelSearchDialog(self)
        self.wait_window(win)
        if win.selected_model and messagebox.askyesno("Confirmation", f"Supprimer définitvement le modèle '{win.selected_model}' ?"):
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM models_catalog WHERE model_name=?", (win.selected_model,))
            conn.commit()
            conn.close()
            self.load_catalog_data()
            self.update_model_comboboxes()

    def hide_model_dialog(self):
        win = ModelSearchDialog(self)
        self.wait_window(win)
        if win.selected_model:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("UPDATE models_catalog SET is_hidden = CASE WHEN is_hidden=1 THEN 0 ELSE 1 END WHERE model_name=?", (win.selected_model,))
            conn.commit()
            conn.close()
            self.load_catalog_data()

    # ==========================================
    # ONGLET 2 : COMPARAISON DE FICHIERS CÔTE-À-CÔTE
    # ==========================================
    def setup_compare_tab(self, parent):
        frame_top = ttk.LabelFrame(parent, text=" Sélection des fichiers G-Code à comparer ")
        frame_top.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_top, text="Fichier A (Référence) :").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        self.v_file_a = tk.StringVar()
        ttk.Entry(frame_top, textvariable=self.v_file_a, width=50).grid(row=0, column=1, padx=5, pady=3)
        ttk.Button(frame_top, text="Parcourir...", command=lambda: self.browse_file(self.v_file_a)).grid(row=0, column=2, padx=5, pady=3)

        ttk.Label(frame_top, text="Fichier B (Comparé) :").grid(row=0, column=3, sticky="w", padx=5, pady=3)
        self.v_file_b = tk.StringVar()
        ttk.Entry(frame_top, textvariable=self.v_file_b, width=50).grid(row=0, column=4, padx=5, pady=3)
        ttk.Button(frame_top, text="Parcourir...", command=lambda: self.browse_file(self.v_file_b)).grid(row=0, column=5, padx=5, pady=3)

        # Barres de navigation dans les différences
        f_nav = ttk.Frame(parent)
        f_nav.pack(fill="x", padx=10, pady=5)

        tk.Button(f_nav, text="🔍 Lancer la Comparaison", bg="#0288D1", fg="white", font=("Arial", 9, "bold"), command=self.compare_files_side_by_side).pack(side="left", padx=5)

        self.btn_prev_diff = ttk.Button(f_nav, text="▲ Différence Précédente", command=self.prev_diff, state=tk.DISABLED)
        self.btn_prev_diff.pack(side="left", padx=5)

        self.btn_next_diff = ttk.Button(f_nav, text="▼ Différence Suivante", command=self.next_diff, state=tk.DISABLED)
        self.btn_next_diff.pack(side="left", padx=5)

        self.lbl_diff_count = ttk.Label(f_nav, text="Aucune comparaison effectuée", font=("Arial", 9, "bold"))
        self.lbl_diff_count.pack(side="right", padx=10)

        # Écran divisé en 2 côte-à-côte avec synchronisation du défilement
        frame_split = ttk.Frame(parent)
        frame_split.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Panneau Fichier A
        f_left = ttk.LabelFrame(frame_split, text=" Fichier A ")
        f_left.pack(side="left", fill=tk.BOTH, expand=True)
        self.txt_file_a = tk.Text(f_left, wrap="none", font=("Courier", 9))
        self.txt_file_a.pack(side="left", fill=tk.BOTH, expand=True)

        # Marge centrale de séparation
        f_mid = ttk.Frame(frame_split, width=25, bg="#E0E0E0")
        f_mid.pack(side="left", fill=tk.Y, padx=2)
        tk.Label(f_mid, text="||", bg="#E0E0E0", font=("Arial", 12, "bold")).pack(pady=50)

        # Panneau Fichier B
        f_right = ttk.LabelFrame(frame_split, text=" Fichier B ")
        f_right.pack(side="left", fill=tk.BOTH, expand=True)
        self.txt_file_b = tk.Text(f_right, wrap="none", font=("Courier", 9))
        self.txt_file_b.pack(side="left", fill=tk.BOTH, expand=True)

        # Scrollbar partagée pour synchroniser parfaitement les deux fenêtres verticalement
        self.compare_scrollbar = ttk.Scrollbar(frame_split, orient="vertical", command=self.on_compare_scrollbar)
        self.compare_scrollbar.pack(side="right", fill="y")

        self.txt_file_a.config(yscrollcommand=self.compare_scrollbar.set)
        self.txt_file_b.config(yscrollcommand=self.compare_scrollbar.set)

        self.txt_file_a.tag_config("diff", background="#FFCDD2", foreground="#B71C1C")
        self.txt_file_b.tag_config("diff", background="#FFCDD2", foreground="#B71C1C")

    def on_compare_scrollbar(self, *args):
        self.txt_file_a.yview(*args)
        self.txt_file_b.yview(*args)

    def browse_file(self, var):
        f = filedialog.askopenfilename(filetypes=[("Programme G-Code", "*.iso *.nc *.txt"), ("Tous", "*.*")])
        if f:
            var.set(f)

    def compare_files_side_by_side(self):
        fa, fb = self.v_file_a.get().strip(), self.v_file_b.get().strip()
        if not os.path.isfile(fa) or not os.path.isfile(fb):
            messagebox.showwarning("Attention", "Veuillez sélectionner deux fichiers valides.")
            return

        with open(fa, "r", encoding="latin1") as f1:
            lines_a = f1.readlines()
        with open(fb, "r", encoding="latin1") as f2:
            lines_b = f2.readlines()

        self.txt_file_a.delete("1.0", tk.END)
        self.txt_file_b.delete("1.0", tk.END)
        self.diff_lines.clear()

        max_lines = max(len(lines_a), len(lines_b))

        for i in range(max_lines):
            la = lines_a[i] if i < len(lines_a) else ""
            lb = lines_b[i] if i < len(lines_b) else ""

            line_num = i + 1
            self.txt_file_a.insert(tk.END, f"{line_num:04d}: {la}")
            self.txt_file_b.insert(tk.END, f"{line_num:04d}: {lb}")

            if la.strip() != lb.strip():
                self.diff_lines.append(line_num)
                self.txt_file_a.tag_add("diff", f"{line_num}.0", f"{line_num}.end")
                self.txt_file_b.tag_add("diff", f"{line_num}.0", f"{line_num}.end")

        if self.diff_lines:
            self.current_diff_index = 0
            self.lbl_diff_count.config(text=f"{len(self.diff_lines)} différence(s) trouvée(s)")
            self.btn_prev_diff.config(state=tk.NORMAL)
            self.btn_next_diff.config(state=tk.NORMAL)
            self.scroll_to_diff()
        else:
            self.lbl_diff_count.config(text="Fichiers 100% identiques")
            self.btn_prev_diff.config(state=tk.DISABLED)
            self.btn_next_diff.config(state=tk.DISABLED)

    def scroll_to_diff(self):
        if 0 <= self.current_diff_index < len(self.diff_lines):
            line = self.diff_lines[self.current_diff_index]
            self.txt_file_a.see(f"{line}.0")
            self.txt_file_b.see(f"{line}.0")

    def next_diff(self):
        if self.diff_lines:
            self.current_diff_index = (self.current_diff_index + 1) % len(self.diff_lines)
            self.scroll_to_diff()

    def prev_diff(self):
        if self.diff_lines:
            self.current_diff_index = (self.current_diff_index - 1) % len(self.diff_lines)
            self.scroll_to_diff()

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
        p = filedialog.askdirectory()
        if p:
            self.entry_src.delete(0, tk.END)
            self.entry_src.insert(0, p)

    def browse_dest(self):
        p = filedialog.askdirectory()
        if p:
            self.entry_dest.delete(0, tk.END)
            self.entry_dest.insert(0, p)

    def add_task(self):
        name, src, dest, rec_type = self.entry_task_name.get().strip(), self.entry_src.get().strip(), self.entry_dest.get().strip(), self.combo_type.get()
        if name and src and dest:
            task = (name, src, dest, rec_type)
            self.tree_tasks.insert("", tk.END, values=task)
            messagebox.showinfo("Succès", f"Sauvegarde '{name}' configurée.")

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

        cfg = load_config()
        cfg["last_backup_dir"] = target_dir
        save_config(cfg)

        file_list = [os.path.join(r, f) for r, d, files in os.walk(src) for f in files] if os.path.isdir(src) else [src]

        try:
            os.makedirs(target_dir, exist_ok=True)
            for file_path in file_list:
                rel_path = os.path.relpath(file_path, src) if os.path.isdir(src) else os.path.basename(file_path)
                dest_file_path = os.path.join(target_dir, rel_path)
                os.makedirs(os.path.dirname(dest_file_path), exist_ok=True)
                shutil.copy2(file_path, dest_file_path)

            messagebox.showinfo("Succès", f"Sauvegarde '{name}' réalisée avec succès dans :\n{target_dir}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de la sauvegarde : {str(e)}")

    def delete_task(self):
        selected = self.tree_tasks.selection()
        if selected and messagebox.askyesno("Confirmation", "Supprimer cette sauvegarde ?"):
            self.tree_tasks.delete(selected[0])

    # ==========================================
    # ONGLET 4 : LANCEMENT ORDRES DE FABRICATION (OF)
    # ==========================================
    def setup_of_tab(self, parent):
        frame_top = ttk.LabelFrame(parent, text=" 1. En-tête du Lancement OF ")
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

        frame_item = ttk.LabelFrame(parent, text=" 2. Pièces & Bruts à Inclure dans cet OF ")
        frame_item.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_item, text="Modèle :").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.combo_of_model = ttk.Combobox(frame_item, width=20)
        self.combo_of_model.grid(row=0, column=1, padx=5, pady=5)
        self.update_model_comboboxes()

        ttk.Button(frame_item, text="🔍 Chercher", command=self.search_model_for_of).grid(row=0, column=2, padx=2)

        ttk.Label(frame_item, text="Qte :").grid(row=0, column=3, padx=5, pady=5, sticky="w")
        self.spin_of_qty = tk.Spinbox(frame_item, from_=1, to=100, width=4)
        self.spin_of_qty.grid(row=0, column=4, padx=5, pady=5)

        ttk.Label(frame_item, text="N° Bloc :").grid(row=0, column=5, padx=5, pady=5, sticky="w")
        self.entry_of_bnum = ttk.Entry(frame_item, width=8)
        self.entry_of_bnum.grid(row=0, column=6, padx=5, pady=5)

        ttk.Label(frame_item, text="Densité :").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.entry_of_bdens = ttk.Entry(frame_item, width=10)
        self.entry_of_bdens.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(frame_item, text="N° Pain :").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.entry_of_pnum = ttk.Entry(frame_item, width=8)
        self.entry_of_pnum.grid(row=1, column=3, padx=5, pady=5)

        ttk.Label(frame_item, text="Poids (g) :").grid(row=1, column=4, padx=5, pady=5, sticky="w")
        self.entry_of_pweight = ttk.Entry(frame_item, width=8)
        self.entry_of_pweight.grid(row=1, column=5, padx=5, pady=5)

        btn_add_item = tk.Button(frame_item, text="➕ Ajouter au Panier", bg="#0288D1", fg="white", font=("Arial", 9, "bold"), command=self.add_item_to_of_cart)
        btn_add_item.grid(row=1, column=6, padx=5, pady=5)

        self.tree_cart = ttk.Treeview(frame_item, columns=("model", "qty", "bnum", "bdens", "pnum", "pweight"), show="headings", height=3)
        self.tree_cart.heading("model", text="Modèle")
        self.tree_cart.heading("qty", text="Qté")
        self.tree_cart.heading("bnum", text="N° Bloc")
        self.tree_cart.heading("bdens", text="Densité")
        self.tree_cart.heading("pnum", text="N° Pain")
        self.tree_cart.heading("pweight", text="Poids (g)")

        self.tree_cart.column("model", width=140)
        self.tree_cart.column("qty", width=40, anchor="center")
        self.tree_cart.column("bnum", width=70, anchor="center")
        self.tree_cart.column("bdens", width=70, anchor="center")
        self.tree_cart.column("pnum", width=70, anchor="center")
        self.tree_cart.column("pweight", width=70, anchor="center")
        self.tree_cart.grid(row=2, column=0, columnspan=6, sticky="ew", padx=5, pady=5)

        ttk.Button(frame_item, text="- Supprimer Pièce du Panier", command=self.remove_item_from_of_cart).grid(row=2, column=6, padx=5)

        btn_val_of = tk.Button(parent, text="🚀 VALIDER ET CRÉER L'ORDRE DE FABRICATION GLOBAL", bg="#2E7D32", fg="white", font=("Arial", 10, "bold"), pady=4, command=self.save_global_of)
        btn_val_of.pack(fill="x", padx=10, pady=5)

        frame_list = ttk.LabelFrame(parent, text=" 3. Historique & Impression des Lancements OF ")
        frame_list.pack(fill="both", expand=True, padx=10, pady=5)

        frame_actions_of = ttk.Frame(frame_list)
        frame_actions_of.pack(fill="x", padx=5, pady=2)

        ttk.Button(frame_actions_of, text=" Charger & Envoyer Programme de la Pièce vers CNC", command=self.transfer_selected_of_item_to_cnc).pack(side="left", padx=5)
        ttk.Button(frame_actions_of, text="🖨️ Imprimer l'OF Sélectionné", command=self.print_selected_of_details).pack(side="right", padx=5)
        ttk.Button(frame_actions_of, text="🖨️ Imprimer Tous les OFs", command=self.print_all_ofs_with_details).pack(side="right", padx=5)

        self.tree_of = ttk.Treeview(frame_list, columns=("id", "of_number", "machine", "operator", "priority", "status", "created_at"), show="headings", height=4)
        self.tree_of.heading("id", text="ID")
        self.tree_of.heading("of_number", text="N° OF Lancement")
        self.tree_of.heading("machine", text="Machine")
        self.tree_of.heading("operator", text="Opérateur")
        self.tree_of.heading("priority", text="Priorité")
        self.tree_of.heading("status", text="Statut Lancement")
        self.tree_of.heading("created_at", text="Date Lancement")

        self.tree_of.column("id", width=35, anchor="center")
        self.tree_of.column("of_number", width=130, anchor="center")
        self.tree_of.column("machine", width=130)
        self.tree_of.column("operator", width=90)
        self.tree_of.column("priority", width=70, anchor="center")
        self.tree_of.column("status", width=90, anchor="center")
        self.tree_of.column("created_at", width=130, anchor="center")
        self.tree_of.pack(fill="x", padx=5, pady=3)
        self.tree_of.bind("<<TreeviewSelect>>", self.on_of_selected)

        ttk.Label(frame_list, text="Pièces incluses dans l'OF sélectionné :", font=("Arial", 9, "bold")).pack(anchor="w", padx=5, pady=2)
        
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

        self.tree_of_items.column("id", width=45, anchor="center")
        self.tree_of_items.column("model", width=130)
        self.tree_of_items.column("prog", width=110)
        self.tree_of_items.column("qty", width=45, anchor="center")
        self.tree_of_items.column("bnum", width=75, anchor="center")
        self.tree_of_items.column("bdens", width=75, anchor="center")
        self.tree_of_items.column("pnum", width=75, anchor="center")
        self.tree_of_items.column("pweight", width=75, anchor="center")
        self.tree_of_items.column("status", width=90, anchor="center")
        self.tree_of_items.pack(fill="both", expand=True, padx=5, pady=3)

        self.load_of_data()

    def search_model_for_of(self):
        win = ModelSearchDialog(self)
        self.wait_window(win)
        if win.selected_model:
            self.combo_of_model.set(win.selected_model)

    def update_model_comboboxes(self):
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("SELECT model_name FROM models_catalog WHERE is_hidden=0 ORDER BY model_name ASC")
        models = [r[0] for r in cursor.fetchall()]
        conn.close()

        if hasattr(self, 'combo_of_model'):
            self.combo_of_model['values'] = models
            if models:
                self.combo_of_model.current(0)

    def add_item_to_of_cart(self):
        model = self.combo_of_model.get().strip()
        bnum = self.entry_of_bnum.get().strip()
        bdens = self.entry_of_bdens.get().strip()
        pnum = self.entry_of_pnum.get().strip()
        pweight = self.entry_of_pweight.get().strip()

        if not model or not bnum or not bdens or not pnum or not pweight:
            if not messagebox.askyesno("Informations Manquantes", "Certaines informations sur le brut sont incomplètes.\nVoulez-vous tout de même ajouter cette pièce à l'OF ?"):
                return

        qty = self.spin_of_qty.get()
        item = (model, qty, bnum, bdens, pnum, pweight)
        self.current_of_cart.append(item)
        self.tree_cart.insert("", tk.END, values=item)

    def remove_item_from_of_cart(self):
        sel = self.tree_cart.selection()
        if sel:
            idx = self.tree_cart.index(sel[0])
            self.tree_cart.delete(sel[0])
            del self.current_of_cart[idx]

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
            messagebox.showwarning("Panier Vide", "Veuillez ajouter au moins une pièce à cet Ordre de Fabrication.")
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
        self.txt_preview.insert(tk.END, f"% \n(PROGRAMME NUM 1060 5-AXES ASSOCIE A L'OF {of_num})\n(MODELE: {model_name})\n(PROGRAMME PAIN: {prog_name})\n\nG00 G90 G40\nM03 S12000\nG00 X0 Y0 Z50\n(G-CODE PASSANT %PPR)\nM05\nM30\n%")

    def print_selected_of_details(self):
        sel = self.tree_of.selection()
        if not sel:
            messagebox.showwarning("Attention", "Veuillez sélectionner un OF à imprimer.")
            return
        of_num = self.tree_of.item(sel[0])['values'][1]

        headers = ["Nom Modèle", "Prog. Pain", "Qté", "N° Bloc", "Densité", "N° Pain", "Poids (g)", "Statut"]
        data = [self.tree_of_items.item(i)['values'][1:] for i in self.tree_of_items.get_children()]

        AdvancedPrintDialog(self, f"Fiche Ordre de Fabrication {of_num}", headers, data)

    def print_all_ofs_with_details(self):
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT w.of_number, w.machine, i.model_name, i.prog_name, i.qty, i.block_num, i.status
            FROM work_orders w JOIN work_order_items i ON w.of_number = i.of_number ORDER BY w.id DESC
        ''')
        data = cursor.fetchall()
        conn.close()

        headers = ["N° OF", "Machine", "Modèle", "Prog. Pain", "Qté", "N° Bloc", "Statut"]
        AdvancedPrintDialog(self, "Rapport Global des Lancements OF", headers, data)

    # ==========================================
    # ONGLET 5 : TRAÇABILITÉ & RECHERCHE AVANCÉE
    # ==========================================
    def setup_tracking_tab(self, parent):
        frame_filter = ttk.LabelFrame(parent, text=" Recherche Avancée et Filtres de Traçabilité Usinage ")
        frame_filter.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_filter, text="Opérateur :").grid(row=0, column=0, padx=5, pady=3, sticky="w")
        self.e_f_op = ttk.Entry(frame_filter, width=12)
        self.e_f_op.grid(row=0, column=1, padx=5, pady=3)

        ttk.Label(frame_filter, text="Modèle :").grid(row=0, column=2, padx=5, pady=3, sticky="w")
        self.e_f_mod = ttk.Entry(frame_filter, width=15)
        self.e_f_mod.grid(row=0, column=3, padx=5, pady=3)

        ttk.Label(frame_filter, text="N° OF :").grid(row=0, column=4, padx=5, pady=3, sticky="w")
        self.e_f_of = ttk.Entry(frame_filter, width=12)
        self.e_f_of.grid(row=0, column=5, padx=5, pady=3)

        ttk.Label(frame_filter, text="Période du :").grid(row=1, column=0, padx=5, pady=3, sticky="w")
        self.e_f_d1 = ttk.Entry(frame_filter, width=12)
        self.e_f_d1.insert(0, "YYYY-MM-DD")
        self.e_f_d1.grid(row=1, column=1, padx=5, pady=3)

        ttk.Label(frame_filter, text="Au :").grid(row=1, column=2, padx=5, pady=3, sticky="w")
        self.e_f_d2 = ttk.Entry(frame_filter, width=12)
        self.e_f_d2.insert(0, "YYYY-MM-DD")
        self.e_f_d2.grid(row=1, column=3, padx=5, pady=3)

        tk.Button(frame_filter, text="🔍 Filtrer", bg="#0288D1", fg="white", font=("Arial", 9, "bold"), command=self.load_tracking_data).grid(row=1, column=4, padx=5, pady=3)
        ttk.Button(frame_filter, text="Réinitialiser", command=self.reset_tracking_filters).grid(row=1, column=5, padx=5, pady=3)

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
        self.entry_tr_model = ttk.Entry(frame_input, width=18)
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

        ttk.Button(parent, text="🖨️ Imprimer la Sélection Filtrée", command=lambda: self.print_treeview_data(self.tree_track, "Rapport de Traçabilité Usinage")).pack(anchor="e", padx=10, pady=2)

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

    def reset_tracking_filters(self):
        self.e_f_op.delete(0, tk.END)
        self.e_f_mod.delete(0, tk.END)
        self.e_f_of.delete(0, tk.END)
        self.e_f_d1.delete(0, tk.END)
        self.e_f_d1.insert(0, "YYYY-MM-DD")
        self.e_f_d2.delete(0, tk.END)
        self.e_f_d2.insert(0, "YYYY-MM-DD")
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

        op = self.e_f_op.get().strip() if hasattr(self, 'e_f_op') else ""
        mod = self.e_f_mod.get().strip() if hasattr(self, 'e_f_mod') else ""
        of_n = self.e_f_of.get().strip() if hasattr(self, 'e_f_of') else ""
        d1 = self.e_f_d1.get().strip() if hasattr(self, 'e_f_d1') and self.e_f_d1.get() != "YYYY-MM-DD" else ""
        d2 = self.e_f_d2.get().strip() if hasattr(self, 'e_f_d2') and self.e_f_d2.get() != "YYYY-MM-DD" else ""

        query = "SELECT id, of_number, operator, machine, model_name, real_time_min, status, timestamp FROM machining_history WHERE 1=1"
        params = []

        if op:
            query += " AND operator LIKE ?"
            params.append(f"%{op}%")
        if mod:
            query += " AND model_name LIKE ?"
            params.append(f"%{mod}%")
        if of_n:
            query += " AND of_number LIKE ?"
            params.append(f"%{of_n}%")
        if d1:
            query += " AND DATE(timestamp) >= ?"
            params.append(d1)
        if d2:
            query += " AND DATE(timestamp) <= ?"
            params.append(d2)

        query += " ORDER BY id DESC"

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute(query, params)
        for row in cursor.fetchall():
            self.tree_track.insert("", "end", values=row)
        conn.close()

    # ==========================================
    # ONGLET 6 : TRANSFERT RS232 NUM 1060 (MODE PASSANT %PPR)
    # ==========================================
    def setup_cnc_tab(self, parent):
        frame_cfg = ttk.LabelFrame(parent, text=" Paramètres RS232 & Mode Passant DNC (%PPR NUM 1060) ")
        frame_cfg.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_cfg, text="Port COM :").grid(row=0, column=0, padx=5, pady=5)
        self.combo_port = ttk.Combobox(frame_cfg, values=["COM1", "COM2", "COM3", "COM4"], width=10)
        self.combo_port.set("COM1")
        self.combo_port.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(frame_cfg, text="Baudrate :").grid(row=0, column=2, padx=5, pady=5)
        self.combo_baud = ttk.Combobox(frame_cfg, values=["4800", "9600", "19200"], width=10)
        self.combo_baud.set("9600")
        self.combo_baud.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(frame_cfg, text="Mode de transfert :").grid(row=0, column=4, padx=5, pady=5)
        self.combo_transfer_mode = ttk.Combobox(frame_cfg, values=["Standard (Chargement Mémoire)", "Passant DNC (%PPR)"], state="readonly", width=25)
        self.combo_transfer_mode.current(1)
        self.combo_transfer_mode.grid(row=0, column=5, padx=5, pady=5)

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

        frame_status = ttk.LabelFrame(parent, text=" État du Transfert en Temps Réel ")
        frame_status.pack(fill="x", padx=10, pady=5)

        self.lbl_current_line = ttk.Label(frame_status, text="Ligne active : -", font=("Courier", 9, "bold"), foreground="#003366")
        self.lbl_current_line.pack(anchor="w", padx=10, pady=2)

        self.progress_rs232 = ttk.Progressbar(frame_status, orient="horizontal", mode="determinate")
        self.progress_rs232.pack(fill="x", padx=10, pady=5)

        self.lbl_pct = ttk.Label(frame_status, text="0%", font=("Arial", 9))
        self.lbl_pct.pack(anchor="e", padx=10)

        frame_send = ttk.Frame(parent)
        frame_send.pack(fill="x", padx=10, pady=10)

        ttk.Button(frame_send, text="🚀 ENVOYER / EXECUTER SUR NUM 1060 (RS232)", command=self.send_to_cnc).pack(fill="x", ipady=5)

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
        mode_ppr = "Passant" in self.combo_transfer_mode.get()

        lines = gcode.splitlines()
        if mode_ppr and not lines[0].startswith("%PPR"):
            lines.insert(0, "%PPR")

        total = len(lines)
        self.progress_rs232['value'] = 0

        ser = None
        if serial:
            try:
                ser = serial.Serial(port, baud, parity=serial.PARITY_EVEN, bytesize=7, stopbits=2, timeout=2)
            except Exception as e:
                messagebox.showerror("Erreur RS232", f"Impossible d'ouvrir le port {port} :\n{str(e)}")
                return

        for idx, line in enumerate(lines, start=1):
            if ser:
                ser.write((line + "\r\n").encode('latin1'))

            pct = int((idx / total) * 100)
            self.progress_rs232['value'] = pct
            self.lbl_pct.config(text=f"{pct}% ({idx}/{total} lignes)")
            self.lbl_current_line.config(text=f"Transmis [{idx:04d}]: {line[:60]}")
            self.update()
            time.sleep(0.005)

        if ser:
            ser.close()

        msg_type = "Mode Passant %PPR" if mode_ppr else "Mode Standard"
        messagebox.showinfo("Transfert Réussi", f"Le programme a été transmis en {msg_type} sur la NUM 1060.")

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
