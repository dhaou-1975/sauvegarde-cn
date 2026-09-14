import os
import sys
import hashlib
import json
import filecmp
import tempfile
import webbrowser
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# Tentative d'importation de ReportLab pour la génération du PDF A4 parfait
HAS_REPORTLAB = False
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False


CONFIG_FILE = "config_cnc.json"
DEFAULT_PASSWORD_HASH = hashlib.sha256("1234".encode()).hexdigest() # Mot de passe par défaut : 1234


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
class CNCBackupManager(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GESTIONNAIRE DE SAUVEGARDE DES PROGRAMMES CNC")
        self.geometry("1000x640")

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
        # Barre de menu supérieur
        menubar = tk.Menu(self)
        menu_admin = tk.Menu(menubar, tearoff=0)
        menu_admin.add_command(label="Changer le mot de passe", command=self.change_password)
        menu_admin.add_separator()
        menu_admin.add_command(label="Quitter", command=self.quit)
        menubar.add_cascade(label="Sécurité / Options", menu=menu_admin)
        self.config(menu=menubar)

        # En-tête principal
        header = tk.Label(self, text="GESTIONNAIRE DE SAUVEGARDE DES PROGRAMMES CNC",
                          bg="#0B3C5D", fg="white", font=("Arial", 14, "bold"), py=8)
        header.pack(fill=tk.X)

        # Onglets
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        tab_comp = ttk.Frame(notebook)
        notebook.add(tab_comp, text="Configuration & Planification des Sauvegardes")

        # Zone Sélection des Dossiers
        frame_dirs = ttk.LabelFrame(tab_comp, text="Sélection des dossiers à comparer")
        frame_dirs.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(frame_dirs, text="Dossier A (Référence) :").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.var_dir_a = tk.StringVar()
        ttk.Entry(frame_dirs, textvariable=self.var_dir_a, width=68).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame_dirs, text="Parcourir", command=lambda: self.browse_dir(self.var_dir_a)).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(frame_dirs, text="Dossier B (Comparé) :").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.var_dir_b = tk.StringVar()
        ttk.Entry(frame_dirs, textvariable=self.var_dir_b, width=68).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame_dirs, text="Parcourir", command=lambda: self.browse_dir(self.var_dir_b)).grid(row=1, column=2, padx=5, pady=5)

        # Double bouton de comparaison
        frame_btns = ttk.Frame(tab_comp)
        frame_btns.pack(fill=tk.X, padx=10, pady=5)

        btn_fast = tk.Button(frame_btns, text="Lancer la Comparaison Rapide (Dates/Tailles)", 
                             bg="#0288D1", fg="white", font=("Arial", 10, "bold"), py=4,
                             command=lambda: self.run_comparison(deep=False))
        btn_fast.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        btn_deep = tk.Button(frame_btns, text="🔍 Comparaison Approfondie (Contenu Texte CNC)", 
                             bg="#2E7D32", fg="white", font=("Arial", 10, "bold"), py=4,
                             command=lambda: self.run_comparison(deep=True))
        btn_deep.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        # Tableau des résultats
        frame_grid = ttk.LabelFrame(tab_comp, text="Résultats de la comparaison")
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

        # Tags de couleur : Fond ROUGE pour statut DIFFERENT
        self.tree.tag_configure("different_tag", background="#D32F2F", foreground="white")
        self.tree.tag_configure("identique_tag", background="white", foreground="black")

        scrollbar = ttk.Scrollbar(frame_grid, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Zone Bas de fenêtre (Exports & Impression)
        frame_bottom = ttk.Frame(tab_comp)
        frame_bottom.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(frame_bottom, text="Exporter en TXT", command=self.export_txt).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_bottom, text="Exporter en Excel (.csv)", command=self.export_csv).pack(side=tk.LEFT, padx=5)

        btn_print = tk.Button(frame_bottom, text="🖨️ Imprimer (Menu Impression Système)", 
                              bg="#424242", fg="white", font=("Arial", 9, "bold"),
                              command=self.print_a4_formatted)
        btn_print.pack(side=tk.RIGHT, padx=5)

    def browse_dir(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def run_comparison(self, deep=False):
        dir_a = self.var_dir_a.get()
        dir_b = self.var_dir_b.get()

        if not os.path.isdir(dir_a) or not os.path.isdir(dir_b):
            messagebox.showwarning("Attention", "Veuillez sélectionner deux dossiers valides.")
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
                    # Comparaison du contenu texte réel (ignore les dates si le texte est identique)
                    is_same = filecmp.cmp(path_a, path_b, shallow=False)
                else:
                    # Comparaison rapide basée sur la date et la taille
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
        doc = SimpleDocTemplate(pdf_filename, pagesize=A4,
                                rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)

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
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0B3C5D")),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('TOPPADDING', (0,0), (-1,-1), 3),
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
        import csv
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow(["Statut", "Fichier", "Date A", "Date B"])
                for item in self.tree.get_children():
                    writer.writerow(self.tree.item(item, "values"))
            messagebox.showinfo("Export", "Export CSV/Excel réussi !")

if __name__ == "__main__":
    app = CNCBackupManager()
    app.mainloop()
