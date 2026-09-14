import os
import sys
import shutil
import subprocess
import threading
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Imports optionnels pour impression native Windows et Excel
try:
    import win32api
    import win32print
    WIN32_PRINT_AVAILABLE = True
except ImportError:
    WIN32_PRINT_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

MAX_BACKUPS = 5

def get_timestamp():
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

def get_timestamp_readable():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def clean_old_backups(target_root_dir, log_widget=None):
    if not os.path.exists(target_root_dir):
        return
    
    subdirs = [
        os.path.join(target_root_dir, d) for d in os.listdir(target_root_dir)
        if os.path.isdir(os.path.join(target_root_dir, d))
    ]
    subdirs.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    
    if len(subdirs) > MAX_BACKUPS:
        to_delete = subdirs[MAX_BACKUPS:]
        for folder in to_delete:
            msg = f"Nettoyage : suppression -> {os.path.basename(folder)}\n"
            if log_widget:
                log_widget.insert(tk.END, msg)
            try:
                shutil.rmtree(folder)
            except Exception as e:
                if log_widget:
                    log_widget.insert(tk.END, f"Erreur de suppression : {e}\n")

def run_robocopy_backup(source_path, destination_root_dir, folder_prefix, log_widget=None, status_label=None):
    if log_widget:
        log_widget.insert(tk.END, f"[{get_timestamp_readable()}] Vérification source : {source_path}\n", "info")
        log_widget.see(tk.END)
    
    if not os.path.exists(source_path):
        err_msg = f"Source inaccessible : {source_path}\nVérifiez le réseau ou le chemin."
        if status_label:
            status_label.config(text="Erreur : Source inaccessible", fg="red")
        if log_widget:
            log_widget.insert(tk.END, f"[ERREUR] {err_msg}\n", "error")
            log_widget.see(tk.END)
        return False, err_msg

    if status_label:
        status_label.config(text="Copie des fichiers en cours...", fg="blue")
    
    timestamp = get_timestamp()
    folder_name = f"{folder_prefix}_{timestamp}" if folder_prefix else f"Sauvegarde_{timestamp}"
    final_destination = os.path.join(destination_root_dir, folder_name)

    if log_widget:
        log_widget.insert(tk.END, f"Destination : {final_destination}\n", "info")
        log_widget.see(tk.END)
    os.makedirs(final_destination, exist_ok=True)

    if os.path.isfile(source_path):
        cmd = f'copy /Y "{source_path}" "{final_destination}\\"'
    else:
        cmd = f'robocopy "{source_path}" "{final_destination}" /E /R:2 /W:3 /NP /NDL'
        
    subprocess.call(cmd, shell=True)

    clean_old_backups(destination_root_dir, log_widget)
    
    if status_label:
        status_label.config(text="Sauvegarde terminée avec succès !", fg="green")
    if log_widget:
        log_widget.insert(tk.END, f"[{get_timestamp_readable()}] [SUCCÈS] Sauvegarde réussie dans : {final_destination}\n\n", "success")
        log_widget.see(tk.END)
    
    return True, final_destination

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestionnaire de Sauvegarde & Comparaison Industrielle CNC V2")
        self.root.geometry("920x720")
        self.root.configure(bg="#f4f6f9")

        # Scheduler State
        self.scheduler_thread = None
        self.scheduler_running = False

        self.setup_ui()

    def setup_ui(self):
        # Header Banner
        header = tk.Frame(self.root, bg="#003366", pady=12)
        header.pack(fill="x")
        lbl_title = tk.Label(header, text="SAUVEGARDE & COMPARAISON INDUSTRIELLE CNC - V2", 
                             font=("Segoe UI", 14, "bold"), fg="white", bg="#003366")
        lbl_title.pack()

        # Notebook (Tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Tab 1: Sauvegarde
        self.tab_backup = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_backup, text=" 💾 Sauvegarde ")
        self.build_backup_tab()

        # Tab 2: Comparaison
        self.tab_compare = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_compare, text=" 🔍 Comparaison ")
        self.build_compare_tab()

        # Tab 3: Planification
        self.tab_schedule = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_schedule, text=" ⏰ Planification Auto ")
        self.build_schedule_tab()

        # Status Bar
        self.status_frame = tk.Frame(self.root, bg="#e0e0e0", height=25)
        self.status_frame.pack(fill="x", side="bottom")
        self.status_label = tk.Label(self.status_frame, text="Prêt.", font=("Segoe UI", 9, "italic"), bg="#e0e0e0", fg="#333333")
        self.status_label.pack(side="left", padx=10)

    # ------------------------------------------------------------------
    # TAB 1 : SAUVEGARDE
    # ------------------------------------------------------------------
    def build_backup_tab(self):
        frame = tk.Frame(self.tab_backup, bg="#f4f6f9", padx=15, pady=15)
        frame.pack(fill="both", expand=True)

        # Quick Presets Frame
        preset_frame = tk.LabelFrame(frame, text=" Raccourcis Pre-configurés ", font=("Segoe UI", 10, "bold"), bg="#f4f6f9")
        preset_frame.pack(fill="x", pady=5)

        btn_p1 = tk.Button(preset_frame, text="PC 1 (CNC Actif)", bg="#d9edf7", font=("Segoe UI", 9, "bold"),
                           command=lambda: self.quick_backup(r"\\PC-CN\usinage\usinage", r"E:\sauvegarde-usinage", "usinage"))
        btn_p1.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        btn_p2 = tk.Button(preset_frame, text="PC 2 (CNC Réserve)", bg="#d9edf7", font=("Segoe UI", 9, "bold"),
                           command=lambda: self.quick_backup(r"\\PC2-CN\usinage2\usinage", r"E:\sauvegarde-usinage", "usinage2"))
        btn_p2.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        btn_p3 = tk.Button(preset_frame, text="Poste Local", bg="#d9edf7", font=("Segoe UI", 9, "bold"),
                           command=lambda: self.quick_backup(r"C:\ESPACE-TRAVAIL", r"F:\SAUVGARDES", "ESPACE-TRAVAIL"))
        btn_p3.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        preset_frame.columnconfigure(0, weight=1)
        preset_frame.columnconfigure(1, weight=1)
        preset_frame.columnconfigure(2, weight=1)

        # Custom Config Frame
        custom_frame = tk.LabelFrame(frame, text=" Configuration sur mesure ", font=("Segoe UI", 10, "bold"), bg="#f4f6f9")
        custom_frame.pack(fill="x", pady=10)

        # Source Selection
        tk.Label(custom_frame, text="Dossier / Fichier Source :", bg="#f4f6f9", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_src = tk.Entry(custom_frame, width=50, font=("Segoe UI", 9))
        self.entry_src.grid(row=0, column=1, padx=5, pady=5)
        self.entry_src.insert(0, r"C:\ESPACE-TRAVAIL")

        btn_browse_src_dir = tk.Button(custom_frame, text="📁 Dossier...", command=self.browse_src_dir)
        btn_browse_src_dir.grid(row=0, column=2, padx=2, pady=5)

        btn_browse_src_file = tk.Button(custom_frame, text="📄 Fichier...", command=self.browse_src_file)
        btn_browse_src_file.grid(row=0, column=3, padx=2, pady=5)

        btn_explore_src = tk.Button(custom_frame, text="👁️ Explorer Source", bg="#fff8dc", font=("Segoe UI", 8, "bold"), command=self.explore_source_folder)
        btn_explore_src.grid(row=0, column=4, padx=5, pady=5)

        # Destination Selection
        tk.Label(custom_frame, text="Dossier Destination :", bg="#f4f6f9", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_dst = tk.Entry(custom_frame, width=50, font=("Segoe UI", 9))
        self.entry_dst.grid(row=1, column=1, padx=5, pady=5)
        self.entry_dst.insert(0, r"F:\SAUVGARDES")

        btn_browse_dst = tk.Button(custom_frame, text="📁 Parcourir...", command=self.browse_dst_dir)
        btn_browse_dst.grid(row=1, column=2, columnspan=2, sticky="ew", padx=2, pady=5)

        # Prefix Name
        tk.Label(custom_frame, text="Prefixe Sauvegarde :", bg="#f4f6f9", font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.entry_prefix = tk.Entry(custom_frame, width=30, font=("Segoe UI", 9))
        self.entry_prefix.grid(row=2, column=1, sticky="w", padx=5, pady=5)
        self.entry_prefix.insert(0, "Manuel")

        # Start Custom Backup Button
        btn_run_custom = tk.Button(custom_frame, text="🚀 Lancer la Sauvegarde Manuel", font=("Segoe UI", 10, "bold"), bg="#5cb85c", fg="white",
                                   command=self.run_custom_backup)
        btn_run_custom.grid(row=3, column=0, columnspan=5, pady=10, sticky="ew", padx=5)

        # Log Window
        log_frame = tk.LabelFrame(frame, text=" Journal des opérations ", font=("Segoe UI", 10, "bold"), bg="#f4f6f9")
        log_frame.pack(fill="both", expand=True, pady=5)

        self.backup_log = tk.Text(log_frame, font=("Consolas", 9), bg="#1e1e1e", fg="#ffffff", insertbackground="white")
        self.backup_log.pack(fill="both", expand=True, padx=5, pady=5)

        # Tag configurations for colors
        self.backup_log.tag_config("info", foreground="#00bfff")
        self.backup_log.tag_config("success", foreground="#00ff00")
        self.backup_log.tag_config("error", foreground="#ff4d4d")

    def browse_src_dir(self):
        d = filedialog.askdirectory(title="Sélectionner le dossier source")
        if d:
            self.entry_src.delete(0, tk.END)
            self.entry_src.insert(0, d)

    def browse_src_file(self):
        f = filedialog.askopenfilename(title="Sélectionner le fichier source")
        if f:
            self.entry_src.delete(0, tk.END)
            self.entry_src.insert(0, f)

    def browse_dst_dir(self):
        d = filedialog.askdirectory(title="Sélectionner le dossier de destination")
        if d:
            self.entry_dst.delete(0, tk.END)
            self.entry_dst.insert(0, d)

    def explore_source_folder(self):
        path = self.entry_src.get().strip()
        if not path:
            messagebox.showwarning("Chemin vide", "Veuillez spécifier un chemin source.")
            return
        
        target = path if os.path.isdir(path) else os.path.dirname(path)
        if os.path.exists(target):
            os.startfile(target)
        else:
            messagebox.showerror("Dossier introuvable", f"Le dossier n'existe pas :\n{target}")

    def quick_backup(self, src, dst, prefix):
        threading.Thread(target=run_robocopy_backup, args=(src, dst, prefix, self.backup_log, self.status_label), daemon=True).start()

    def run_custom_backup(self):
        src = self.entry_src.get().strip()
        dst = self.entry_dst.get().strip()
        prefix = self.entry_prefix.get().strip()
        if not src or not dst:
            messagebox.showwarning("Champs requis", "Veuillez renseigner la source et la destination.")
            return
        threading.Thread(target=run_robocopy_backup, args=(src, dst, prefix, self.backup_log, self.status_label), daemon=True).start()

    # ------------------------------------------------------------------
    # TAB 2 : COMPARAISON
    # ------------------------------------------------------------------
    def build_compare_tab(self):
        frame = tk.Frame(self.tab_compare, bg="#f4f6f9", padx=15, pady=15)
        frame.pack(fill="both", expand=True)

        # Selection Frame
        sel_frame = tk.LabelFrame(frame, text=" Sélection des éléments à comparer ", font=("Segoe UI", 10, "bold"), bg="#f4f6f9")
        sel_frame.pack(fill="x", pady=5)

        # Target A
        tk.Label(sel_frame, text="Élément A (Récent / Référence) :", bg="#f4f6f9").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_comp_a = tk.Entry(sel_frame, width=45, font=("Segoe UI", 9))
        self.entry_comp_a.grid(row=0, column=1, padx=5, pady=5)
        btn_a_dir = tk.Button(sel_frame, text="📁 Dossier", command=lambda: self.browse_compare_path(self.entry_comp_a, is_file=False))
        btn_a_dir.grid(row=0, column=2, padx=2, pady=5)
        btn_a_file = tk.Button(sel_frame, text="📄 Fichier", command=lambda: self.browse_compare_path(self.entry_comp_a, is_file=True))
        btn_a_file.grid(row=0, column=3, padx=2, pady=5)

        # Target B
        tk.Label(sel_frame, text="Élément B (Ancien / Comparé) :", bg="#f4f6f9").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_comp_b = tk.Entry(sel_frame, width=45, font=("Segoe UI", 9))
        self.entry_comp_b.grid(row=1, column=1, padx=5, pady=5)
        btn_b_dir = tk.Button(sel_frame, text="📁 Dossier", command=lambda: self.browse_compare_path(self.entry_comp_b, is_file=False))
        btn_b_dir.grid(row=1, column=2, padx=2, pady=5)
        btn_b_file = tk.Button(sel_frame, text="📄 Fichier", command=lambda: self.browse_compare_path(self.entry_comp_b, is_file=True))
        btn_b_file.grid(row=1, column=3, padx=2, pady=5)

        # Action Buttons
        btn_run_comp = tk.Button(sel_frame, text="🔍 Lancer la Comparaison", font=("Segoe UI", 10, "bold"), bg="#0275d8", fg="white",
                                 command=self.run_comparison)
        btn_run_comp.grid(row=2, column=0, columnspan=4, pady=8, sticky="ew", padx=5)

        # Treeview / Table Results
        res_frame = tk.LabelFrame(frame, text=" Résultats de la comparaison ", font=("Segoe UI", 10, "bold"), bg="#f4f6f9")
        res_frame.pack(fill="both", expand=True, pady=5)

        cols = ("statut", "fichier", "date_a", "date_b", "taille_a", "taille_b")
        self.tree_comp = ttk.Treeview(res_frame, columns=cols, show="headings", height=10)
        self.tree_comp.heading("statut", text="Statut")
        self.tree_comp.heading("fichier", text="Fichier / Chemin Relatif")
        self.tree_comp.heading("date_a", text="Date Modification (A)")
        self.tree_comp.heading("date_b", text="Date Modification (B)")
        self.tree_comp.heading("taille_a", text="Taille A (Octets)")
        self.tree_comp.heading("taille_b", text="Taille B (Octets)")

        self.tree_comp.column("statut", width=110, anchor="center")
        self.tree_comp.column("fichier", width=250, anchor="w")
        self.tree_comp.column("date_a", width=140, anchor="center")
        self.tree_comp.column("date_b", width=140, anchor="center")
        self.tree_comp.column("taille_a", width=100, anchor="e")
        self.tree_comp.column("taille_b", width=100, anchor="e")

        scrollbar = ttk.Scrollbar(res_frame, orient="vertical", command=self.tree_comp.yview)
        self.tree_comp.configure(yscroll=scrollbar.set)
        
        self.tree_comp.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Color tags for Treeview
        self.tree_comp.tag_configure("MODIFIÉ", foreground="#d9534f", background="#f2dede") # Red
        self.tree_comp.tag_configure("AJOUTÉ", foreground="#0275d8", background="#d9edf7")  # Blue
        self.tree_comp.tag_configure("SUPPRIMÉ", foreground="#f0ad4e", background="#fcf8e3") # Yellow/Orange
        self.tree_comp.tag_configure("IDENTIQUE", foreground="#5cb85c")                     # Green

        # Export & Print Buttons Bar
        exp_frame = tk.Frame(frame, bg="#f4f6f9")
        exp_frame.pack(fill="x", pady=5)

        btn_exp_txt = tk.Button(exp_frame, text="💾 Exporter en TXT", font=("Segoe UI", 9), command=self.export_txt)
        btn_exp_txt.pack(side="left", padx=5)

        btn_exp_xls = tk.Button(exp_frame, text="📊 Exporter en Excel (.xlsx)", font=("Segoe UI", 9), command=self.export_excel)
        btn_exp_xls.pack(side="left", padx=5)

        btn_print = tk.Button(exp_frame, text="🖨️ Imprimer / Imprimante System", font=("Segoe UI", 9, "bold"), bg="#6c757d", fg="white",
                              command=self.print_results)
        btn_print.pack(side="right", padx=5)

        self.comparison_data = [] # Stores comparison records

    def browse_compare_path(self, entry_widget, is_file=False):
        if is_file:
            path = filedialog.askopenfilename(title="Sélectionner un fichier")
        else:
            path = filedialog.askdirectory(title="Sélectionner un dossier")
        if path:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, path)

    def run_comparison(self):
        path_a = self.entry_comp_a.get().strip()
        path_b = self.entry_comp_b.get().strip()

        if not path_a or not path_b:
            messagebox.showwarning("Champs requis", "Veuillez sélectionner l'élément A et l'élément B.")
            return

        if not os.path.exists(path_a) or not os.path.exists(path_b):
            messagebox.showerror("Erreur de chemin", "L'un des chemins spécifiés n'existe pas.")
            return

        # Clear existing items
        for item in self.tree_comp.get_children():
            self.tree_comp.delete(item)

        self.comparison_data = []

        # Case 1: Both are Files
        if os.path.isfile(path_a) and os.path.isfile(path_b):
            stat_a, stat_b = os.stat(path_a), os.stat(path_b)
            d_a = datetime.fromtimestamp(stat_a.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            d_b = datetime.fromtimestamp(stat_b.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
            
            status = "IDENTIQUE"
            if stat_a.st_size != stat_b.st_size or abs(stat_a.st_mtime - stat_b.st_mtime) > 2:
                status = "MODIFIÉ"

            rec = {
                "statut": status,
                "fichier": os.path.basename(path_a),
                "date_a": d_a,
                "date_b": d_b,
                "taille_a": stat_a.st_size,
                "taille_b": stat_b.st_size
            }
            self.comparison_data.append(rec)

        # Case 2: Both are Folders
        elif os.path.isdir(path_a) and os.path.isdir(path_b):
            dict_a = {os.path.relpath(os.path.join(r, f), path_a): os.path.join(r, f)
                      for r, _, files in os.walk(path_a) for f in files}
            dict_b = {os.path.relpath(os.path.join(r, f), path_b): os.path.join(r, f)
                      for r, _, files in os.walk(path_b) for f in files}

            all_rel_paths = sorted(list(set(dict_a.keys()).union(set(dict_b.keys()))))

            for rel_p in all_rel_paths:
                in_a = rel_p in dict_a
                in_b = rel_p in dict_b

                if in_a and not in_b:
                    st_a = os.stat(dict_a[rel_p])
                    d_a = datetime.fromtimestamp(st_a.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                    rec = {"statut": "AJOUTÉ", "fichier": rel_p, "date_a": d_a, "date_b": "-", "taille_a": st_a.st_size, "taille_b": "-"}
                elif not in_a and in_b:
                    st_b = os.stat(dict_b[rel_p])
                    d_b = datetime.fromtimestamp(st_b.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                    rec = {"statut": "SUPPRIMÉ", "fichier": rel_p, "date_a": "-", "date_b": d_b, "taille_a": "-", "taille_b": st_b.st_size}
                else:
                    st_a, st_b = os.stat(dict_a[rel_p]), os.stat(dict_b[rel_p])
                    d_a = datetime.fromtimestamp(st_a.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                    d_b = datetime.fromtimestamp(st_b.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                    
                    status = "IDENTIQUE"
                    if st_a.st_size != st_b.st_size or abs(st_a.st_mtime - st_b.st_mtime) > 2:
                        status = "MODIFIÉ"

                    rec = {"statut": status, "fichier": rel_p, "date_a": d_a, "date_b": d_b, "taille_a": st_a.st_size, "taille_b": st_b.st_size}

                self.comparison_data.append(rec)
        else:
            messagebox.showwarning("Incompatibilité", "Veuillez comparer soit deux dossiers, soit deux fichiers de même nature.")
            return

        # Populate Treeview
        for rec in self.comparison_data:
            self.tree_comp.insert("", "end", values=(
                rec["statut"], rec["fichier"], rec["date_a"], rec["date_b"], rec["taille_a"], rec["taille_b"]
            ), tags=(rec["statut"],))

        self.status_label.config(text=f"Comparaison terminée : {len(self.comparison_data)} élément(s) traités.", fg="green")

    def export_txt(self):
        if not self.comparison_data:
            messagebox.showwarning("Aucune donnée", "Aucun résultat de comparaison à exporter.")
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")], title="Enregistrer le rapport TXT")
        if not save_path:
            return

        with open(save_path, "w", encoding="utf-8") as f:
            f.write("========================================================\n")
            f.write("RAPPORT DE COMPARAISON INDUSTRIEL CNC\n")
            f.write(f"Date du rapport : {get_timestamp_readable()}\n")
            f.write(f"Élément A : {self.entry_comp_a.get()}\n")
            f.write(f"Élément B : {self.entry_comp_b.get()}\n")
            f.write("========================================================\n\n")

            for r in self.comparison_data:
                f.write(f"[{r['statut']}] {r['fichier']}\n")
                f.write(f"  Date A: {r['date_a']} | Taille A: {r['taille_a']}\n")
                f.write(f"  Date B: {r['date_b']} | Taille B: {r['taille_b']}\n")
                f.write("-" * 50 + "\n")

        messagebox.showinfo("Exportation réussie", f"Rapport TXT sauvegardé sous :\n{save_path}")

    def export_excel(self):
        if not self.comparison_data:
            messagebox.showwarning("Aucune donnée", "Aucun résultat de comparaison à exporter.")
            return

        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Module manquant", "Le module 'openpyxl' n'est pas disponible pour générer un vrai fichier Excel.")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], title="Enregistrer le rapport Excel")
        if not save_path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Rapport Comparaison"

        # Headers
        headers = ["Statut", "Fichier / Chemin Relatif", "Date Modif (A)", "Date Modif (B)", "Taille A (octets)", "Taille B (octets)"]
        ws.append(headers)

        # Style Header
        header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # Color fills for status
        fills = {
            "MODIFIÉ": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),  # Light Red
            "AJOUTÉ": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),   # Light Green
            "SUPPRIMÉ": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"), # Light Yellow
            "IDENTIQUE": PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        }

        for r in self.comparison_data:
            row_vals = [r["statut"], r["fichier"], r["date_a"], r["date_b"], r["taille_a"], r["taille_b"]]
            ws.append(row_vals)
            
            # Apply color to the status cell
            curr_row = ws.max_row
            status_cell = ws.cell(row=curr_row, column=1)
            if r["statut"] in fills:
                status_cell.fill = fills[r["statut"]]

        # Auto-adjust column width
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        wb.save(save_path)
        messagebox.showinfo("Export Excel", f"Fichier Excel créé avec succès :\n{save_path}")

    def print_results(self):
        if not self.comparison_data:
            messagebox.showwarning("Aucune donnée", "Aucun résultat à imprimer.")
            return

        # Generate temporary print file
        temp_txt = os.path.join(os.environ.get("TEMP", "C:\\Temp"), f"Print_Report_{get_timestamp()}.txt")
        with open(temp_txt, "w", encoding="utf-8") as f:
            f.write("========================================================\n")
            f.write("         RAPPORT DE COMPARAISON CNC - IMPRESSION       \n")
            f.write(f"Date : {get_timestamp_readable()}\n")
            f.write("========================================================\n\n")
            for r in self.comparison_data:
                f.write(f"[{r['statut']}] {r['fichier']}\n")
                f.write(f"   A: {r['date_a']} ({r['taille_a']} octets)\n")
                f.write(f"   B: {r['date_b']} ({r['taille_b']} octets)\n")
                f.write("-" * 50 + "\n")

        if WIN32_PRINT_AVAILABLE and sys.platform.startswith("win"):
            try:
                # Triggers Windows Native Print Dialog
                win32api.ShellExecute(0, "printui", f'/p /n "" "{temp_txt}"', None, ".", 0)
                messagebox.showinfo("Impression", "Le document a été envoyé au gestionnaire d'impression Windows.")
            except Exception as e:
                # Fallback print via notepad dialog
                os.system(f'notepad /p "{temp_txt}"')
        else:
            # Fallback for systems without win32api
            os.system(f'notepad /p "{temp_txt}"')

    # ------------------------------------------------------------------
    # TAB 3 : PLANIFICATION AUTOMATIQUE
    # ------------------------------------------------------------------
    def build_schedule_tab(self):
        frame = tk.Frame(self.tab_schedule, bg="#f4f6f9", padx=15, pady=15)
        frame.pack(fill="both", expand=True)

        sched_box = tk.LabelFrame(frame, text=" Configuration de la Planification Automatique ", font=("Segoe UI", 10, "bold"), bg="#f4f6f9")
        sched_box.pack(fill="x", pady=10)

        # Mode Selection
        tk.Label(sched_box, text="Mode d'exécution :", bg="#f4f6f9").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.combo_mode = ttk.Combobox(sched_box, values=["Quotidien (Chaque jour)", "Hebdomadaire (Une fois / semaine)"], state="readonly", width=35)
        self.combo_mode.grid(row=0, column=1, padx=5, pady=5)
        self.combo_mode.current(0)

        # Time Selection
        tk.Label(sched_box, text="Heure d'exécution (HH:MM) :", bg="#f4f6f9").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.entry_time = tk.Entry(sched_box, width=10, font=("Segoe UI", 10))
        self.entry_time.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        self.entry_time.insert(0, "18:00")

        # Source / Destination Planified
        tk.Label(sched_box, text="Dossier Source à Planifier :", bg="#f4f6f9").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.entry_sched_src = tk.Entry(sched_box, width=45, font=("Segoe UI", 9))
        self.entry_sched_src.grid(row=2, column=1, padx=5, pady=5)
        self.entry_sched_src.insert(0, r"C:\ESPACE-TRAVAIL")

        tk.Label(sched_box, text="Dossier Destination Auto :", bg="#f4f6f9").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.entry_sched_dst = tk.Entry(sched_box, width=45, font=("Segoe UI", 9))
        self.entry_sched_dst.grid(row=3, column=1, padx=5, pady=5)
        self.entry_sched_dst.insert(0, r"F:\SAUVGARDES_AUTO")

        # Toggle Button
        self.btn_toggle_sched = tk.Button(sched_box, text="▶ Activer la Planification", font=("Segoe UI", 10, "bold"), bg="#5cb85c", fg="white",
                                         command=self.toggle_scheduler)
        self.btn_toggle_sched.grid(row=4, column=0, columnspan=2, pady=15, sticky="ew", padx=5)

        # Status Display
        self.lbl_sched_status = tk.Label(frame, text="Planificateur Inactif", font=("Segoe UI", 10, "italic"), bg="#f4f6f9", fg="gray")
        self.lbl_sched_status.pack(pady=10)

    def toggle_scheduler(self):
        if not self.scheduler_running:
            target_time = self.entry_time.get().strip()
            src = self.entry_sched_src.get().strip()
            dst = self.entry_sched_dst.get().strip()

            if not src or not dst:
                messagebox.showwarning("Incomplet", "Veuillez spécifier la source et la destination pour la planification.")
                return

            self.scheduler_running = True
            self.btn_toggle_sched.config(text="⏹ Arrêter la Planification", bg="#d9534f")
            self.lbl_sched_status.config(text=f"Planification ACTIVE (Prochaine exécution programmée à {target_time})", fg="green")
            
            self.scheduler_thread = threading.Thread(target=self.run_scheduler_loop, args=(target_time, src, dst), daemon=True)
            self.scheduler_thread.start()
        else:
            self.scheduler_running = False
            self.btn_toggle_sched.config(text="▶ Activer la Planification", bg="#5cb85c")
            self.lbl_sched_status.config(text="Planificateur Inactif", fg="gray")

    def run_scheduler_loop(self, target_time, src, dst):
        last_run_date = None
        while self.scheduler_running:
            now = datetime.now()
            current_time_str = now.strftime("%H:%M")
            current_date_str = now.strftime("%Y-%m-%d")

            if current_time_str == target_time and last_run_date != current_date_str:
                last_run_date = current_date_str
                run_robocopy_backup(src, dst, "AUTO", self.backup_log, self.status_label)

            time.sleep(30)


if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
