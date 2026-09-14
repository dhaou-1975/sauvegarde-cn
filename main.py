import os
import sys
import shutil
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

MAX_BACKUPS = 5

def get_timestamp():
    return datetime.now().strftime("%Y-%m-%d_%H-%M")

def clean_old_backups(target_root_dir, log_widget):
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
            log_widget.insert(tk.END, f"Nettoyage : suppression -> {os.path.basename(folder)}\n")
            try:
                shutil.rmtree(folder)
            except Exception as e:
                log_widget.insert(tk.END, f"Erreur de suppression : {e}\n")

def execute_backup(source_path, destination_root_dir, folder_prefix, log_widget, status_label):
    log_widget.delete("1.0", tk.END)
    log_widget.insert(tk.END, f"Vérification de la source : {source_path}\n")
    
    if not os.path.exists(source_path):
        messagebox.showerror("Erreur Réseau / CNC", 
                             f"Source inaccessible : {source_path}\n\n"
                             "Vérifiez que le PC CNC est allumé et connecté au réseau.")
        log_widget.insert(tk.END, "[ERREUR] Source inaccessible ou machine déconnectée.\n")
        return

    status_label.config(text="Copie des fichiers en cours...", fg="blue")
    timestamp = get_timestamp()
    folder_name = f"{folder_prefix}_{timestamp}"
    final_destination = os.path.join(destination_root_dir, folder_name)

    log_widget.insert(tk.END, f"Destination : {final_destination}\n")
    os.makedirs(final_destination, exist_ok=True)

    log_widget.insert(tk.END, "Lancement de Robocopy...\n")
    log_widget.update()

    cmd = f'robocopy "{source_path}" "{final_destination}" /E /R:2 /W:3 /NP /NDL'
    subprocess.call(cmd, shell=True)

    clean_old_backups(destination_root_dir, log_widget)
    
    status_label.config(text="Sauvegarde terminée avec succès !", fg="green")
    log_widget.insert(tk.END, "\n[SUCCÈS] Sauvegarde effectuée avec succès.\n")
    messagebox.showinfo("Succès", f"Sauvegarde réussie dans :\n{final_destination}")

def compare_backups(root_dir, log_widget):
    log_widget.delete("1.0", tk.END)
    if not os.path.exists(root_dir):
        messagebox.showwarning("Dossier introuvable", f"Le dossier {root_dir} n'existe pas.")
        return

    subdirs = [
        os.path.join(root_dir, d) for d in os.listdir(root_dir)
        if os.path.isdir(os.path.join(root_dir, d))
    ]
    subdirs.sort(key=lambda x: os.path.getmtime(x), reverse=True)

    if len(subdirs) < 2:
        messagebox.showwarning("Comparaison impossible", 
                               "Il faut au moins 2 dossiers de sauvegarde pour effectuer une comparaison.")
        return

    dir1, dir2 = subdirs[0], subdirs[1]
    
    log_widget.insert(tk.END, f"Comparaison automatique des 2 plus récents :\n")
    log_widget.insert(tk.END, f"A (RÉCENT) : {os.path.basename(dir1)}\n")
    log_widget.insert(tk.END, f"B (ANCIEN) : {os.path.basename(dir2)}\n\n")

    report_name = f"Rapport_Comparaison_{get_timestamp()}.txt"
    report_path = os.path.join(root_dir, report_name)

    dict1 = {os.path.relpath(os.path.join(r, f), dir1): os.path.join(r, f) 
             for r, _, files in os.walk(dir1) for f in files}
    dict2 = {os.path.relpath(os.path.join(r, f), dir2): os.path.join(r, f) 
             for r, _, files in os.walk(dir2) for f in files}

    with open(report_path, "w", encoding="utf-8") as rep:
        rep.write("==================================================\n")
        rep.write("RAPPORT DE COMPARAISON DE SAUVEGARDES\n")
        rep.write(f"Date : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        rep.write(f"Dossier A : {dir1}\n")
        rep.write(f"Dossier B : {dir2}\n")
        rep.write("==================================================\n\n")

        rep.write("--- FICHIERS MODIFIÉS OU AJOUTÉS DANS A ---\n")
        for rel_p, full_p in dict1.items():
            if rel_p not in dict2:
                rep.write(f"[AJOUTÉ DANS A] {rel_p}\n")
                log_widget.insert(tk.END, f"[+] Ajouté: {rel_p}\n")
            else:
                stat1, stat2 = os.stat(full_p), os.stat(dict2[rel_p])
                if stat1.st_size != stat2.st_size or abs(stat1.st_mtime - stat2.st_mtime) > 2:
                    rep.write(f"[MODIFIÉ] {rel_p}\n")
                    log_widget.insert(tk.END, f"[M] Modifié: {rel_p}\n")

        rep.write("\n--- FICHIERS ABSENTS DE A ---\n")
        for rel_p in dict2.keys():
            if rel_p not in dict1:
                rep.write(f"[SUPPRIMÉ DANS A] {rel_p}\n")
                log_widget.insert(tk.END, f"[-] Supprimé: {rel_p}\n")

    log_widget.insert(tk.END, f"\n[SUCCÈS] Rapport généré : {report_path}\n")
    messagebox.showinfo("Rapport généré", f"Rapport sauvegardé sous :\n{report_path}")

class BackupApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestionnaire de Sauvegarde Industrielle CNC")
        self.root.geometry("680x520")
        self.root.configure(bg="#f0f0f0")

        # Titre
        title_frame = tk.Frame(root, bg="#003366", py=10)
        title_frame.pack(fill="x")
        lbl_title = tk.Label(title_frame, text="SAUVEGARDE & COMPARAISON CNC", 
                             font=("Segoe UI", 14, "bold"), fg="white", bg="#003366")
        lbl_title.pack()

        # Cadre des Boutons
        btn_frame = tk.LabelFrame(root, text=" Actions Disponibles ", font=("Segoe UI", 10, "bold"), bg="#f0f0f0", padx=10, pady=10)
        btn_frame.pack(fill="x", padx=15, pady=10)

        # Boutons de Sauvegarde
        btn1 = tk.Button(btn_frame, text=" Sauvegarder PC 1 (CNC Actif)", font=("Segoe UI", 9, "bold"), bg="#e1e1e1",
                         command=lambda: execute_backup(r"\\PC-CN\usinage\usinage", r"E:\sauvegarde-usinage", "usinage", self.log_area, self.status_label))[cite: 4]
        btn1.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        btn2 = tk.Button(btn_frame, text=" Sauvegarder PC 2 (CNC Réserve)", font=("Segoe UI", 9, "bold"), bg="#e1e1e1",
                         command=lambda: execute_backup(r"\\PC2-CN\usinage2\usinage", r"E:\sauvegarde-usinage", "usinage2", self.log_area, self.status_label))[cite: 4]
        btn2.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        btn3 = tk.Button(btn_frame, text=" Sauvegarder Poste Local", font=("Segoe UI", 9, "bold"), bg="#e1e1e1",
                         command=lambda: execute_backup(r"C:\ESPACE-TRAVAIL", r"F:\SAUVGARDES", "ESPACE-TRAVAIL", self.log_area, self.status_label))[cite: 4]
        btn3.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        # Boutons de Comparaison
        btn4 = tk.Button(btn_frame, text=" Comparer Sauvegardes CNC", font=("Segoe UI", 9), bg="#d9edf7",
                         command=lambda: compare_backups(r"E:\sauvegarde-usinage", self.log_area))[cite: 4]
        btn4.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

        btn5 = tk.Button(btn_frame, text=" Comparer Sauvegardes Local", font=("Segoe UI", 9), bg="#d9edf7",
                         command=lambda: compare_backups(r"F:\SAUVGARDES", self.log_area))[cite: 4]
        btn5.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

        btn_frame.columnconfigure(0, weight=1)
        btn_frame.columnconfigure(1, weight=1)

        # Statut
        self.status_label = tk.Label(root, text="Prêt.", font=("Segoe UI", 10, "italic"), bg="#f0f0f0", fg="gray")
        self.status_label.pack(anchor="w", padx=15)

        # Log Text Box
        log_frame = tk.LabelFrame(root, text=" Journal des opérations ", font=("Segoe UI", 10, "bold"), bg="#f0f0f0")
        log_frame.pack(fill="both", expand=True, padx=15, pady=10)

        self.log_area = tk.Text(log_frame, font=("Consolas", 9), bg="#1e1e1e", fg="#00ff00", insertbackground="white")
        self.log_area.pack(fill="both", expand=True, padx=5, pady=5)

if __name__ == "__main__":
    root = tk.Tk()
    app = BackupApp(root)
    root.mainloop()
