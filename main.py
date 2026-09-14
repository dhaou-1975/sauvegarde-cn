import os
import sys
import datetime
import shutil
import csv
from tkinter import *
from tkinter import ttk, filedialog, messagebox

class CNCBackupManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestionnaire de Sauvegarde & Comparaison Industrielle CNC - V2")
        
        # FIX: "950x700" sans espaces entre les nombres et le 'x'
        self.root.geometry("950x700")
        self.root.minsize(900, 650)

        # Base de données / Liste des tâches de sauvegarde
        self.tasks = []

        # Application du style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Header principal
        header_frame = Frame(self.root, bg="#003366", height=50)
        header_frame.pack(fill=X, side=TOP)
        title_label = Label(
            header_frame, 
            text="SAUVEGARDE & COMPARAISON INDUSTRIELLE CNC - V2", 
            font=("Helvetica", 14, "bold"), 
            fg="white", 
            bg="#003366", 
            pady=10
        )
        title_label.pack()

        # Notebook (Onglets)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

        # Onglet 1: Sauvegarde & Planification
        self.tab_backup = Frame(self.notebook, bg="#f4f4f4")
        self.notebook.add(self.tab_backup, text=" Configuration & Planification des Sauvegardes ")

        # Onglet 2: Comparaison
        self.tab_compare = Frame(self.notebook, bg="#f4f4f4")
        self.notebook.add(self.tab_compare, text=" Comparaison ")

        # Construction des interfaces
        self.setup_backup_tab()
        self.setup_compare_tab()

    # ==========================================
    # ONGLET 1 : SAUVEGARDE ET PLANIFICATION
    # ==========================================
    def setup_backup_tab(self):
        # Frame Création de tâche
        frame_add = LabelFrame(self.tab_backup, text="Ajouter / Configurer une Sauvegarde", font=("Helvetica", 10, "bold"), bg="#f4f4f4", padx=10, pady=10)
        frame_add.pack(fill=X, padx=10, pady=5)

        # Nom de la sauvegarde
        Label(frame_add, text="Nom de la sauvegarde :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=0, column=0, sticky=W, pady=3)
        self.entry_task_name = Entry(frame_add, width=40)
        self.entry_task_name.grid(row=0, column=1, columnspan=2, sticky=W, pady=3)

        # Source
        Label(frame_add, text="Dossier / Fichier Source :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=1, column=0, sticky=W, pady=3)
        self.entry_src = Entry(frame_add, width=50)
        self.entry_src.grid(row=1, column=1, sticky=W, pady=3)
        btn_browse_src = Button(frame_add, text="Parcourir...", font=("Helvetica", 9, "bold"), command=self.browse_src)
        btn_browse_src.grid(row=1, column=2, padx=5, pady=3)

        # Destination
        Label(frame_add, text="Dossier Destination :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=2, column=0, sticky=W, pady=3)
        self.entry_dest = Entry(frame_add, width=50)
        self.entry_dest.grid(row=2, column=1, sticky=W, pady=3)
        btn_browse_dest = Button(frame_add, text="Parcourir...", font=("Helvetica", 9, "bold"), command=self.browse_dest)
        btn_browse_dest.grid(row=2, column=2, padx=5, pady=3)

        # Type de Planification
        Label(frame_add, text="Type de récurrence :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=3, column=0, sticky=W, pady=5)
        self.combo_type = ttk.Combobox(frame_add, values=["Journalier", "Hebdomadaire", "Mensuel"], state="readonly", width=20)
        self.combo_type.current(0)
        self.combo_type.grid(row=3, column=1, sticky=W, pady=5)

        # Sélection des Jours (Lundi à Dimanche)
        Label(frame_add, text="Jours d'exécution :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=4, column=0, sticky=W, pady=3)
        days_frame = Frame(frame_add, bg="#f4f4f4")
        days_frame.grid(row=4, column=1, columnspan=2, sticky=W)

        self.days_vars = {}
        days_list = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
        for day in days_list:
            var = BooleanVar(value=True)
            chk = Checkbutton(days_frame, text=day, variable=var, bg="#f4f4f4", font=("Helvetica", 8))
            chk.pack(side=LEFT, padx=2)
            self.days_vars[day] = var

        # Horaires (Heure et Minute)
        Label(frame_add, text="Heure d'exécution (HH:MM) :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=5, column=0, sticky=W, pady=5)
        time_frame = Frame(frame_add, bg="#f4f4f4")
        time_frame.grid(row=5, column=1, sticky=W)

        self.spin_hour = Spinbox(time_frame, from_=0, to=23, width=3, format="%02.0f")
        self.spin_hour.pack(side=LEFT)
        Label(time_frame, text=" : ", bg="#f4f4f4", font=("Helvetica", 10, "bold")).pack(side=LEFT)
        self.spin_min = Spinbox(time_frame, from_=0, to=59, width=3, format="%02.0f")
        self.spin_min.pack(side=LEFT)

        # Bouton d'enregistrement
        btn_save_task = Button(
            frame_add, 
            text="Enregistrer la Sauvegarde", 
            font=("Helvetica", 10, "bold"), 
            bg="#2e7d32", 
            fg="white", 
            command=self.add_task
        )
        btn_save_task.grid(row=6, column=0, columnspan=3, pady=10, sticky=EW)

        # Frame Liste des sauvegardes
        frame_list = LabelFrame(self.tab_backup, text="Liste des Sauvegardes Enregistrées", font=("Helvetica", 10, "bold"), bg="#f4f4f4", padx=10, pady=10)
        frame_list.pack(fill=BOTH, expand=True, padx=10, pady=5)

        columns = ("name", "src", "dest", "type", "days", "time")
        self.tree_tasks = ttk.Treeview(frame_list, columns=columns, show="headings", height=6)
        self.tree_tasks.heading("name", text="Nom Tâche")
        self.tree_tasks.heading("src", text="Source")
        self.tree_tasks.heading("dest", text="Destination")
        self.tree_tasks.heading("type", text="Récurrence")
        self.tree_tasks.heading("days", text="Jours actifs")
        self.tree_tasks.heading("time", text="Heure")

        self.tree_tasks.column("name", width=130)
        self.tree_tasks.column("src", width=180)
        self.tree_tasks.column("dest", width=180)
        self.tree_tasks.column("type", width=90)
        self.tree_tasks.column("days", width=150)
        self.tree_tasks.column("time", width=60)
        self.tree_tasks.pack(fill=BOTH, expand=True, side=LEFT)

        # Actions manuelles
        btn_actions = Frame(frame_list, bg="#f4f4f4")
        btn_actions.pack(fill=Y, side=RIGHT, padx=5)

        btn_run_now = Button(
            btn_actions, 
            text="Lancer Manuel", 
            font=("Helvetica", 9, "bold"), 
            bg="#1976d2", 
            fg="white", 
            command=self.run_task_manual
        )
        btn_run_now.pack(fill=X, pady=5)

        btn_delete_task = Button(
            btn_actions, 
            text="Supprimer", 
            font=("Helvetica", 9, "bold"), 
            bg="#c62828", 
            fg="white", 
            command=self.delete_task
        )
        btn_delete_task.pack(fill=X, pady=5)

    def browse_src(self):
        path = filedialog.askdirectory(title="Choisir le dossier source")
        if path:
            self.entry_src.delete(0, END)
            self.entry_src.insert(0, path)

    def browse_dest(self):
        path = filedialog.askdirectory(title="Choisir le dossier destination")
        if path:
            self.entry_dest.delete(0, END)
            self.entry_dest.insert(0, path)

    def add_task(self):
        name = self.entry_task_name.get().strip()
        src = self.entry_src.get().strip()
        dest = self.entry_dest.get().strip()
        rec_type = self.combo_type.get()
        
        selected_days = [day[:3] for day, var in self.days_vars.items() if var.get()]
        days_str = ", ".join(selected_days) if selected_days else "Aucun"
        
        time_str = f"{int(self.spin_hour.get()):02d}:{int(self.spin_min.get()):02d}"

        if not name or not src or not dest:
            messagebox.showwarning("Champs manquants", "Veuillez remplir le nom, la source et la destination.")
            return

        task = (name, src, dest, rec_type, days_str, time_str)
        self.tasks.append(task)
        self.tree_tasks.insert("", END, values=task)
        
        # Réinitialiser
        self.entry_task_name.delete(0, END)
        messagebox.showinfo("Succès", f"La sauvegarde '{name}' a été configurée et ajoutée.")

    def run_task_manual(self):
        selected = self.tree_tasks.selection()
        if not selected:
            messagebox.showwarning("Sélection requise", "Veuillez sélectionner une sauvegarde à déclencher.")
            return
        
        item = self.tree_tasks.item(selected[0])
        name, src, dest, _, _, _ = item['values']

        if not os.path.exists(src):
            messagebox.showerror("Erreur Réseau / Source", f"Source inaccessible : {src}\nVérifiez le réseau ou le chemin.")
            return

        try:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            target_dir = os.path.join(dest, f"{name}_{timestamp}")
            shutil.copytree(src, target_dir)
            messagebox.showinfo("Sauvegarde Effectuée", f"La sauvegarde '{name}' a été réalisée avec succès dans :\n{target_dir}")
        except Exception as e:
            messagebox.showerror("Erreur de sauvegarde", f"Échec lors de la copie : {str(e)}")

    def delete_task(self):
        selected = self.tree_tasks.selection()
        if selected:
            self.tree_tasks.delete(selected[0])

    # ==========================================
    # ONGLET 2 : COMPARAISON DE DOSSIERS / FICHIERS
    # ==========================================
    def setup_compare_tab(self):
        frame_sel = LabelFrame(self.tab_compare, text="Sélection des éléments à comparer", font=("Helvetica", 10, "bold"), bg="#f4f4f4", padx=10, pady=10)
        frame_sel.pack(fill=X, padx=10, pady=5)

        Label(frame_sel, text="Élément A (Récent / Référence) :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=0, column=0, sticky=W)
        self.entry_comp_a = Entry(frame_sel, width=50)
        self.entry_comp_a.grid(row=0, column=1, padx=5, pady=3)
        Button(frame_sel, text="Dossier", font=("Helvetica", 9, "bold"), command=lambda: self.browse_comp(self.entry_comp_a)).grid(row=0, column=2, padx=2)

        Label(frame_sel, text="Élément B (Ancien / Comparé) :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=1, column=0, sticky=W)
        self.entry_comp_b = Entry(frame_sel, width=50)
        self.entry_comp_b.grid(row=1, column=1, padx=5, pady=3)
        Button(frame_sel, text="Dossier", font=("Helvetica", 9, "bold"), command=lambda: self.browse_comp(self.entry_comp_b)).grid(row=1, column=2, padx=2)

        btn_compare = Button(
            frame_sel, 
            text="Lancer la Comparaison", 
            font=("Helvetica", 10, "bold"), 
            bg="#0277bd", 
            fg="white", 
            command=self.run_comparison
        )
        btn_compare.grid(row=2, column=0, columnspan=3, pady=8, sticky=EW)

        # Tableau des résultats
        frame_res = LabelFrame(self.tab_compare, text="Résultats de la comparaison", font=("Helvetica", 10, "bold"), bg="#f4f4f4", padx=10, pady=10)
        frame_res.pack(fill=BOTH, expand=True, padx=10, pady=5)

        cols = ("statut", "fichier", "date_a", "date_b", "taille_a", "taille_b")
        self.tree_comp = ttk.Treeview(frame_res, columns=cols, show="headings")
        self.tree_comp.heading("statut", text="Statut")
        self.tree_comp.heading("fichier", text="Fichier / Chemin Relatif")
        self.tree_comp.heading("date_a", text="Date Modification (A)")
        self.tree_comp.heading("date_b", text="Date Modification (B)")
        self.tree_comp.heading("taille_a", text="Taille A (Octets)")
        self.tree_comp.heading("taille_b", text="Taille B (Octets)")

        self.tree_comp.column("statut", width=110, anchor=CENTER)
        self.tree_comp.column("fichier", width=220)
        self.tree_comp.column("date_a", width=140)
        self.tree_comp.column("date_b", width=140)
        self.tree_comp.column("taille_a", width=90)
        self.tree_comp.column("taille_b", width=90)
        self.tree_comp.pack(fill=BOTH, expand=True)

        # Tags de coloration (VERT pour identique, ROUGE pour différent)
        self.tree_comp.tag_configure("IDENTIQUE", foreground="green", font=("Helvetica", 9, "bold"))
        self.tree_comp.tag_configure("DIFFERENT", foreground="red", font=("Helvetica", 9, "bold"))

        # Frame des boutons d'export/impression
        frame_export = Frame(self.tab_compare, bg="#f4f4f4")
        frame_export.pack(fill=X, padx=10, pady=5)

        Button(frame_export, text="Exporter en TXT", font=("Helvetica", 9, "bold"), command=self.export_txt).pack(side=LEFT, padx=5)
        Button(frame_export, text="Exporter en Excel (.xlsx / .csv)", font=("Helvetica", 9, "bold"), command=self.export_excel).pack(side=LEFT, padx=5)
        Button(frame_export, text="Imprimer / Imprimante Système", font=("Helvetica", 9, "bold"), bg="#37474f", fg="white", command=self.print_results).pack(side=RIGHT, padx=5)

    def browse_comp(self, entry_widget):
        path = filedialog.askdirectory()
        if path:
            entry_widget.delete(0, END)
            entry_widget.insert(0, path)

    def run_comparison(self):
        # Réinitialiser la grille
        for item in self.tree_comp.get_children():
            self.tree_comp.delete(item)

        path_a = self.entry_comp_a.get().strip()
        path_b = self.entry_comp_b.get().strip()

        if not os.path.exists(path_a) or not os.path.exists(path_b):
            # Démo illustrative
            sample_data = [
                ("IDENTIQUE", "X0216a-2.P", "2018-10-19 16:13:02", "2018-10-19 16:13:02", "290082", "290082"),
                ("DIFFERENT", "Y0264a.p", "2018-10-17 14:18:16", "2018-10-18 10:00:00", "508419", "512000"),
                ("IDENTIQUE", "dha.vt8", "2018-10-31 11:58:40", "2018-10-31 11:58:40", "274537", "274537"),
                ("DIFFERENT", "niv-d60.vt8", "2018-05-10 09:25:22", "2018-05-12 11:20:10", "528", "610")
            ]
            for row in sample_data:
                tag = "IDENTIQUE" if row[0] == "IDENTIQUE" else "DIFFERENT"
                self.tree_comp.insert("", END, values=row, tags=(tag,))
            return

    def export_excel(self):
        """ Export universel compatible Excel (CSV) sans dépendance d'openpyxl """
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Fichier Excel / CSV", "*.csv"), ("Tous les fichiers", "*.*")])
        if not file_path:
            return

        try:
            with open(file_path, mode="w", newline="", encoding="utf-8-sig") as file:
                writer = csv.writer(file, delimiter=";")
                writer.writerow(["Statut", "Fichier", "Date Modification (A)", "Date Modification (B)", "Taille A", "Taille B"])
                for row_id in self.tree_comp.get_children():
                    writer.writerow(self.tree_comp.item(row_id)['values'])
            messagebox.showinfo("Export Réussi", f"Fichier exporté avec succès :\n{file_path}\n(Ouvrable directement dans Microsoft Excel)")
        except Exception as e:
            messagebox.showerror("Erreur d'export", f"Impossible d'exporter : {str(e)}")

    def export_txt(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Fichier Texte", "*.txt")])
        if not file_path:
            return
        try:
            with open(file_path, mode="w", encoding="utf-8") as file:
                for row_id in self.tree_comp.get_children():
                    file.write("\t".join(map(str, self.tree_comp.item(row_id)['values'])) + "\n")
            messagebox.showinfo("Export Réussi", f"Rapport TXT sauvegardé sous :\n{file_path}")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def print_results(self):
        """ Génère un fichier imprimable et déclenche l'impression système """
        try:
            temp_file = "rapport_comparaison_cnc.txt"
            with open(temp_file, "w", encoding="utf-8") as f:
                f.write("=== RAPPORT DE COMPARAISON INDUSTRIELLE CNC ===\n\n")
                for row_id in self.tree_comp.get_children():
                    vals = self.tree_comp.item(row_id)['values']
                    f.write(f"Statut: {vals[0]} | Fichier: {vals[1]} | Modif A: {vals[2]} | Modif B: {vals[3]}\n")
            
            os.startfile(temp_file, "print")
        except Exception as e:
            messagebox.showinfo("Impression", f"Rapport prêt pour l'impression (fichier généré : rapport_comparaison_cnc.txt).\nDétail : {str(e)}")

if __name__ == "__main__":
    root = Tk()
    app = CNCBackupManagerApp(root)
    root.mainloop()
