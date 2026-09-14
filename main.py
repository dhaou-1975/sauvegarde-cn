import os
import sys
import time
import datetime
import shutil
import csv
from tkinter import *
from tkinter import ttk, filedialog, messagebox

# ==========================================
# INFORMATIONS DE L'APPLICATION
# ==========================================
APP_NAME = "Gestionnaire de Sauvegarde des Programmes CNC"
APP_VERSION = "2.1.0"
APP_AUTHOR = "Bouzaien Dhaou"
APP_EMAIL = "bouzaien.dhaou@gmail.com"
DATE_CREATED = "14/09/2026"
DATE_MODIFIED = "14/09/2026"


class BackupProgressBarDialog(Toplevel):
    """ Fenêtre modale avec barre de progression verte pour la sauvegarde """
    def __init__(self, parent, title="Sauvegarde en cours..."):
        super().__init__(parent)
        self.title(title)
        self.geometry("450x200")
        self.resizable(False, False)
        self.grab_set()  # Fenêtre modale
        self.protocol("WM_DELETE_WINDOW", lambda: None)  # Empêche la fermeture prématurée

        self.label_status = Label(self, text="Préparation de la sauvegarde...", font=("Helvetica", 10, "bold"))
        self.label_status.pack(pady=15)

        # Style pour la barre de progression en vert
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure("Green.Horizontal.TProgressbar", foreground='#4CAF50', background='#4CAF50', thickness=20)

        self.progress = ttk.Progressbar(self, style="Green.Horizontal.TProgressbar", length=350, mode='determinate')
        self.progress.pack(pady=10)

        self.label_percent = Label(self, text="0%", font=("Helvetica", 10))
        self.label_percent.pack()

        self.btn_close = Button(self, text="Fermer", font=("Helvetica", 10, "bold"), bg="#2e7d32", fg="white", state=DISABLED, command=self.destroy)
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
        self.btn_close.config(state=NORMAL)
        self.protocol("WM_DELETE_WINDOW", self.destroy)


class CNCBackupManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("980x730")
        self.root.minsize(900, 650)

        self.tasks = []

        # Menu Barre Supérieure
        self.menubar = Menu(self.root)
        self.root.config(menu=self.menubar)

        # Menu Aide / À propos
        self.help_menu = Menu(self.menubar, tearoff=0)
        self.help_menu.add_command(label="À propos", command=self.show_about)
        self.menubar.add_cascade(label="Aide", menu=self.help_menu)

        # Header principal
        header_frame = Frame(self.root, bg="#003366", height=50)
        header_frame.pack(fill=X, side=TOP)
        title_label = Label(
            header_frame, 
            text=APP_NAME.upper(), 
            font=("Helvetica", 14, "bold"), 
            fg="white", 
            bg="#003366", 
            pady=10
        )
        title_label.pack()

        # Onglets
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)

        self.tab_backup = Frame(self.notebook, bg="#f4f4f4")
        self.notebook.add(self.tab_backup, text=" Configuration & Planification des Sauvegardes ")

        self.tab_compare = Frame(self.notebook, bg="#f4f4f4")
        self.notebook.add(self.tab_compare, text=" Comparaison de Dossiers ")

        self.setup_backup_tab()
        self.setup_compare_tab()

    def show_about(self):
        """ Affichage de la boîte À Propos """
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
        frame_add = LabelFrame(self.tab_backup, text="Ajouter / Configurer une Sauvegarde", font=("Helvetica", 10, "bold"), bg="#f4f4f4", padx=10, pady=10)
        frame_add.pack(fill=X, padx=10, pady=5)

        Label(frame_add, text="Nom de la sauvegarde :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=0, column=0, sticky=W, pady=3)
        self.entry_task_name = Entry(frame_add, width=40)
        self.entry_task_name.grid(row=0, column=1, columnspan=2, sticky=W, pady=3)

        Label(frame_add, text="Dossier / Fichier Source :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=1, column=0, sticky=W, pady=3)
        self.entry_src = Entry(frame_add, width=50)
        self.entry_src.grid(row=1, column=1, sticky=W, pady=3)
        Button(frame_add, text="Parcourir...", font=("Helvetica", 9, "bold"), command=self.browse_src).grid(row=1, column=2, padx=5, pady=3)

        Label(frame_add, text="Dossier Destination :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=2, column=0, sticky=W, pady=3)
        self.entry_dest = Entry(frame_add, width=50)
        self.entry_dest.grid(row=2, column=1, sticky=W, pady=3)
        Button(frame_add, text="Parcourir...", font=("Helvetica", 9, "bold"), command=self.browse_dest).grid(row=2, column=2, padx=5, pady=3)

        Label(frame_add, text="Type de récurrence :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=3, column=0, sticky=W, pady=5)
        self.combo_type = ttk.Combobox(frame_add, values=["Journalier", "Hebdomadaire", "Mensuel"], state="readonly", width=20)
        self.combo_type.current(0)
        self.combo_type.grid(row=3, column=1, sticky=W, pady=5)
        self.combo_type.bind("<<ComboboxSelected>>", self.on_recurrence_change)

        # Jours d'exécution hebdomadaire
        self.label_days = Label(frame_add, text="Jours actifs (Hebdo) :", font=("Helvetica", 9, "bold"), bg="#f4f4f4")
        self.label_days.grid(row=4, column=0, sticky=W, pady=3)
        
        self.days_frame = Frame(frame_add, bg="#f4f4f4")
        self.days_frame.grid(row=4, column=1, columnspan=2, sticky=W)

        self.days_vars = {}
        self.days_checkbuttons = []
        for day in ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]:
            var = BooleanVar(value=True)
            chk = Checkbutton(self.days_frame, text=day, variable=var, bg="#f4f4f4", font=("Helvetica", 8))
            chk.pack(side=LEFT, padx=2)
            self.days_vars[day] = var
            self.days_checkbuttons.append(chk)

        # Jour du mois pour le mode Mensuel
        self.label_month_day = Label(frame_add, text="Jour du mois (1-31) :", font=("Helvetica", 9, "bold"), bg="#f4f4f4")
        self.label_month_day.grid(row=5, column=0, sticky=W, pady=5)

        self.spin_month_day = Spinbox(frame_add, from_=1, to=31, width=5, format="%02.0f", state=DISABLED)
        self.spin_month_day.grid(row=5, column=1, sticky=W, pady=5)

        # Heure d'exécution
        Label(frame_add, text="Heure d'exécution (HH:MM) :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=6, column=0, sticky=W, pady=5)
        time_frame = Frame(frame_add, bg="#f4f4f4")
        time_frame.grid(row=6, column=1, sticky=W)

        self.spin_hour = Spinbox(time_frame, from_=0, to=23, width=3, format="%02.0f")
        self.spin_hour.pack(side=LEFT)
        Label(time_frame, text=" : ", bg="#f4f4f4", font=("Helvetica", 10, "bold")).pack(side=LEFT)
        self.spin_min = Spinbox(time_frame, from_=0, to=59, width=3, format="%02.0f")
        self.spin_min.pack(side=LEFT)

        Button(
            frame_add, 
            text="Enregistrer la Sauvegarde", 
            font=("Helvetica", 10, "bold"), 
            bg="#2e7d32", 
            fg="white", 
            command=self.add_task
        ).grid(row=7, column=0, columnspan=3, pady=10, sticky=EW)

        # Liste des sauvegardes
        frame_list = LabelFrame(self.tab_backup, text="Liste des Sauvegardes Enregistrées", font=("Helvetica", 10, "bold"), bg="#f4f4f4", padx=10, pady=10)
        frame_list.pack(fill=BOTH, expand=True, padx=10, pady=5)

        columns = ("name", "src", "dest", "type", "days", "time")
        self.tree_tasks = ttk.Treeview(frame_list, columns=columns, show="headings", height=6)
        self.tree_tasks.heading("name", text="Nom Tâche")
        self.tree_tasks.heading("src", text="Source")
        self.tree_tasks.heading("dest", text="Destination")
        self.tree_tasks.heading("type", text="Récurrence")
        self.tree_tasks.heading("days", text="Planification / Jour")
        self.tree_tasks.heading("time", text="Heure")

        self.tree_tasks.column("name", width=130)
        self.tree_tasks.column("src", width=180)
        self.tree_tasks.column("dest", width=180)
        self.tree_tasks.column("type", width=90)
        self.tree_tasks.column("days", width=150)
        self.tree_tasks.column("time", width=60)
        self.tree_tasks.pack(fill=BOTH, expand=True, side=LEFT)

        btn_actions = Frame(frame_list, bg="#f4f4f4")
        btn_actions.pack(fill=Y, side=RIGHT, padx=5)

        Button(btn_actions, text="Lancer Manuel", font=("Helvetica", 9, "bold"), bg="#1976d2", fg="white", command=self.run_task_manual).pack(fill=X, pady=5)
        Button(btn_actions, text="Supprimer", font=("Helvetica", 9, "bold"), bg="#c62828", fg="white", command=self.delete_task).pack(fill=X, pady=5)

        # État initial selon récurrence
        self.on_recurrence_change()

    def on_recurrence_change(self, event=None):
        rec_type = self.combo_type.get()
        if rec_type == "Mensuel":
            # Activer la case du jour du mois et désactiver la sélection par jour de semaine
            self.spin_month_day.config(state=NORMAL)
            for chk in self.days_checkbuttons:
                chk.config(state=DISABLED)
        elif rec_type == "Hebdomadaire":
            self.spin_month_day.config(state=DISABLED)
            for chk in self.days_checkbuttons:
                chk.config(state=NORMAL)
        else:  # Journalier
            self.spin_month_day.config(state=DISABLED)
            for chk in self.days_checkbuttons:
                chk.config(state=NORMAL)

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
        self.tree_tasks.insert("", END, values=task)
        self.entry_task_name.delete(0, END)
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

        # Lister tous les fichiers à copier pour la barre de progression
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

        # Ouvrir la fenêtre modale de progression
        progress_dialog = BackupProgressBarDialog(self.root, title=f"Sauvegarde : {name}")

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
        frame_sel = LabelFrame(self.tab_compare, text="Sélection des dossiers à comparer", font=("Helvetica", 10, "bold"), bg="#f4f4f4", padx=10, pady=10)
        frame_sel.pack(fill=X, padx=10, pady=5)

        Label(frame_sel, text="Dossier A (Référence) :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=0, column=0, sticky=W)
        self.entry_comp_a = Entry(frame_sel, width=50)
        self.entry_comp_a.grid(row=0, column=1, padx=5, pady=3)
        Button(frame_sel, text="Parcourir", font=("Helvetica", 9, "bold"), command=lambda: self.browse_comp(self.entry_comp_a)).grid(row=0, column=2, padx=2)

        Label(frame_sel, text="Dossier B (Comparé) :", font=("Helvetica", 9, "bold"), bg="#f4f4f4").grid(row=1, column=0, sticky=W)
        self.entry_comp_b = Entry(frame_sel, width=50)
        self.entry_comp_b.grid(row=1, column=1, padx=5, pady=3)
        Button(frame_sel, text="Parcourir", font=("Helvetica", 9, "bold"), command=lambda: self.browse_comp(self.entry_comp_b)).grid(row=1, column=2, padx=2)

        Button(
            frame_sel, 
            text="Lancer la Comparaison", 
            font=("Helvetica", 10, "bold"), 
            bg="#0277bd", 
            fg="white", 
            command=self.run_comparison
        ).grid(row=2, column=0, columnspan=3, pady=8, sticky=EW)

        # Tableau des résultats
        frame_res = LabelFrame(self.tab_compare, text="Résultats de la comparaison", font=("Helvetica", 10, "bold"), bg="#f4f4f4", padx=10, pady=10)
        frame_res.pack(fill=BOTH, expand=True, padx=10, pady=5)

        cols = ("statut", "fichier", "date_a", "date_b")
        self.tree_comp = ttk.Treeview(frame_res, columns=cols, show="headings")
        self.tree_comp.heading("statut", text="Statut")
        self.tree_comp.heading("fichier", text="Fichier / Chemin Relatif")
        self.tree_comp.heading("date_a", text="Date Modification (A)")
        self.tree_comp.heading("date_b", text="Date Modification (B)")

        self.tree_comp.column("statut", width=120, anchor=CENTER)
        self.tree_comp.column("fichier", width=350)
        self.tree_comp.column("date_a", width=180)
        self.tree_comp.column("date_b", width=180)
        self.tree_comp.pack(fill=BOTH, expand=True)

        # Style de coloration des lignes
        self.tree_comp.tag_configure("IDENTIQUE", foreground="green", font=("Helvetica", 9, "bold"))
        self.tree_comp.tag_configure("DIFFERENT", foreground="red", font=("Helvetica", 9, "bold"))

        # Boutons d'exportation / impression
        frame_export = Frame(self.tab_compare, bg="#f4f4f4")
        frame_export.pack(fill=X, padx=10, pady=5)

        Button(frame_export, text="Exporter en TXT", font=("Helvetica", 9, "bold"), command=self.export_txt).pack(side=LEFT, padx=5)
        Button(frame_export, text="Exporter en Excel (.csv)", font=("Helvetica", 9, "bold"), command=self.export_excel).pack(side=LEFT, padx=5)
        Button(frame_export, text="Imprimer (Menu Impression Système)", font=("Helvetica", 9, "bold"), bg="#37474f", fg="white", command=self.print_results).pack(side=RIGHT, padx=5)

    def browse_comp(self, entry_widget):
        path = filedialog.askdirectory()
        if path:
            entry_widget.delete(0, END)
            entry_widget.insert(0, path)

    def run_comparison(self):
        for item in self.tree_comp.get_children():
            self.tree_comp.delete(item)

        dir_a = self.entry_comp_a.get().strip()
        dir_b = self.entry_comp_b.get().strip()

        if not os.path.exists(dir_a) or not os.path.exists(dir_b):
            messagebox.showerror("Erreur Dossier", "Veuillez sélectionner deux dossiers valides à comparer.")
            return

        files_a = {}
        for root_dir, _, files in os.walk(dir_a):
            for f in files:
                full_p = os.path.join(root_dir, f)
                rel_p = os.path.relpath(full_p, dir_a)
                mtime = os.path.getmtime(full_p)
                files_a[rel_p] = (mtime, datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S'))

        files_b = {}
        for root_dir, _, files in os.walk(dir_b):
            for f in files:
                full_p = os.path.join(root_dir, f)
                rel_p = os.path.relpath(full_p, dir_b)
                mtime = os.path.getmtime(full_p)
                files_b[rel_p] = (mtime, datetime.datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S'))

        all_rel_paths = sorted(list(set(files_a.keys()).union(set(files_b.keys()))))

        for rel_p in all_rel_paths:
            info_a = files_a.get(rel_p)
            info_b = files_b.get(rel_p)

            date_a_str = info_a[1] if info_a else "Absent dans A"
            date_b_str = info_b[1] if info_b else "Absent dans B"

            if info_a and info_b:
                if abs(info_a[0] - info_b[0]) < 1.0:
                    status = "IDENTIQUE"
                else:
                    status = "DIFFERENT"
            else:
                status = "DIFFERENT"

            self.tree_comp.insert("", END, values=(status, rel_p, date_a_str, date_b_str), tags=(status,))

    def export_excel(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Fichier CSV (Excel)", "*.csv"), ("Tous les fichiers", "*.*")]
        )
        if not file_path:
            return

        try:
            with open(file_path, mode="w", newline="", encoding="utf-8-sig") as file:
                writer = csv.writer(file, delimiter=";")
                writer.writerow(["Statut", "Fichier / Chemin Relatif", "Date Modification (A)", "Date Modification (B)"])
                for row_id in self.tree_comp.get_children():
                    writer.writerow(self.tree_comp.item(row_id)['values'])
            messagebox.showinfo("Export Réussi", f"Rapport exporté avec succès :\n{file_path}")
        except Exception as e:
            messagebox.showerror("Erreur Export", str(e))

    def export_txt(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Fichier Texte", "*.txt")])
        if not file_path:
            return
        try:
            with open(file_path, mode="w", encoding="utf-8") as file:
                file.write("=== RAPPORT DE COMPARAISON INDUSTRIELLE CNC ===\n\n")
                for row_id in self.tree_comp.get_children():
                    file.write("\t".join(map(str, self.tree_comp.item(row_id)['values'])) + "\n")
            messagebox.showinfo("Export Réussi", f"Rapport TXT sauvegardé sous :\n{file_path}")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def print_results(self):
        """ Impression native sous Windows sans dépendance externe """
        try:
            temp_file = os.path.join(os.environ.get("TEMP", "."), "rapport_comparaison_cnc.txt")
            with open(temp_file, "w", encoding="utf-8") as f:
                f.write(f"=== {APP_NAME.upper()} ===\n")
                f.write(f"Auteur: {APP_AUTHOR} | Version: {APP_VERSION}\n")
                f.write(f"Date du rapport: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                f.write(f"{'STATUT':<12} | {'FICHIER':<40} | {'DATE MODIF (A)':<20} | {'DATE MODIF (B)':<20}\n")
                f.write("-" * 95 + "\n")
                for row_id in self.tree_comp.get_children():
                    v = self.tree_comp.item(row_id)['values']
                    f.write(f"{v[0]:<12} | {v[1]:<40} | {v[2]:<20} | {v[3]:<20}\n")

            if sys.platform == "win32":
                os.startfile(temp_file, "print")
            else:
                messagebox.showinfo("Impression", f"Fichier généré pour impression : {temp_file}")
        except Exception as e:
            messagebox.showerror("Erreur d'impression", f"Impossible de lancer l'impression :\n{str(e)}")


if __name__ == "__main__":
    root = Tk()
    app = CNCBackupManagerApp(root)
    root.mainloop()
