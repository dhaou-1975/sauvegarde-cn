import os
import sys
import shutil
import subprocess
from datetime import datetime

MAX_BACKUPS = 5

def get_timestamp():
    return datetime.now().strftime("%Y-%m-%d_%H-%M")

def clean_old_backups(target_root_dir):
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
            print(f"Nettoyage : suppression de l'ancienne sauvegarde -> {folder}")
            try:
                shutil.rmtree(folder)
            except Exception as e:
                print(f"Erreur de suppression : {e}")

def execute_backup(source_path, destination_root_dir, folder_prefix):
    print("\n--------------------------------------------------")
    print(f"Vérification de la source : {source_path}")
    
    if not os.path.exists(source_path):
        print("[ERREUR] Source inaccessible ou machine CNC déconnectée/éteinte !")
        print("Vérifiez la connexion réseau et la mise sous tension de la CNC.")
        return

    timestamp = get_timestamp()
    folder_name = f"{folder_prefix}_{timestamp}"
    final_destination = os.path.join(destination_root_dir, folder_name)

    print(f"Création du dossier de destination : {final_destination}")
    os.makedirs(final_destination, exist_ok=True)

    print("Copie des fichiers en cours...")
    cmd = f'robocopy "{source_path}" "{final_destination}" /E /R:2 /W:3 /NP /NDL'
    subprocess.call(cmd, shell=True)

    print("[SUCCÈS] Sauvegarde effectuée.")
    clean_old_backups(destination_root_dir)

def compare_backups(root_dir):
    print("\n--- Comparaison de deux dossiers de sauvegarde ---")
    if not os.path.exists(root_dir):
        print(f"[ERREUR] Le dossier {root_dir} n'existe pas.")
        return

    subdirs = [
        os.path.join(root_dir, d) for d in os.listdir(root_dir)
        if os.path.isdir(os.path.join(root_dir, d))
    ]
    subdirs.sort(key=lambda x: os.path.getmtime(x), reverse=True)

    if len(subdirs) < 2:
        print("[AVERTISSEMENT] Au moins 2 sauvegardes sont requises pour effectuer une comparaison.")
        return

    print("Sauvegardes disponibles :")
    for idx, path in enumerate(subdirs):
        print(f" [{idx}] {os.path.basename(path)}")

    try:
        idx1 = int(input("Sélectionnez le numéro du 1er dossier (ex: le plus récent) : "))
        idx2 = int(input("Sélectionnez le numéro du 2nd dossier (ex: le plus ancien) : "))
        dir1 = subdirs[idx1]
        dir2 = subdirs[idx2]
    except (ValueError, IndexError):
        print("[ERREUR] Sélection invalide.")
        return

    report_name = f"Rapport_Comparaison_{get_timestamp()}.txt"
    report_path = os.path.join(root_dir, report_name)

    dict1 = {}
    for root, _, files in os.walk(dir1):
        for f in files:
            full_p = os.path.join(root, f)
            rel_p = os.path.relpath(full_p, dir1)
            dict1[rel_p] = full_p

    dict2 = {}
    for root, _, files in os.walk(dir2):
        for f in files:
            full_p = os.path.join(root, f)
            rel_p = os.path.relpath(full_p, dir2)
            dict2[rel_p] = full_p

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
            else:
                stat1 = os.stat(full_p)
                stat2 = os.stat(dict2[rel_p])
                if stat1.st_size != stat2.st_size or abs(stat1.st_mtime - stat2.st_mtime) > 2:
                    rep.write(f"[MODIFIÉ] {rel_p} (Taille A: {stat1.st_size} | Taille B: {stat2.st_size})\n")

        rep.write("\n--- FICHIERS ABSENTS DE A (PRÉSENTS DANS B) ---\n")
        for rel_p in dict2.keys():
            if rel_p not in dict1:
                rep.write(f"[SUPPRIMÉ DANS A] {rel_p}\n")

    print(f"\n[SUCCÈS] Rapport généré : {report_path}")

def main():
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("==========================================================")
        print("   APPLICATION DE SAUVEGARDE INDUSTRIELLE (WIN 7)        ")
        print("==========================================================")
        print(r" 1. Sauvegarder PC 1 (CNC Actif  : \\PC-CN\usinage\usinage)")[cite: 4]
        print(r" 2. Sauvegarder PC 2 (CNC Réserve: \\PC2-CN\usinage2\usinage)")[cite: 4]
        print(r" 3. Sauvegarder Poste Local     : C:\ESPACE-TRAVAIL")[cite: 4]
        print(" --------------------------------------------------------")
        print(r" 4. Comparer 2 sauvegardes CNC (E:\sauvegarde-usinage)")[cite: 4]
        print(r" 5. Comparer 2 sauvegardes Local (F:\SAUVGARDES)")[cite: 4]
        print(" --------------------------------------------------------")
        print(" 6. Quitter")
        print("==========================================================")
        
        choice = input("Faites votre choix (1-6) : ").strip()

        if choice == '1':
            execute_backup(r"\\PC-CN\usinage\usinage", r"E:\sauvegarde-usinage", "usinage")[cite: 4]
        elif choice == '2':
            execute_backup(r"\\PC2-CN\usinage2\usinage", r"E:\sauvegarde-usinage", "usinage2")[cite: 4]
        elif choice == '3':
            execute_backup(r"C:\ESPACE-TRAVAIL", r"F:\SAUVGARDES", "ESPACE-TRAVAIL")[cite: 4]
        elif choice == '4':
            compare_backups(r"E:\sauvegarde-usinage")[cite: 4]
        elif choice == '5':
            compare_backups(r"F:\SAUVGARDES")[cite: 4]
        elif choice == '6':
            print("Fermeture de l'application...")
            break
        else:
            print("Choix invalide !")

        input("\nAppuyez sur Entrée pour continuer...")

if __name__ == "__main__":
    main()
