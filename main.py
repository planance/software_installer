import os
import sys
import ctypes
import pythoncom
from win32com.client import Dispatch
import requests
import zipfile
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import customtkinter

dev_name = "NOM_DU_DEV"
software_name = "NOM_DU_LOGICIEL"
program_files_path = os.environ.get("PROGRAMFILES")
dossier_cible = os.path.join(program_files_path, dev_name, software_name)
url_zip = "URL DE VOTRE FICHIER .ZIP AVEC LE FICHIER .EXE DEDANS"

def telecharger_fichier(url, destination):
    response = requests.get(url)
    with open(destination, 'wb') as f:
        f.write(response.content)

def extraire_zip(archive, destination):
    with zipfile.ZipFile(archive, 'r') as zip_ref:
        zip_ref.extractall(destination)

def choisir_destination(destination_actuelle=None):
    if destination_actuelle:
        return destination_actuelle
    else:
        return filedialog.askdirectory()

def mise_a_jour_progression(current, total):
    pourcentage = (current / total) * 100
    progress_bar['value'] = pourcentage
    fenetre.update_idletasks()

def create_shortcut(target_path, shortcut_path):
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.TargetPath = target_path
    shortcut.save()

def trouver_fichier_exe(destination):
    fichiers_dossier = os.listdir(destination)

    exe_files = [fichier for fichier in fichiers_dossier if fichier.lower().endswith('.exe')]

    if len(exe_files) == 1:
        return os.path.join(destination, exe_files[0])
    elif len(exe_files) > 1:
        print("Plus d'un fichier .exe trouvé dans le dossier. Impossible de créer le raccourci.")
    else:
        print("Aucun fichier .exe trouvé dans le dossier. Impossible de créer le raccourci.")
    return None

def lancer_installation(destination):
    destination_zip = "t_e_m_p.zip"
    telecharger_fichier(url_zip, destination_zip)

    extraire_zip(destination_zip, destination)

    total_fichiers = 10
    for i in range(total_fichiers):
        mise_a_jour_progression(i + 1, total_fichiers)

    os.remove(destination_zip)
    label.configure(text="Installation terminée !")
    fermer_fenetre()

    exe_file = trouver_fichier_exe(destination)

    if exe_file:
        bureau = os.path.join(os.path.expanduser("~"), "Desktop")
        shortcut_path = os.path.join(bureau, software_name + ".lnk")
        create_shortcut(exe_file, shortcut_path)

        ctypes.windll.user32.MessageBoxW(0, "Raccourci créé sur le bureau.", "Installation terminée", 1)

def fermer_fenetre():
    destination_entry.pack_forget()
    parcourir_button.pack_forget()
    commencer_button.pack_forget()
    progress_bar.pack_forget()

    fermer_button = customtkinter.CTkButton(fenetre, text="Fermer", command=fenetre.destroy)
    fermer_button.pack()

fenetre = customtkinter.CTk()
fenetre.title("Installation de " + software_name + " ("+ dev_name + ")")

screen_width = fenetre.winfo_screenwidth()
screen_height = fenetre.winfo_screenheight()
window_width = 400
window_height = 200
x_position = (screen_width - window_width) // 2
y_position = (screen_height - window_height) // 2
fenetre.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

fenetre.resizable(False, False)

label = customtkinter.CTkLabel(fenetre, text="Sélectionnez le dossier de destination:")
label.pack()

destination_var = customtkinter.StringVar()
destination_var.set(dossier_cible)

destination_entry = customtkinter.CTkEntry(fenetre, textvariable=destination_var, width=325)
destination_entry.pack()

def parcourir_dossier():
    destination = filedialog.askdirectory()
    if destination:
        destination_var.set(destination)

parcourir_button = customtkinter.CTkButton(fenetre, text="Choisir une autre destination (facultatif)", command=parcourir_dossier)
parcourir_button.pack()

customtkinter.CTkLabel(fenetre, text=" ").pack()

commencer_button = customtkinter.CTkButton(fenetre, text="Commencer l'installation", command=lambda: lancer_installation(destination_var.get()))
commencer_button.pack()

customtkinter.CTkLabel(fenetre, text=" ").pack()

progress_bar = ttk.Progressbar(fenetre, orient="horizontal", length=300, mode="determinate")
progress_bar.pack()

fenetre.mainloop()
