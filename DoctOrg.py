import os
import shutil
import pandas as pd
from datetime import datetime
import docx
from PyPDF2 import PdfReader
from PIL import Image, ImageTk, ExifTags
import openpyxl
import pptx
import tkinter as tk
from tkinter import filedialog
from ttkbootstrap.dialogs import Messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.toast import ToastNotification
import threading


import logging

# Configure logging
LOG_FILE = "error_log.txt"
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s"
)


POSSIBLE_EXTENSIONS = ['.docx', '.pdf', '.txt', '.csv', '.jpg', '.jpeg', '.jfif', '.png', '.xlsx', '.xls', '.pptx', '.ppt', '.zip', '.rar']
VERSION = "v1.0.0"

def get_file_metadata(file_path):
    extension = os.path.splitext(file_path)[1].lower()
    if extension == '.docx':
        return get_docx_metadata(file_path)
    elif extension == '.pdf':
        return get_pdf_metadata(file_path)
    elif extension == '.txt':
        return get_txt_metadata(file_path)
    elif extension == '.csv':
        return get_csv_metadata(file_path)
    elif extension in ['.jpg', '.jpeg', '.jfif', '.png']:
        return get_image_metadata(file_path)
    elif extension == '.xlsx':
        return get_xlsx_metadata(file_path)
    elif extension == '.xls':
        return get_xls_metadata(file_path)
    elif extension in ['.pptx', '.ppt']:
        return get_ppt_metadata(file_path)
    elif extension == '.zip':
        return get_zip_metadata(file_path)
    elif extension == '.rar':
        return get_rar_metadata(file_path)
    else:
        return {'Author': 'Unknown'}

def get_docx_metadata(file_path):
    doc = docx.Document(file_path)
    core_properties = doc.core_properties
    return {'Author': core_properties.author}

def get_pdf_metadata(file_path):
    reader = PdfReader(file_path)
    info = reader.metadata
    return {'Author': info.author if info.author else 'Unknown'}

def get_txt_metadata(file_path):
    return {'Author': 'Unknown'}

def get_csv_metadata(file_path):
    return {'Author': 'Unknown'}

def get_image_metadata(file_path):
    image = Image.open(file_path)
    exif_data = image._getexif()
    if exif_data is not None:
        metadata = {ExifTags.TAGS.get(tag): value for tag, value in exif_data.items()}
        return {'Author': metadata.get('Artist', 'Unknown')}
    return {'Author': 'Unknown'}

def get_xlsx_metadata(file_path):
    wb = openpyxl.load_workbook(file_path)
    props = wb.properties
    return {'Author': props.creator}

def get_xls_metadata(file_path):
    '''wb = xlrd.open_workbook(file_path)
    # author name'''
    return {'Author': 'Unknown'}

def get_ppt_metadata(file_path):
    prs = pptx.Presentation(file_path)
    props = prs.core_properties
    return {'Author': props.author}

def get_zip_metadata(file_path):
    '''with zipfile.ZipFile(file_path, 'r') as zip_ref:
        info_list = zip_ref.infolist()
        if info_list:
            return {'Author': 'Unknown'}'''
    return {'Author': 'Unknown'}

def get_rar_metadata(file_path):
    '''with rarfile.RarFile(file_path, 'r') as rar_ref:
        info_list = rar_ref.infolist()
        if info_list:
            return {'Author': 'Unknown'}'''
    return {'Author': 'Unknown'}

def extract_file_info(self, directory, selected_extensions):
    names = []
    file_types = []
    last_modified_dates = []
    paths = []
    creator_names = []

    for root, dirs, files in os.walk(directory):
        files = [f for f in files if not f[0] == '.']
        length = len(files + dirs)
        i = 0
        for name in files + dirs:
            item_path = os.path.join(root.replace('/','\\'), name)
            file_type = os.path.splitext(name)[1].lower()
            if file_type in selected_extensions or 'all' in selected_extensions:
                names.append(name)
                file_types.append(file_type)
                last_modified = os.path.getmtime(item_path)
                last_modified_date = datetime.fromtimestamp(last_modified).strftime('%Y-%m-%d %H:%M:%S')
                last_modified_dates.append(last_modified_date)
                paths.append(item_path)
                metadata = get_file_metadata(item_path)
                author = metadata.get('Author', 'Unknown')
                creator_names.append(author)
            self.show_progress(length, i + 1)

    data = {
        'Nom': names,
        'Type de fichier': file_types,
        'Nom du créateur': creator_names,
        'Date de dernière modification': last_modified_dates,
        'Chemin': paths
        
    }
    self.show_progress(1, 1)
    return pd.DataFrame(data)

def export_to_csv(df, file_name):
    df.to_csv(file_name, index=False, sep=';')

class FileAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Doct'Org - Gestion documentaire")
        self.root.geometry("800x500")
        self.root.resizable(False, False)
        self.style = ttk.Style("cosmo")  # Modern theme
        self.directory = ttk.StringVar()
        self.selected_extensions = ttk.StringVar(value='all')
        self.dataframe = None
        self.moved_files = []
        self.user_id = tk.StringVar()

        self.logo_path = "./doctorg.png"  
        self.logo_image = None
        if os.path.exists(self.logo_path):
            self.logo_image = Image.open(self.logo_path)
            self.logo_image = self.logo_image.resize((80, 80))
            self.logo_photo = ImageTk.PhotoImage(self.logo_image)
            # l'icon de la fenêtre
            self.root.iconphoto(False, self.logo_photo)


        self.create_widgets()

    def create_widgets(self):

        theme_frame = ttk.Frame(self.root, padding=5)
        theme_frame.pack(side=tk.TOP, anchor=tk.NW, padx=10, pady=5)

        ttk.Label(theme_frame, text="Thème:", font=("Helvetica", 10)).pack(side=tk.LEFT, padx=5)
        self.theme_selector = ttk.Combobox(
            theme_frame,
            values=["vapor", "cyborg", "solar", "darkly", "cosmo", "flatly", "pulse", "morph"],
            state="readonly",
            width=15
        )
        self.theme_selector.set(self.style.theme.name)
        self.theme_selector.pack(side=tk.LEFT, padx=5)
        self.theme_selector.bind("<<ComboboxSelected>>", self.change_theme)

        header_frame = ttk.Frame(self.root, padding=10)
        header_frame.pack(fill=X, pady=10)

        if self.logo_image:
            logo_label = ttk.Label(header_frame, image=self.logo_photo)
            logo_label.pack(side=LEFT, padx=10)

        title_label = ttk.Label(
            header_frame,
            text="Doct'Org",
            font=("Helvetica", 20, "bold"),
            bootstyle=PRIMARY
        )
        title_label.pack(side=LEFT, padx=10)

        subtitle_label = ttk.Label(
            header_frame,
            text="Gestion documentaire simplifiée - " + VERSION,
            font=("Helvetica", 10, "italic"),
            bootstyle=SECONDARY
        )
        subtitle_label.pack(side=LEFT, padx=10)

        # Frame pour la sélection du répertoire
        directory_frame = ttk.LabelFrame(self.root, text="Sélection du répertoire", padding=10)
        directory_frame.pack(fill=X, padx=10, pady=5)

        ttk.Label(directory_frame, text="Répertoire:", font=("Helvetica", 10)).grid(row=0, column=0, padx=5, pady=5, sticky=W)
        ttk.Entry(directory_frame, textvariable=self.directory, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(directory_frame, text="Parcourir", bootstyle=PRIMARY, command=self.browse_directory).grid(row=0, column=2, padx=5, pady=5)

        # Frame pour la sélection des types de fichiers
        filetype_frame = ttk.LabelFrame(self.root, text="Types de fichiers", padding=10)
        filetype_frame.pack(fill=X, padx=10, pady=5)

        ttk.Label(filetype_frame, text="Extensions:", font=("Helvetica", 10)).grid(row=0, column=0, padx=5, pady=5, sticky=W)
        ttk.Combobox(filetype_frame, textvariable=self.selected_extensions, values=['all'] + POSSIBLE_EXTENSIONS).grid(row=0, column=1, padx=5, pady=5)

        ttk.Separator(self.root, orient=HORIZONTAL).pack(fill=X, padx=10, pady=10)

        # Frame pour les actions
        action_frame = ttk.LabelFrame(self.root, text="Actions", padding=10)
        action_frame.pack(fill=X, padx=10, pady=5)

        self.retrieve_button = ttk.Button(action_frame, text="Récupérer les métadonnées", bootstyle=SUCCESS, command=self.extract_information)
        self.retrieve_button.grid(row=0, column=0, padx=5, pady=5)
        self.export_button = ttk.Button(action_frame, text="Exporter les informations", bootstyle=INFO, command=self.export_information, state=DISABLED)
        self.export_button.grid(row=0, column=1, padx=5, pady=5)
        self.view_button = ttk.Button(action_frame, text="Visualiser les informations", bootstyle=SECONDARY, command=self.view_information, state=DISABLED)
        self.view_button.grid(row=0, column=2, padx=5, pady=5)

        # Bouton pour générer le rapport
        self.report_button = ttk.Button(action_frame, text="Générer le rapport", bootstyle=WARNING, command=self.generate_report, state=DISABLED)
        self.report_button.grid(row=0, column=3, padx=5, pady=5)

        self.progress_bar = ttk.Progressbar(self.root, orient=HORIZONTAL, length=400, mode='determinate')
        self.progress_bar.pack(pady=1)

        # Label pour le copyright
        copyright_label = ttk.Label(
        self.root,
        text="© 2025 Mohamed BERTE - Doct'Org. Tous droits réservés.",
        font=("Helvetica", 8),
        bootstyle=SECONDARY,
        anchor="center"
        )
        copyright_label.pack(side=tk.BOTTOM, pady=5)
    
    def change_theme(self, event):
        """Change le thème de l'interface."""
        selected_theme = self.theme_selector.get()
        self.style.theme_use(selected_theme)

    def browse_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.directory.set(directory)

    def extract_information(self):
        directory = self.directory.get()
        selected_extensions = self.selected_extensions.get()
        if not directory:
            Messagebox.show_error(title="Erreur", message="Veuillez sélectionner un répertoire")
            return
        
        toast_start = ToastNotification(
            title="Démarrage",
            message="La récupération des informations a démarré...",
            duration=5000, 
            bootstyle=INFO
        )
        toast_start.show_toast()

        self.disable_buttons()

        thread = threading.Thread(target=self._run_extraction, args=(directory, selected_extensions))
        thread.start()
        
    def _run_extraction(self, directory, selected_extensions):
        """Exécute la récupération des informations (tâche longue)."""
        try:
            self.dataframe = extract_file_info(self=self, directory=directory, selected_extensions=[selected_extensions])
            self.root.after(0, self._on_extraction_complete, True)

        except Exception as e:
            logging.error(f"Erreur lors de la récupération des informations : {e}")
            self.root.after(0, self._on_extraction_complete, False)
        finally:
            self.root.after(0, self.enable_buttons)

    def _on_extraction_complete(self, success):
        """Mise à jour de l'interface utilisateur après la récupération."""
        if success:
            if not self.dataframe.empty:
                self.export_button.config(state=tk.NORMAL)
                self.view_button.config(state=tk.NORMAL)

        toast_end = ToastNotification(
            title="Terminée",
            message="La récupération des informations est terminée.",
            duration=5000,
            bootstyle=SUCCESS
        )
        toast_end.show_toast()

    def disable_buttons(self):
        """Désactive les boutons pendant la récupération."""
        self.retrieve_button.config(state=tk.DISABLED)

    def enable_buttons(self):
        """Réactive les boutons après la récupération."""
        self.retrieve_button.config(state=tk.NORMAL)

    def show_progress(self, total, current):
        progress = int((current / total) * 100)
        self.progress_bar["value"] = progress
        self.root.update_idletasks()

    def export_information(self):
        if self.dataframe is not None:
            file_name = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
            if file_name:
                export_to_csv(self.dataframe, file_name)
                toast = ToastNotification(
                    title="Succès",
                    message=f"Informations extraites et enregistrées dans {file_name}",
                    duration=3000,
                    bootstyle=SUCCESS
                )
                toast.show_toast()
            else:
                toast = ToastNotification(
                    title="Annulé",
                    message="Exportation annulée.",
                    duration=3000,
                    bootstyle=WARNING
                )
                toast.show_toast()
        else:
            toast = ToastNotification(
                title="Erreur",
                message="Aucune donnée à exporter.",
                duration=3000,
                bootstyle=DANGER
            )
            toast.show_toast()

    def revert_last_move(self):
        """Annule le dernier déplacement effectué."""
        if not self.moved_files:
            toast = ToastNotification(
                title="Avertissement",
                message="Aucun déplacement à annuler.",
                duration=3000,
                bootstyle=WARNING
            )
            toast.show_toast()
            return

        # Récupérer les informations du dernier fichier déplacé
        last_move = self.moved_files.pop()  # Supprime et récupère le dernier élément
        source_path = last_move['Chemin source']
        destination_path = last_move['Chemin destination']

        try:
            
            shutil.move(destination_path, source_path)

            self.source_tree.insert(
                "",
                "end",
                values=[
                    last_move['Nom'],
                    last_move['Type de fichier'],
                    last_move['Nom du créateur'],
                    last_move['Date de dernière modification'],
                    source_path
                ]
            )

            for item in self.destination_tree.get_children():
                if self.destination_tree.item(item, 'values')[-1] == destination_path:
                    self.destination_tree.delete(item)
                    break

            new_row = pd.DataFrame([{
                'Nom': last_move['Nom'],
                'Type de fichier': last_move['Type de fichier'],
                'Nom du créateur': last_move['Nom du créateur'],
                'Date de dernière modification': last_move['Date de dernière modification'],
                'Chemin': source_path
            }])
            self.dataframe = pd.concat([self.dataframe, new_row], ignore_index=True)

            toast = ToastNotification(
                title="Succès",
                message=f"Le déplacement de '{last_move['Nom']}' a été annulé.",
                duration=3000,
                bootstyle=SUCCESS
            )
            toast.show_toast()

        except Exception as e:
            # Gérer les erreurs et afficher une notification d'erreur
            toast = ToastNotification(
                title="Erreur",
                message=f"Erreur lors de l'annulation du déplacement : {e}",
                duration=3000,
                bootstyle=DANGER
            )
            toast.show_toast()

    def view_information(self):
        if self.dataframe is not None:
            # Création de la fenêtre
            view_window = ttk.Toplevel(self.root)
            view_window.title("Gestion des fichiers - Source et Destination")
            view_window.geometry("1700x600")

            # Frame principale divisée en deux colonnes
            main_frame = ttk.Frame(view_window)
            main_frame.pack(fill=BOTH, expand=True)

            # Zone de gauche : Répertoire source
            source_frame = ttk.LabelFrame(main_frame, text="Répertoire Source", padding=10)
            source_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)

            self.source_tree = ttk.Treeview(source_frame, bootstyle=INFO, selectmode='extended')
            self.source_tree.pack(expand=True, fill=BOTH)

            self.source_tree["columns"] = list(self.dataframe.columns)
            self.source_tree["show"] = "headings"

            for column in self.source_tree["columns"]:
                self.source_tree.heading(column, text=column, command=lambda c=column: self.sort_treeview(c, False))
                self.source_tree.column(column, width=150)

            for index, row in self.dataframe.iterrows():
                self.source_tree.insert("", "end", values=list(row))

            # Barre de recherche
            search_frame = ttk.Frame(source_frame, padding=5)
            search_frame.pack(fill=tk.X, padx=5, pady=5)

            ttk.Label(search_frame, text="Rechercher:", font=("Helvetica", 10)).pack(side=tk.LEFT, padx=5)
            self.search_var = ttk.StringVar()
            search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
            search_entry.pack(side=tk.LEFT, padx=5)
            search_entry.bind("<KeyRelease>", self.filter_treeview)

            ttk.Button(search_frame, text="Réinitialiser", bootstyle=INFO, command=self.reset_treeview_filter).pack(side=tk.LEFT, padx=5)

            # Zone de droite : Répertoire de destination
            destination_frame = ttk.LabelFrame(main_frame, text="Répertoire Destination", padding=10)
            destination_frame.pack(side=RIGHT, fill=BOTH, expand=True, padx=5, pady=5)

            # Affichage des dossiers et fichiers du répertoire de destination
            self.destination_tree = ttk.Treeview(destination_frame, bootstyle=SUCCESS, selectmode='browse')
            self.destination_tree.pack(expand=True, fill=BOTH)

            self.destination_tree["columns"] = ["Nom", "Type", "Chemin"]
            self.destination_tree["show"] = "headings"

            self.destination_tree.heading("Nom", text="Nom")
            self.destination_tree.column("Nom", width=200)
            self.destination_tree.heading("Type", text="Type")
            self.destination_tree.column("Type", width=100)
            self.destination_tree.heading("Chemin", text="Chemin")
            self.destination_tree.column("Chemin", width=400)

            # Boutons pour charger le répertoire de destination et effectuer des actions
            destination_controls = ttk.Frame(destination_frame)
            destination_controls.pack(fill=X, pady=5)

            ttk.Button(destination_controls, text="Charger Répertoire", bootstyle=PRIMARY, command=self.browse_destination).pack(side=LEFT, padx=5)
            ttk.Button(destination_controls, text="Créer un Dossier", bootstyle=INFO, command=self.create_new_folder).pack(side=LEFT, padx=5)
            
            self.destination_tree.bind("<Double-1>", self.on_destination_double_click)
            self.source_tree.bind("<Double-1>", self.open_file_on_double_click)

            # Boutons pour déplacer les fichiers
            action_controls = ttk.Frame(main_frame)
            action_controls.pack(side=LEFT, padx=10, pady=10)

            ttk.Button(action_controls, text="Déplacer >", bootstyle=SUCCESS, command=self.move_selected_files).pack(pady=5)
            ttk.Button(action_controls, text="< Annuler", bootstyle=WARNING, command=self.revert_last_move).pack(pady=5)
        else:
            Messagebox.show_warning(title="Avertissement", message="Aucune donnée à afficher.")

    def filter_treeview(self, event=None):
        """Filtre les fichiers affichés dans le Treeview en fonction du mot-clé."""
        keyword = self.search_var.get().lower()
        for item in self.source_tree.get_children():
            values = self.source_tree.item(item, "values")
            if any(keyword in str(value).lower() for value in values):
                self.source_tree.item(item, tags=())
            else:
                self.source_tree.item(item, tags=("hidden",))                           

        self.source_tree.tag_configure("hidden", foreground="gray")

    def reset_treeview_filter(self):
        """Réinitialise le filtre du Treeview."""
        self.search_var.set("")
        for item in self.source_tree.get_children():
            self.source_tree.item(item, tags=())
    
    def browse_destination(self):
        directory = filedialog.askdirectory()
        if directory:
            self.destination_directory = directory
            self.populate_destination_tree(directory)

    def populate_destination_tree(self, directory):
        """Remplit l'arborescence de destination avec les dossiers et fichiers."""
        self.current_destination_directory = directory  # Stocke le répertoire actuel
        for item in self.destination_tree.get_children():
            self.destination_tree.delete(item)

        # Ajouter une option pour remonter au dossier parent
        parent_directory = os.path.dirname(directory)
        if parent_directory and parent_directory != directory:
            self.destination_tree.insert("", "end", values=["..", "Dossier", parent_directory])

        for root, dirs, files in os.walk(directory):
            for d in dirs:
                self.destination_tree.insert("", "end", values=[d, "Dossier", os.path.join(root, d)])
            for f in files:
                self.destination_tree.insert("", "end", values=[f, "Fichier", os.path.join(root, f)])
            break  # Ne pas descendre dans les sous-dossiers
    
    def on_destination_double_click(self, event):
        """Navigue dans les dossiers de destination en double-cliquant."""
        selected_item = self.destination_tree.selection()
        if selected_item:
            selected_path = self.destination_tree.item(selected_item[0], 'values')[-1]
            if os.path.isdir(selected_path):  # Si c'est un dossier, naviguer dedans
                self.populate_destination_tree(selected_path)

    def create_new_folder(self):
        """Crée un nouveau dossier dans le répertoire courant de destination."""
        if not hasattr(self, 'current_destination_directory') or not self.current_destination_directory:
            Messagebox.show_error(title="Erreur", message="Veuillez charger un répertoire de destination.", parent=self.root)
            return

        folder_name = tk.simpledialog.askstring("Nouveau Dossier", "Entrez le nom du nouveau dossier :")
        if folder_name:
            new_folder_path = os.path.join(self.current_destination_directory, folder_name)
            try:
                os.makedirs(new_folder_path, exist_ok=True)
                self.populate_destination_tree(self.current_destination_directory)
                toast = ToastNotification(
                title="Succès",
                message=f"Dossier '{folder_name}' créé avec succès.",
                duration=3000,
                bootstyle=SUCCESS
                )
                toast.show_toast()
            except Exception as e:
                toast = ToastNotification(
                title="Erreur",
                message=f"Impossible de créer le dossier : {e}",
                duration=3000,
                bootstyle=DANGER
                )
                toast.show_toast()

    def move_selected_files(self):
        """Déplace les fichiers sélectionnés du répertoire source vers le dossier sélectionné ou la racine."""
        if not hasattr(self, 'destination_directory') or not self.destination_directory:
            #Messagebox.show_error(title="Erreur", message="Veuillez charger un répertoire de destination.", parent=self.root)
            toast = ToastNotification(
            title="Erreur",
            message="Veuillez charger un répertoire de destination.",
            duration=3000,
            bootstyle=DANGER
            )
            return

        selected_items = self.source_tree.selection()
        if not selected_items:
            toast = ToastNotification(
            title="Avertissement",
            message="Aucun fichier sélectionné.",
            duration=3000,
            bootstyle=WARNING
            )
            toast.show_toast()
            return

        selected_dest = self.destination_tree.selection()
        if selected_dest:
            dest_path = self.destination_tree.item(selected_dest[0], 'values')[-1]
            if not os.path.isdir(dest_path):
                #Messagebox.show_error(title="Erreur", message="Veuillez sélectionner un dossier valide.", parent=self.root)
                toast = ToastNotification(
                title="Erreur",
                message="Veuillez sélectionner un dossier valide.",
                duration=3000,
                bootstyle=DANGER
                )
                return
        else:
            dest_path = self.current_destination_directory  # Si rien n'est sélectionné, utiliser la racine

        for item in selected_items:
            file_path = self.source_tree.item(item, 'values')[-1]
            file_name = os.path.basename(file_path)
            destination = os.path.join(dest_path, file_name)

            print(file_path)
            print(destination)

            try:
                shutil.move(file_path, destination)
                
                # Ajouter les informations des fichiers déplacés
                self.moved_files.append({
                    'Nom': self.source_tree.item(item, 'values')[0],
                    'Type de fichier': self.source_tree.item(item, 'values')[1],
                    'Nom du créateur': self.source_tree.item(item, 'values')[2],
                    'Date de dernière modification': self.source_tree.item(item, 'values')[3],
                    'Chemin source': file_path,
                    'Chemin destination': destination
                })

                
                self.report_button.config(state=NORMAL)
                self.source_tree.delete(item)
                self.dataframe = self.dataframe[self.dataframe['Chemin'] != file_path]
                self.populate_destination_tree(self.current_destination_directory)

                toast = ToastNotification(
                title="Succès",
                message=f"Fichiers déplacés vers {destination}",
                duration=3000,
                bootstyle=SUCCESS
                )
                toast.show_toast()

            except Exception as e:
                #Messagebox.show_error(title="Erreur", message=f"Erreur lors du déplacement de {file_name} : {e}")
                toast = ToastNotification(
                title="Erreur",
                message=f"Erreur lors du déplacement de {file_name} : {e}",
                duration=3000,
                bootstyle=DANGER
                )

        #Messagebox.show_info(title="Succès", message="Fichiers déplacés avec succès.")
        toast=ToastNotification(
            title="Succès",
            message="Fichiers déplacés avec succès.",
            duration=3000,
            bootstyle=SUCCESS
        )
    
    def sort_treeview(self, col, reverse):
        """Trie les colonnes du Treeview."""
        data = [(self.source_tree.set(k, col), k) for k in self.source_tree.get_children("")]
        data.sort(reverse=reverse)

        for index, (val, k) in enumerate(data):
            self.source_tree.move(k, "", index)

        self.source_tree.heading(col, command=lambda: self.sort_treeview(col, not reverse))

    def open_file_on_double_click(self, event):
        """Ouvre un fichier en double-cliquant sur une ligne du Treeview."""
        selected_item = self.source_tree.selection()
        if selected_item:
            file_path = self.source_tree.item(selected_item[0], 'values')[-1]
            if os.path.exists(file_path):
                os.startfile(file_path)  # Ouvre le fichier avec l'application par défaut
            else:
                Messagebox.show_error(title="Erreur", message=f"Le fichier '{file_path}' n'existe pas.")

    def generate_report(self):
        if not self.moved_files:
            Messagebox.show_error(title="Erreur", message="Aucun fichier déplacé pour générer un rapport")
            return

        user_id = tk.simpledialog.askstring("Identifiant requis", "Veuillez entrer votre identifiant (ETXXXXX):")
        if not user_id or not user_id.startswith("ET"):
            Messagebox.show_error(title="Erreur", message="Identifiant invalide. Veuillez entrer un identifiant valide (ETXXXXX).")
            return

        report_df = pd.DataFrame(self.moved_files)

        report_df['Responsable'] = user_id

        file_name = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_name:
            report_df.to_csv(file_name, index=False, sep=';')
            Messagebox.show_info(title="Succès", message=f"Rapport généré et enregistré dans {file_name}")

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = FileAnalyzerApp(root)
        root.mainloop()

    except Exception as e:
        logging.error(f"Error DETECTED: {e}")
