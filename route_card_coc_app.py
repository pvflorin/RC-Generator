# =============================================================
# 1. IMPORTS & CONFIGURATION
# =============================================================

# Standard library imports
import sys
import os
import json
import logging
import re

# Windows-specific imports
if sys.platform.startswith('win'):
    import winreg
import argparse
import webbrowser
import subprocess
import platform
from datetime import datetime
from pathlib import Path

# Third-party imports
import pandas as pd  # For reading Excel files
import xlsxwriter   # For writing Excel files
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QListWidgetItem, QLineEdit,
    QMessageBox, QStatusBar, QFileDialog, QTabWidget, QGroupBox,
    QGridLayout, QRadioButton, QInputDialog
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPalette, QColor, QFont

# Windows-specific imports
if sys.platform.startswith('win'):
    import winreg

# =============================================================
# 2. CONSTANTS AND CONFIGURATION
# =============================================================

# Default file names and initial paths
NUME_FISIER_COMENZI = ""
NUME_FISIER_TEHNOLOGII = ""
APP_NAME = "RC Generator"

# Path configuration
if sys.platform.startswith('win'):
    REGISTRY_PATH = r"Software\RCGenApp"
    # Windows standard paths
    CONFIG_DIR = os.path.join(os.getenv('APPDATA', ''), APP_NAME)
    WORK_DIR = os.path.join(os.getenv('USERPROFILE', ''), 'Documents', APP_NAME)
    # Fallbacks if environment variables are not available
    if not CONFIG_DIR or CONFIG_DIR == APP_NAME:
        CONFIG_DIR = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming', APP_NAME)
    if not WORK_DIR or WORK_DIR == APP_NAME:
        WORK_DIR = os.path.join(os.path.expanduser('~'), 'Documents', APP_NAME)
else:
    CONFIG_DIR = os.path.join(os.path.expanduser('~'), 'Library', 'Application Support', APP_NAME)
    WORK_DIR = os.path.join(os.path.expanduser('~'), 'Documents', APP_NAME)

# Create necessary directories with proper permissions
def ensure_dir(path):
    """Create directory with proper permissions if it doesn't exist"""
    if not os.path.exists(path):
        try:
            os.makedirs(path, exist_ok=True)
            if sys.platform.startswith('win'):
                import subprocess
                # Set proper Windows ACL permissions (requires running as administrator first time)
                try:
                    subprocess.run(['icacls', path, '/grant', f'{os.getenv("USERNAME")}:(OI)(CI)F'], 
                                 capture_output=True, text=True, check=True)
                except Exception as e:
                    logging.warning(f"Could not set Windows permissions (non-critical): {e}")
        except Exception as e:
            logging.error(f"Could not create directory {path}: {e}")
            return False
    return True

# Create application directories
ensure_dir(CONFIG_DIR)
ensure_dir(WORK_DIR)

CONFIG_FILE = os.path.join(CONFIG_DIR, 'config.json')

# Create config directory if it doesn't exist
os.makedirs(CONFIG_DIR, exist_ok=True)

# Logger configuration
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# =============================================================
# 2.5 UI CONTRAST FIX FOR WINDOWS
# =============================================================

def ensure_ui_contrast():
    """Ensure readable text on light background (fix white-on-white on Windows)."""
    import sys, os
    try:
        # Prefer Fusion style on Windows (more consistent)
        if sys.platform == "win32":
            os.environ.setdefault("QT_STYLE_OVERRIDE", "Fusion")
        from PyQt6 import QtWidgets, QtGui
        app = QtWidgets.QApplication.instance()
        created_app = False
        if app is None:
            app = QtWidgets.QApplication([])
            created_app = True
        p = QtGui.QPalette()
        p.setColor(QtGui.QPalette.ColorRole.Window, QtGui.QColor(240, 240, 240))
        p.setColor(QtGui.QPalette.ColorRole.WindowText, QtGui.QColor(0, 0, 0))
        p.setColor(QtGui.QPalette.ColorRole.Base, QtGui.QColor(255, 255, 255))
        p.setColor(QtGui.QPalette.ColorRole.Text, QtGui.QColor(0, 0, 0))
        p.setColor(QtGui.QPalette.ColorRole.Button, QtGui.QColor(240, 240, 240))
        p.setColor(QtGui.QPalette.ColorRole.ButtonText, QtGui.QColor(0, 0, 0))
        app.setPalette(p)
        # Fallback: explicitly force widget text color
        app.setStyleSheet("QLabel, QPushButton, QLineEdit, QTextEdit, QComboBox { color: #000000; }")
        if created_app:
            # we created a temporary app just to set the global palette; close it
            app.quit()
    except Exception:
        # Fail silently if PyQt6 not available or something else fails
        pass

# =============================================================
# 3. FILE HANDLING FUNCTIONS
# =============================================================

def get_saved_file_path(key_name, default_value):
    """Get saved file path from config file or registry"""
    if sys.platform.startswith('win'):
        try:
            import winreg
            # Try to open or create the registry key with full permissions
            try:
                key = winreg.HKEY_CURRENT_USER  # type: ignore
                access_key = winreg.OpenKey(key, REGISTRY_PATH, 0, winreg.KEY_READ)  # type: ignore
            except OSError:
                # Key doesn't exist, try to create it
                try:
                    access_key = winreg.CreateKey(key, REGISTRY_PATH)  # type: ignore
                except Exception as e:
                    logging.error(f"Could not create registry key: {e}")
                    return default_value
                
            try:
                value, _ = winreg.QueryValueEx(access_key, key_name)  # type: ignore
                winreg.CloseKey(access_key)  # type: ignore
                return value
            except OSError:
                winreg.CloseKey(access_key)  # type: ignore
                return default_value
        except Exception as e:
            logging.error(f"Registry access error: {e}")
            return default_value
    else:
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    return config.get(key_name, default_value)
        except Exception:
            pass
    return default_value

def save_file_path(key_name, file_path):
    """Save file path to config file or registry"""
    if sys.platform.startswith('win'):
        try:
            import winreg
            key = winreg.HKEY_CURRENT_USER  # type: ignore
            try:
                # Try to open existing key with write permission
                regkey = winreg.OpenKey(key, REGISTRY_PATH, 0, winreg.KEY_WRITE)  # type: ignore
            except OSError:
                # Key doesn't exist, create it
                try:
                    regkey = winreg.CreateKey(key, REGISTRY_PATH)  # type: ignore
                except Exception as e:
                    logging.error(f"Could not create registry key: {e}")
                    return
            
            try:
                winreg.SetValueEx(regkey, key_name, 0, winreg.REG_SZ, file_path)  # type: ignore
            except Exception as e:
                logging.error(f"Could not write to registry: {e}")
            finally:
                winreg.CloseKey(regkey)  # type: ignore
        except Exception as e:
            logging.error(f"Could not save to registry: {e}")
            # Fall back to config file on registry failure
            try:
                config = {}
                if os.path.exists(CONFIG_FILE):
                    with open(CONFIG_FILE, 'r') as f:
                        config = json.load(f)
                config[key_name] = file_path
                with open(CONFIG_FILE, 'w') as f:
                    json.dump(config, f)
            except Exception as e:
                logging.error(f"Could not save to config file: {e}")
    else:
        try:
            config = {}
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
            config[key_name] = file_path
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f)
        except Exception:
            pass

# =============================================================
# 3. FILE HANDLING FUNCTIONS
# =============================================================

def verifica_si_selecteaza_fisier(nume_fisier, tip_fisier):
    # On Windows, check registry for last used file path
    key_name = "comenzi_path" if tip_fisier == "comenzi" else "tehnologii_path"
    file_path = nume_fisier
    saved_path = get_saved_file_path(key_name, nume_fisier)
    
    if os.path.exists(saved_path):
        return saved_path
    
    if not os.path.exists(file_path):
        # Try to ask the user via a GUI dialog
        try:
            from PyQt6.QtWidgets import QApplication, QMessageBox, QFileDialog
            from PyQt6.QtCore import Qt
            
            # Ensure we have a QApplication instance
            app = QApplication.instance()
            if app is None:
                app = QApplication([])
            
            reply = QMessageBox.question(
                None, 
                "Fișier lipsă", 
                f"Fișierul '{file_path}' nu a fost găsit. Doriți să selectați manual fișierul de {tip_fisier}?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                browsed_path, _ = QFileDialog.getOpenFileName(
                    None,
                    f"Selectați fișierul de {tip_fisier}",
                    "",
                    "Excel Files (*.xlsx)"
                )
                if browsed_path:
                    save_file_path(key_name, browsed_path)
                    return browsed_path
        except Exception:
            # No GUI available; cannot prompt user
            return None
        return None
    else:
        save_file_path(key_name, file_path)
        return file_path

def actualizeaza_cale_fisier(nume_fisier):
    # Return absolute path if file exists in current directory, else just the name
    if os.path.exists(nume_fisier):
        return os.path.abspath(nume_fisier)
    return nume_fisier

def select_output_directory(output_path_widget):
    """Function to select output directory using a folder dialog"""
    global WORK_DIR
    current_dir = output_path_widget.text() or WORK_DIR
    folder = QFileDialog.getExistingDirectory(
        None, 
        "Selectați directorul de salvare", 
        current_dir
    )
    if folder:
        output_path_widget.setText(folder)
        save_file_path('output_dir', folder)
        # Update the global WORK_DIR variable
        WORK_DIR = folder

# =============================================================
# 2. DATA ACCESS FUNCTIONS
# =============================================================

# =============================================================
# 4. DATA ACCESS FUNCTIONS
# =============================================================

def gaseste_detalii_comanda(comanda_interna):
    """
    Reads the Excel file for orders and returns the details for the given internal order.
    Returns a dictionary with details if found, or an error message.
    """
    
    global NUME_FISIER_COMENZI
    NUME_FISIER_COMENZI = verifica_si_selecteaza_fisier(NUME_FISIER_COMENZI, "comenzi") or NUME_FISIER_COMENZI
    NUME_FISIER_COMENZI = actualizeaza_cale_fisier(NUME_FISIER_COMENZI)
    if not os.path.exists(NUME_FISIER_COMENZI):
        return None, f"Fișierul '{NUME_FISIER_COMENZI}' nu a fost găsit sau selectat."

    try:
        # Citim foaia 'Comenzi', sărind peste rândul 0 (dacă există un rând de totaluri)
        df = pd.read_excel(NUME_FISIER_COMENZI, sheet_name='Comenzi', skiprows=[0]) 
        
        # 1. Curățarea spațiilor albe din antetele coloanelor
        df.columns = df.columns.str.strip() 
        
        # 2. Curățarea spațiilor albe din coloana cheie și asigurarea tipului string
        if 'Comanda Interna' in df.columns:
            df['Comanda Interna'] = df['Comanda Interna'].fillna('').astype(str).str.strip()
        else:
            return None, "Coloana 'Comanda Interna' nu a fost găsită în foaia 'Comenzi' din fișierul sursă."

        # Cautam comanda
        rezultat = df[df['Comanda Interna'] == str(comanda_interna)]

        if not rezultat.empty:
            return rezultat.iloc[0].to_dict(), None
        else:
            return None, f"Comanda Internă '{comanda_interna}' nu a fost găsită în fișier."

    except Exception as e:
        return None, f"Eroare la citirea/procesarea fișierului Excel pentru comenzi. Eroare: {e}"

def gaseste_detalii_tehnologie(reper):
    """
    Reads the Excel file for technologies and returns the details for the given part (reper).
    Returns a dictionary with details if found, or an error message.
    """
    
    global NUME_FISIER_TEHNOLOGII
    NUME_FISIER_TEHNOLOGII = verifica_si_selecteaza_fisier(NUME_FISIER_TEHNOLOGII, "tehnologii") or NUME_FISIER_TEHNOLOGII
    NUME_FISIER_TEHNOLOGII = actualizeaza_cale_fisier(NUME_FISIER_TEHNOLOGII)
    if not os.path.exists(NUME_FISIER_TEHNOLOGII):
        return None, f"Fișierul '{NUME_FISIER_TEHNOLOGII}' nu a fost găsit sau selectat."

    try:
        # Citim foaia 'Sheet1' (implicit)
        df = pd.read_excel(NUME_FISIER_TEHNOLOGII, sheet_name=0) 
        
        # 1. Curățarea spațiilor albe din antetele coloanelor
        df.columns = df.columns.str.strip() 
        
        # 2. Curățarea spațiilor albe din coloana cheie și asigurarea tipului string
        if 'Reper' in df.columns:
            df['Reper'] = df['Reper'].fillna('').astype(str).str.strip()
        else:
            return None, "Coloana 'Reper' nu a fost găsită în foaia de Tehnologii."

        # Cautam reperul
        rezultat = df[df['Reper'] == str(reper)]

        if not rezultat.empty:
            return rezultat.iloc[0].to_dict(), None
        else:
            # Try to show a GUI error if possible, otherwise just return the error message
            try:
                from PyQt6.QtWidgets import QApplication, QMessageBox
                
                # Ensure we have a QApplication instance
                app = QApplication.instance()
                if app is None:
                    app = QApplication([])
                
                QMessageBox.critical(
                    None,
                    "Eroare Tehnologii", 
                    f"Reperul '{reper}' nu a fost găsit în fișierul de Tehnologii.xlsx. Nu există operații pentru acest reper."
                )
            except Exception:
                pass
            return None, f"Reperul '{reper}' nu a fost găsit în fișierul de Tehnologii."

    except Exception as e:
        return None, f"Eroare la citirea/procesarea fișierului Excel pentru tehnologii. Eroare: {e}"

def get_or_create_document_folder(detalii_comanda):
    """
    Builds the path for the output folder based on order details and creates it if it doesn't exist.
    Format: [Reper]_[Comanda Interna]_[Fisa Interna Elmet]
    Returns the folder path or an error message.
    """
    
    def sanitize_name(name):
        """Curăță string-ul de caractere problematice pentru a fi folosit ca nume de director."""
        name = str(name).strip()
        # Înlocuiește separatorul de cale al sistemului de operare cu cratimă
        name = name.replace(os.path.sep, '-')
        # Înlocuiește alte caractere problematice (comune pe Windows/Linux) cu cratimă
        for char in ['*', '?', '"', '<', '>', '|', ':', '\\', '/']:
            name = name.replace(char, '-')
        # Înlocuiește spațiile cu underscore-uri
        name = name.replace(' ', '_')
        return name

    comanda_interna = detalii_comanda.get('Comanda Interna', 'NECUNOSCUT')
    reper = detalii_comanda.get('Reper', 'NECUNOSCUT')
    fisa_interna_elmet = detalii_comanda.get('Fisa Interna Elmet', 'NECUNOSCUT')
    
    # Aplicăm curățarea
    reper_safe = sanitize_name(reper)
    comanda_safe = sanitize_name(comanda_interna)
    fisa_safe = sanitize_name(fisa_interna_elmet)
    
    # Construim numele folderului final
    folder_name = f"{reper_safe}_{comanda_safe}_{fisa_safe}"
    
    # Construim calea absolută completă
    folder_path = os.path.join(WORK_DIR, folder_name)
    
    try:
        # Verify/create parent directory first
        if not os.path.exists(WORK_DIR):
            ensure_dir(WORK_DIR)
        
        # Now create the specific folder
        if not os.path.exists(folder_path):
            ensure_dir(folder_path)
            logging.info(f"Folder nou creat: '{folder_name}' pentru documente.")
        else:
            logging.info(f"Folder existent găsit: '{folder_name}'.")
            
            # Verify write permissions on Windows
            if sys.platform.startswith('win'):
                try:
                    test_file = os.path.join(folder_path, '.test_write')
                    with open(test_file, 'w') as f:
                        f.write('')
                    os.remove(test_file)
                except Exception:
                    # If we can't write, try to fix permissions
                    import subprocess
                    try:
                        subprocess.run(['icacls', folder_path, '/grant', f'{os.getenv("USERNAME")}:(OI)(CI)F'],
                                     capture_output=True, text=True, check=True)
                    except Exception as e:
                        return None, f"Nu se poate scrie în folderul '{folder_name}'. Verificați permisiunile: {e}"
            
        return folder_path, None
    except Exception as e:
        return None, f"Eroare la crearea/accesarea folderului '{folder_name}': {e}"


def run_order(comanda_interna, tip='RC', skip_prompts=False, date_suplimentare=None):
    """Process a single order (RC or COC). Returns a dict log entry."""
    log_file = Path(os.getcwd()) / 'rc_coc_runs.jsonl'
    started = datetime.now().isoformat()
    comanda_interna = str(comanda_interna).strip().upper()
    detalii, eroare = gaseste_detalii_comanda(comanda_interna)
    if eroare:
        entry = {'order': comanda_interna, 'status': 'ERROR', 'message': eroare, 'ts_start': started, 'ts_end': datetime.now().isoformat()}
        try:
            with log_file.open('a', encoding='utf-8') as fh:
                fh.write(json.dumps(entry, ensure_ascii=False) + '\n')
        except Exception:
            pass
        return entry

    folder_path, eroare_folder = get_or_create_document_folder(detalii)
    if eroare_folder:
        entry = {'order': comanda_interna, 'status': 'ERROR', 'message': eroare_folder, 'ts_start': started, 'ts_end': datetime.now().isoformat()}
        try:
            with log_file.open('a', encoding='utf-8') as fh:
                fh.write(json.dumps(entry, ensure_ascii=False) + '\n')
        except Exception:
            pass
        return entry

    if tip == 'RC':
        mesaj, succes = genereaza_route_card_excel(detalii, folder_path)
        status = 'OK' if succes else 'ERROR'
        entry = {'order': comanda_interna, 'status': status, 'message': mesaj, 'ts_start': started, 'ts_end': datetime.now().isoformat()}
        try:
            with log_file.open('a', encoding='utf-8') as fh:
                fh.write(json.dumps(entry, ensure_ascii=False) + '\n')
        except Exception:
            pass
        return entry

    # COC
    def build_coc_defaults(order_str):
        m = re.search(r"(\d+)$", order_str)
        if m:
            digits = m.group(1)
            dcir_digits = digits.zfill(6)
            lot_numeric = str(int(digits))
        else:
            dcir_digits = '000000'
            lot_numeric = 'N/A'
        return {
            'Nr. Certificat': f"DCIR{dcir_digits}",
            'Lot Nr.': lot_numeric,
            'Lot Material client': '',
            'Revizie Desen': detalii.get('Revizie', 'N/A') if detalii and isinstance(detalii.get('Revizie', None), str) else 'N/A',
            'Nume Client': 'Elmet International SRL'
        }

    # If caller provided date_suplimentare explicitly, use it.
    if date_suplimentare is None:
        if skip_prompts:
            date_suplimentare = build_coc_defaults(comanda_interna)
        else:
            date_suplimentare = cere_date_suplimentare_coc(comanda_interna)

    # Caută revizia în fișierul de tehnologii folosind "Cod Reper"
    cod_reper_lookup = detalii.get('Cod Reper', detalii.get('Reper', 'N/A') if detalii else 'N/A') if detalii else 'N/A'
    detalii_tehnologie, eroare_tehnologie = gaseste_detalii_tehnologie(cod_reper_lookup)
    revizie = detalii_tehnologie.get('Revizie', 'N/A') if detalii_tehnologie else 'N/A'
    
    # Extrage datele suplimentare
    nr_certificat = date_suplimentare['Nr. Certificat']
    revizie_desen = date_suplimentare['Revizie Desen']

    mesaj, succes = genereaza_declaratie_conformitate_excel(detalii, date_suplimentare, folder_path)
    status = 'OK' if succes else 'ERROR'
    entry = {'order': comanda_interna, 'status': status, 'message': mesaj, 'ts_start': started, 'ts_end': datetime.now().isoformat()}
    try:
        with log_file.open('a', encoding='utf-8') as fh:
            fh.write(json.dumps(entry, ensure_ascii=False) + '\n')
    except Exception:
        pass
    return entry


def read_log_entries(n=10):
    """Read last n entries from the JSONL log file and return list of dicts."""
    log_file = Path(os.getcwd()) / 'rc_coc_runs.jsonl'
    if not log_file.exists():
        return []
    try:
        with log_file.open('r', encoding='utf-8') as fh:
            lines = [l.strip() for l in fh if l.strip()]
        last = lines[-n:]
        entries = []
        for ln in reversed(last):
            try:
                entries.append(json.loads(ln))
            except Exception:
                entries.append({'raw': ln})
        return entries
    except Exception:
        return []


def load_clients(clients_path=None):
    """Load persistent clients list from clients.json. Returns a list of client names.
    If the file doesn't exist, returns a default list with 'Elmet International SRL'.
    """
    if clients_path is None:
        clients_path = os.path.join(CONFIG_DIR, 'clients.json')
    try:
        if clients_path and os.path.exists(clients_path):
            with open(clients_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
    except Exception:
        pass
    # Fallback default
    return ["Elmet International SRL"]


def save_clients(clients, clients_path=None):
    """Save the clients list to clients.json (overwrites).
    """
    if clients_path is None:
        clients_path = os.path.join(CONFIG_DIR, 'clients.json')
    try:
        os.makedirs(os.path.dirname(clients_path), exist_ok=True)
        with open(clients_path, 'w', encoding='utf-8') as f:
            json.dump(list(dict.fromkeys(clients)), f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def build_coc_defaults(order_str, detalii=None):
    """Return default COC fields extracted from order string and optionally detalii."""
    m = re.search(r"(\d+)$", str(order_str))
    if m:
        digits = m.group(1)
        dcir_digits = digits.zfill(6)
        lot_numeric = str(int(digits))
    else:
        dcir_digits = '000000'
        lot_numeric = 'N/A'
    rev = ''
    if detalii:
        rev = detalii.get('Revizie', '') if isinstance(detalii.get('Revizie', None), str) else ''
    return {
        'Nr. Certificat': f"DCIR{dcir_digits}",
        'Lot Nr.': lot_numeric,
        'Lot Material client': '',
        'Revizie Desen': rev or 'N/A',
        'Nume Client': 'Elmet International SRL'
    }


def cere_date_suplimentare_coc(comanda_interna):
    """
    Asks the user for extra details needed for the Declaration of Conformity (COC).
    Returns a dictionary with the required fields.
    """
    # Use build_coc_defaults to prefill fields and only ask for Lot Material client.
    defaults = build_coc_defaults(comanda_interna)
    print("\n--- Date Suplimentare pentru Declarația de Conformitate (COC) ---")
    print(f"Se folosesc valorile implicite pentru majoritatea câmpurilor: Nr. Certificat={defaults['Nr. Certificat']}, Lot Nr.={defaults['Lot Nr.']}, Nume Client={defaults['Nume Client']}")
    lot_material_client = input("Introduceți Lot Material client (ex: MAT123) (lăsați gol pentru valoarea implicită goală): ").strip()
    if not lot_material_client:
        lot_material_client = ''
        print("Lot Material client lăsat gol. Se va păstra necompletat.")
    # Allow optional override for Lot Nr.
    lot_nr_override = input(f"Doriți să suprascrieți Lot Nr. implicit ({defaults['Lot Nr.']})? Lăsați gol pentru a folosi implicit: ").strip()
    if lot_nr_override:
        lot_nr = lot_nr_override
    else:
        lot_nr = defaults['Lot Nr.']

    # Revizie desen fallback
    revizie_desen = defaults.get('Revizie Desen', 'N/A')
    nume_client = defaults.get('Nume Client', 'Elmet International SRL')
    print("--------------------------------------------------------------------------\n")
    return {
        'Nr. Certificat': defaults['Nr. Certificat'],
        'Lot Nr.': lot_nr,
        'Lot Material client': lot_material_client,
        'Revizie Desen': revizie_desen,
        'Nume Client': nume_client
    }

#################################################################
# 3. DOCUMENT GENERATION FUNCTIONS
#    (Excel file creation for COC and Route Card)
#################################################################
# =============================================================
# 5. DOCUMENT GENERATION FUNCTIONS
# =============================================================

def genereaza_declaratie_conformitate_excel(detalii_comanda, date_suplimentare, folder_path):
    """Generates the Declaration of Conformity (COC) Excel file.
    Uses details from the order and extra user input.
    Returns a success message and status.
    """
    
    # Extrage datele
    comanda_interna = detalii_comanda.get('Comanda Interna', 'NECUNOSCUT')
    reper = detalii_comanda.get('Reper', 'NECUNUT')
    cantitate = detalii_comanda.get('Cantitate', 'N/A')
    
    # Caută revizia în fișierul de tehnologii folosind "Cod Reper"
    cod_reper_lookup = detalii_comanda.get('Cod Reper', reper)
    detalii_tehnologie, eroare_tehnologie = gaseste_detalii_tehnologie(cod_reper_lookup)
    revizie = detalii_tehnologie.get('Revizie', 'N/A') if detalii_tehnologie else 'N/A'
    
    # Extrage datele suplimentare
    nr_certificat = date_suplimentare['Nr. Certificat']
    revizie_desen = date_suplimentare['Revizie Desen']
    
    cantitate_str = str(int(cantitate)) if pd.notna(cantitate) and isinstance(cantitate, (int, float)) else str(cantitate)
    
    # Construiește calea completă de salvare
    nume_fisier_local = f"Declaratie_Conformitate_{nr_certificat}_{comanda_interna}_{reper}.xlsx"
    nume_fisier_output = os.path.join(folder_path, nume_fisier_local)
    
    print(f"DEBUG: Calea completă de salvare COC: {nume_fisier_output}") # DEBUG
    
    try:
        # 1. Inițializare Workbook și Worksheet
        workbook = xlsxwriter.Workbook(nume_fisier_output)
        worksheet = workbook.add_worksheet("COC")

        # ==========================================================================
        # SETĂRI DE IMPRIMARE ȘI LĂȚIME GENERALĂ (Folosim 8 coloane A-H)
        # ==========================================================================
        worksheet.set_paper(9)           # A4
        worksheet.set_portrait()         # Portret
        worksheet.set_margins(0.5, 0.5, 0.5, 0.5) 
        worksheet.fit_to_pages(1, 0)     # Fit la o singură pagină pe lățime
        
        # Setăm lățimile coloanelor pentru o aliniere vizuală bună
        worksheet.set_column(0, 1, 15)  # A:B
        worksheet.set_column(2, 3, 15)  # C:D
        worksheet.set_column(4, 5, 15)  # E:F 
        worksheet.set_column(6, 7, 15)  # G:H

        # 2. Definirea stilurilor (Formatelor)
        fmt_title = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'border': 0}) 
        fmt_company_info = workbook.add_format({'font_size': 9, 'align': 'center', 'text_wrap': True})
        
        # Stiluri pentru casetele de date (cu chenar)
        fmt_box_label = workbook.add_format({'bold': True, 'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'bg_color': '#D9D9D9'})
        fmt_box_data = workbook.add_format({'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFFE0'}) 
        
        fmt_header_conformity = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'bottom': 2, 'top': 2}) 
        
        fmt_table_header = workbook.add_format({'bold': True, 'font_size': 9, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'bg_color': '#E0E0E0'})
        fmt_table_data = workbook.add_format({'font_size': 10, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        fmt_signature_label = workbook.add_format({'bold': True, 'font_size': 10, 'align': 'left', 'top': 1})
        
        # 3. Secțiunea I: Antet și Date Companie (INRED GROUP SRL)
        worksheet.merge_range(0, 0, 0, 7, 'S.C. INRED GROUP SRL', fmt_title)
        worksheet.merge_range(1, 0, 1, 7, 'Str.Sat Racauti 599, Comuna Buciumi, Bacau, Romania', fmt_company_info)
        worksheet.merge_range(2, 0, 2, 7, 'Cod Unic Inregistrare: RO24705289', fmt_company_info)
        worksheet.merge_range(3, 0, 3, 7, 'Nr.de ordine in Registrul Comertului: J04/1960/2008', fmt_company_info)

        # 4. Secțiunea II: Elemente de Identificare (Rândul 6-7)
        worksheet.set_row(5, 40)
        
        # Only two blocks: Nr. Certificat and Client, each taking half the width
        worksheet.merge_range(5, 0, 5, 3, 'Nr. Certificat / Coc No.', fmt_box_label)  # A6:D6
        worksheet.merge_range(6, 0, 6, 3, date_suplimentare['Nr. Certificat'], fmt_box_data)  # A7:D7

        worksheet.merge_range(5, 4, 5, 7, 'Client / Customer', fmt_box_label)  # E6:H6
        worksheet.merge_range(6, 4, 6, 7, date_suplimentare['Nume Client'], fmt_box_data)  # E7:H7

        # Rândul 9: Titlu Declarație Conformitate
        worksheet.merge_range(8, 0, 8, 7, 'DECLARAȚIE CONFORMITATE / Declaration of Conformity', fmt_header_conformity)  # A9:H9

        # 5. Secțiunea III: Detalii Comandă (Rândul 11-12)
        
        # Cap de tabel (Rândul 11)
        worksheet.merge_range(10, 0, 10, 1, 'Comanda Interna / Internal order', fmt_table_header)  # A11:B11
        worksheet.merge_range(10, 2, 10, 3, 'Nr. Buc. / No. Pcs', fmt_table_header)  # C11:D11
        worksheet.write(10, 4, 'Lot Nr. / Batch No.', fmt_table_header)  # E11
        worksheet.write(10, 5, 'Comanda client / Client External Order', fmt_table_header)  # F11
        worksheet.merge_range(10, 6, 10, 7, 'Comanda Interna client / Client internal order', fmt_table_header)  # G11:H11

        # Datele (Rândul 12)
        worksheet.merge_range(11, 0, 11, 1, comanda_interna, fmt_table_data)  # A12:B12
        worksheet.merge_range(11, 2, 11, 3, cantitate_str, fmt_table_data)  # C12:D12
        worksheet.write(11, 4, date_suplimentare['Lot Nr.'], fmt_table_data)  # E12
        worksheet.write(11, 5, detalii_comanda.get('Comanda', 'N/A'), fmt_table_data)  # F12
        worksheet.merge_range(11, 6, 11, 7, detalii_comanda.get('Fisa Interna Elmet', 'N/A'), fmt_table_data)  # G12:H12

        # 6. Secțiunea IV: Identificare Produs (Rândul 14-15)
        
        row_prod_id = 14
        
        # Denumire Produs
        worksheet.merge_range(row_prod_id, 0, row_prod_id, 2, 'Denumire Produs / Part Description', fmt_box_label)
        worksheet.merge_range(row_prod_id, 3, row_prod_id, 7, detalii_comanda.get('Denumire', 'N/A'), fmt_box_data)
        
        # Cod Reper / Revizie
        worksheet.merge_range(row_prod_id + 1, 0, row_prod_id + 1, 2, 'Cod Reper / Drawing No.', fmt_box_label)
        worksheet.merge_range(row_prod_id + 1, 3, row_prod_id + 1, 5, detalii_comanda.get('Reper', 'N/A'), fmt_box_data)
        worksheet.write(row_prod_id + 1, 6, 'Rev.', fmt_box_label)
        worksheet.write(row_prod_id + 1, 7, revizie, fmt_box_data)
        
        # Lot Material Client (Rândul 17) - now uses user input
        worksheet.merge_range(row_prod_id + 3, 0, row_prod_id + 3, 2, 'Lot Material client / Client Material Batch No.', fmt_box_label)
        worksheet.merge_range(row_prod_id + 3, 3, row_prod_id + 3, 7, date_suplimentare.get('Lot Material client', 'XXXX'), fmt_box_data)

        # 7. Secțiunea V: Declarații de Conformitate și Semnături
        
        row_conformity = row_prod_id + 5
        
        # Conformitate 
        worksheet.merge_range(row_conformity, 0, row_conformity, 7, 
                              'Este conform specificațiilor din: / Conforms With Specifications Of:', 
                              workbook.add_format({'bold': True, 'font_size': 10, 'align': 'left'}))

        # Căsuța CERINTE CLIENT/ CUSTOMER REQ 
        worksheet.merge_range(row_conformity + 1, 0, row_conformity + 1, 7, 
                              'CERINȚE CLIENT / CUSTOMER REQ', 
                              workbook.add_format({'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 2, 'bg_color': '#E0E0E0'}))

        # Textul Lung - updated as requested
        long_text = (
            "Prin prezenta, declarăm că piesele aferente certificatului de față sunt în conformitate cu cerințele clientului, transmise prin intermediul comenzii ferme.\n"
            "Toate prelucrările au fost efectuate conform specificațiilor tehnice primite, respectând standardele de calitate aplicabile și cerințele contractuale."
        )
        fmt_long_text = workbook.add_format({'font_size': 9, 'align': 'justify', 'valign': 'top', 'text_wrap': True})
        worksheet.merge_range(row_conformity + 3, 0, row_conformity + 5, 7, long_text, fmt_long_text)
        worksheet.set_row(row_conformity + 3, 40)

        # Secțiunea de Semnătură
        row_signature = row_conformity + 9
        
        # Data Emitere Certificat
        worksheet.write(row_signature, 0, 'Data emitere certificat / Issued date:', fmt_signature_label)
        worksheet.merge_range(row_signature, 1, row_signature, 3, datetime.now().strftime("%d.%m.%Y"), fmt_table_data)
        
        # Semnătură - updated to Adm. Slevoaca Bogdan
        worksheet.merge_range(row_signature, 4, row_signature, 7, 'Adm. Slevoaca Bogdan', fmt_signature_label)

        # Declarație Materiale 
        worksheet.merge_range(row_signature + 2, 0, row_signature + 2, 7, 
                              'Declara că toate materialele folosite nu sunt de origine contrafăcută / Declares that all materials used are not counterfeit origin.', 
                              workbook.add_format({'font_size': 9, 'align': 'left', 'italic': True, 'top': 1}))


        # Finalizare
        workbook.close()
        return f"Succes! Declarația de Conformitate a fost generată și salvată în: {nume_fisier_output}", True
    
    except Exception as e:
        return f"Eroare critică la generarea fișierului Excel COC: {e}", False


def genereaza_route_card_excel(detalii_comanda, folder_path):
    """
    Generates the Route Card Excel file and fills it with operations from the Technologies Excel file.
    Returns a success message and status.
    """
    
    comanda_interna = detalii_comanda.get('Comanda Interna', 'NECUNOSCUT')
    reper = detalii_comanda.get('Reper', 'NECUNOSCUT')
    pozitie = detalii_comanda.get('Pozitie', 'N/A')
    
    # ----------------------------------------------------------------------
    # 1. Caută detaliile de tehnologie (Fluxul de Operații)
    # ----------------------------------------------------------------------
    detalii_tehnologie, eroare_tehnologie = gaseste_detalii_tehnologie(reper)

    if eroare_tehnologie:
        print(f"ATENTIE: Nu s-a putut incarca tehnologia pentru reperul '{reper}'. Tabelul de operatii va fi gol. Eroare: {eroare_tehnologie}")
        tech_data = {}
        revizie_desen = ''
        material_brut = ''
        operations = []
    else:
        tech_data = detalii_tehnologie or {}
        revizie_desen = tech_data.get('Revizie', '').strip() if tech_data else ''
        # Auto-fill Material brut from Tehnologii.xlsx column D (should be named 'Material brut')
        material_brut = tech_data.get('Material brut', '').strip() if tech_data else ''
        # Extrage operațiile dinamice
        operations = []
        max_operations = 10
        for i in range(1, max_operations + 1):
            op_num_col = f'OP{i*10}'
            op_time_col = f'TOP{i*10}'
            # Numele coloanei de locatie variază: 'Utilaj/Locație' pentru OP10, 'Utilaj/Locație2' pentru OP20, etc.
            op_loc_col = f'Utilaj/Locație{"" if i == 1 else i}' 
            
            # Verificăm dacă există o denumire de operație validă
            operatie_text = tech_data.get(op_num_col, '') if tech_data else ''
            if operatie_text and str(operatie_text).strip() not in ('', 'nan'):
                operations.append({
                    'nr_op': f'{i*10}',
                    'operatie': str(operatie_text).strip(),
                    'timp': str(tech_data.get(op_time_col, '') if tech_data else '').strip(),
                    'locatie': str(tech_data.get(op_loc_col, '') if tech_data else '').strip(),
                })
            else:
                # Oprim când găsim prima operație goală
                break
    # ----------------------------------------------------------------------

    
    # Construiește calea completă de salvare
    nume_fisier_local = f"Route_Card_{comanda_interna}_{reper}_P{pozitie}.xlsx"
    nume_fisier_output = os.path.join(folder_path, nume_fisier_local)
    
    print(f"DEBUG: Calea completă de salvare Route Card: {nume_fisier_output}") # DEBUG
    
    try:
        # 1. Inițializare Workbook și Worksheet
        workbook = xlsxwriter.Workbook(nume_fisier_output)
        worksheet = workbook.add_worksheet("Route Card")

        # ==========================================================================
        # SETĂRI DE IMPRIMARE ȘI LĂȚIME GENERALĂ PE 8 COLOANE (A-H)
        # ==========================================================================
        worksheet.set_paper(9)           # A4
        worksheet.set_portrait()         # Portret
        worksheet.set_margins(0.5, 0.5, 0.5, 0.5) 
        worksheet.repeat_rows(9, 1)      # Repetăm capul de tabel (Rândurile 10 și 11)
        worksheet.fit_to_pages(1, 0)     # Fit la o singură pagină pe lățime
        
        # Setăm lățimile coloanelor pentru Secțiunea I (A-H)
        worksheet.set_column(0, 0, 10)  # A:A - Etichetă (Lățime mică)
        worksheet.set_column(1, 1, 15)  # B:B - Valoare (Lățime medie)  
        worksheet.set_column(2, 2, 10)  # C:C - Etichetă (Lățime mică)
        worksheet.set_column(3, 3, 15)  # D:D - Valoare (Lățime medie)
        worksheet.set_column(4, 4, 10)  # E:E - Etichetă (Lățime mică)
        worksheet.set_column(5, 5, 8)   # F:F
        worksheet.set_column(6, 6, 8)   # G:G
        worksheet.set_column(7, 7, 9)   # H:H
        
        # 2. Definirea stilurilor (Formatelor)
        fmt_header = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E0E0E0', 'border': 2}) 
        fmt_subheader = workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#F0F0F0', 'border': 1, 'align': 'center'})
        fmt_label = workbook.add_format({'bold': True, 'font_size': 9, 'border': 1, 'align': 'left', 'valign': 'vcenter'})
        fmt_data = workbook.add_format({'font_size': 10, 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        fmt_input_highlight = workbook.add_format({'bold': True, 'font_size': 12, 'bg_color': '#FFFFE0', 'border': 2, 'align': 'center', 'valign': 'vcenter'})
        fmt_table_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#BFBFBF', 'border': 1, 'text_wrap': True, 'font_size': 9})
        fmt_table_subheader = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#D9D9D9', 'border': 1, 'font_size': 8})
        fmt_table_data = workbook.add_format({'align': 'center', 'border': 1})
        fmt_table_data_left = workbook.add_format({'align': 'left', 'border': 1, 'text_wrap': True}) # Text wrap pentru operație/locație
        
        # 3. Secțiunea I: Antet și Identificare Comandă
        
        # Antet (Merge pe toate cele 8 coloane A-H)
        worksheet.merge_range(0, 0, 0, 7, 'ROUTE CARD', fmt_header)  # A1:H1
        worksheet.merge_range(1, 0, 1, 7, 'S.C. INRed Group S.R.L.', fmt_subheader)  # A2:H2
        
        # Rândul 4: Comanda Interna (Input cheie)
        worksheet.write(3, 0, 'COMANDA INTERNĂ:', fmt_label)  # A4
        worksheet.write(3, 1, comanda_interna, fmt_input_highlight)  # B4
        worksheet.write(3, 2, 'Data Generare:', fmt_label)  # C4
        worksheet.write(3, 3, datetime.now().strftime("%Y-%m-%d"), fmt_data)  # D4
        worksheet.write(3, 4, 'Poziție Comandă:', fmt_label)  # E4
        worksheet.merge_range(3, 5, 3, 7, detalii_comanda.get('Pozitie', 'N/A'), fmt_data)  # F4:H4

        # Rândul 5: Informații client și intern
        worksheet.write(4, 0, 'Comandă Client:', fmt_label)  # A5
        worksheet.write(4, 1, detalii_comanda.get('Comanda', 'N/A'), fmt_data)  # B5
        worksheet.write(4, 2, 'Fisa Interna Elmet:', fmt_label)  # C5
        worksheet.write(4, 3, detalii_comanda.get('Fisa Interna Elmet', 'N/A'), fmt_data)  # D5
        worksheet.write(4, 4, 'Data Comanda:', fmt_label)  # E5
        data_comanda = detalii_comanda.get('Data Comanda', 'N/A')
        if isinstance(data_comanda, datetime):
            data_comanda = data_comanda.strftime("%Y-%m-%d")
        worksheet.merge_range(4, 5, 4, 7, data_comanda, fmt_data)  # F5:H5

        # Rândul 6: Reper și Cantitate
        worksheet.write(5, 0, 'Cod Piesă (Reper):', fmt_label)  # A6
        worksheet.write(5, 1, detalii_comanda.get('Reper', 'N/A'), fmt_data)  # B6
        worksheet.write(5, 2, 'Denumire Piesă:', fmt_label)  # C6
        worksheet.write(5, 3, detalii_comanda.get('Denumire', 'N/A'), fmt_data)  # D6
        worksheet.write(5, 4, 'Cantitate Comandată:', fmt_label)  # E6
        cantitate = detalii_comanda.get('Cantitate', 'N/A')
        cantitate_str = str(int(cantitate)) if pd.notna(cantitate) and isinstance(cantitate, (int, float)) else str(cantitate)
        worksheet.merge_range(5, 5, 5, 7, cantitate_str, fmt_data)  # F6:H6
        
        # Rândul 7: Materiale și Revizii (populat acum)
        worksheet.write(6, 0, 'Material Brut:', fmt_label)  # A7
        worksheet.write(6, 1, material_brut, fmt_data)  # B7 - Material Brut
        worksheet.write(6, 2, 'Revizie Desen:', fmt_label)  # C7
        worksheet.write(6, 3, revizie_desen, fmt_data)  # D7 - Populat din Tehnologii.xlsx
        worksheet.write(6, 4, 'Status Material:', fmt_label)  # E7
        worksheet.merge_range(6, 5, 6, 7, detalii_comanda.get('Status Material', 'N/A'), fmt_data)  # F7:H7

        # 4. Secțiunea II: Fluxul Operațiilor (Tabel de lucru)
        
        # Antet Secțiune (Merge pe 8 coloane A-H)
        worksheet.merge_range(8, 0, 8, 7, 'II. FLUXUL OPERAȚIILOR ȘI ÎNREGISTRĂRI (Populat din Tehnologii.xlsx)', fmt_subheader)  # A9:H9
        
        # Re-setăm lățimile coloanelor pentru tabelul de operații (A-H)
        worksheet.set_column(0, 0, 6)   # A:A - Nr. Op.
        worksheet.set_column(1, 1, 15)  # B:B - Operație
        worksheet.set_column(2, 2, 15)  # C:C - Utilaj / Locație
        worksheet.set_column(3, 3, 10)  # D:D - Timp Standard
        worksheet.set_column(4, 4, 7)   # E:E - Cantitate OK
        worksheet.set_column(5, 5, 7)   # F:F - Cantitate REJ
        worksheet.set_column(6, 6, 12)  # G:G - Data / Ora
        worksheet.set_column(7, 7, 18)  # H:H - Operator / Semnătură
        
        # Antet Tabel - Rândul 10 și 11 (Indicii 9 și 10)
        
        # Coloane care se întind pe 2 rânduri
        worksheet.merge_range(9, 0, 10, 0, 'Nr. Op.', fmt_table_header)      
        worksheet.merge_range(9, 1, 10, 1, 'Operație', fmt_table_header)     
        worksheet.merge_range(9, 2, 10, 2, 'Utilaj / Locație', fmt_table_header) 
        worksheet.merge_range(9, 3, 10, 3, 'Timp Standard (min)', fmt_table_header) 
        
        # Cantitate realizată - Eticheta mare 
        worksheet.merge_range(9, 4, 9, 5, 'Cantitate REALIZATĂ', fmt_table_header) 
        
        # Sub-etichetele (OK, REJ) pe rândul 11 
        worksheet.write(10, 4, 'OK', fmt_table_subheader)   
        worksheet.write(10, 5, 'REJ', fmt_table_subheader)  
        
        # Coloane care se întind pe 2 rânduri
        worksheet.merge_range(9, 6, 10, 6, 'Data / Ora', fmt_table_header)      
        worksheet.merge_range(9, 7, 10, 7, 'Operator / Semnătură', fmt_table_header) 

        # Linii pentru înregistrare (dinamic + rânduri libere) - MAX 12 RANDURI
        start_row = 11
        max_rows = 12
        num_dynamic_rows = len(operations)

        # Scriere operații dinamice
        for idx, op in enumerate(operations):
            row = start_row + idx
            worksheet.set_row(row, 20) # Setăm o înălțime de rând mai mare pentru vizibilitate
            worksheet.write(row, 0, op['nr_op'], fmt_table_data)          # Nr. Op.
            worksheet.write(row, 1, op['operatie'], fmt_table_data_left)  # Operație
            worksheet.write(row, 2, op['locatie'], fmt_table_data_left)   # Utilaj / Locație
            worksheet.write(row, 3, op['timp'], fmt_table_data)           # Timp Standard
            worksheet.write(row, 4, '', fmt_table_data)                  # Cantitate OK (Manual)
            worksheet.write(row, 5, '', fmt_table_data)                  # Cantitate REJ (Manual)
            worksheet.write(row, 6, '', fmt_table_data)                  # Data / Ora (Manual)
            worksheet.write(row, 7, '', fmt_table_data)                  # Operator / Semnătură (Manual)

        # Scriere rânduri libere rămase (până la max_rows)
        for row_idx in range(num_dynamic_rows, max_rows):
            row = start_row + row_idx
            worksheet.set_row(row, 20) # Setăm aceeași înălțime
            worksheet.write(row, 0, '', fmt_table_data)          
            worksheet.write(row, 1, '', fmt_table_data_left)     
            worksheet.write(row, 2, '', fmt_table_data_left)     
            worksheet.write(row, 3, '', fmt_table_data)          
            worksheet.write(row, 4, '', fmt_table_data)          
            worksheet.write(row, 5, '', fmt_table_data)          
            worksheet.write(row, 6, '', fmt_table_data)          
            worksheet.write(row, 7, '', fmt_table_data)          
            
        # 5. Secțiunea III: Finalizare
        
        row_final = 24 # Începe după cele 12 rânduri de operații (de la rândul 11 la 22) + 1 rând liber (23)
        if (start_row + max_rows) > 23:
            row_final = start_row + max_rows + 1

        # Titlu Secțiune 
        worksheet.merge_range(row_final, 0, row_final, 7, 'III. CONTROL FINAL ȘI ÎNCHEIERE', fmt_subheader) 

        # Rândul 26 
        worksheet.write(row_final + 2, 0, 'Total Cantitate OK Finală:', fmt_label)
        worksheet.merge_range(row_final + 2, 1, row_final + 2, 3, '', fmt_data) 
        worksheet.write(row_final + 2, 4, 'Data Control Final:', fmt_label)
        worksheet.merge_range(row_final + 2, 5, row_final + 2, 7, '', fmt_data) 
        
        # Rândul 27 
        worksheet.write(row_final + 3, 0, 'Inspector QC Final:', fmt_label)
        worksheet.merge_range(row_final + 3, 1, row_final + 3, 3, '', fmt_data) 
        worksheet.write(row_final + 3, 4, 'Semnătură QC:', fmt_label)
        worksheet.merge_range(row_final + 3, 5, row_final + 3, 7, '', fmt_data) 
        
        # Rândul 29 
        worksheet.write(row_final + 5, 0, 'Manager Producție:', fmt_label)
        worksheet.merge_range(row_final + 5, 1, row_final + 5, 7, '', fmt_data) 
        
        # Finalizare
        workbook.close()
        return f"Succes! Route Card-ul a fost generat și salvat în: {nume_fisier_output}", True
    
    except Exception as e:
        return f"Eroare critică la generarea fișierului Excel: {e}", False



#################################################################
# 4. GUI (Tkinter) & MAIN APPLICATION LOGIC
#################################################################
# =============================================================
# 6. GUI FUNCTIONS
# =============================================================

def ruleaza_aplicatia_pyqt():
    """
    Rulează aplicația folosind PyQt6 pentru o interfață grafică modernă.
    """
    try:
        from PyQt6.QtWidgets import (
            QApplication, QWidget, QVBoxLayout, QHBoxLayout,
            QPushButton, QLabel, QRadioButton, QListWidget,
            QMessageBox, QInputDialog, QListWidgetItem, QLineEdit,
            QStatusBar
        )
        from PyQt6.QtGui import QPalette, QColor, QFont
        from PyQt6.QtCore import Qt, QSize
    except ImportError:
        # Fallback la Tkinter dacă PyQt6 nu este instalat
        return ruleaza_aplicatia_gui()

    # Paletă de culori modernă (dark mode)
    dark_palette = QPalette()
    dark_palette.setColor(QPalette.ColorRole.Window, QColor(30, 30, 30))
    dark_palette.setColor(QPalette.ColorRole.WindowText, QColor(255, 255, 255))  # white
    dark_palette.setColor(QPalette.ColorRole.Base, QColor(45, 45, 45))
    dark_palette.setColor(QPalette.ColorRole.AlternateBase, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(255, 255, 255))  # white
    dark_palette.setColor(QPalette.ColorRole.ToolTipText, QColor(255, 255, 255))  # white
    dark_palette.setColor(QPalette.ColorRole.Text, QColor(255, 255, 255))  # white
    dark_palette.setColor(QPalette.ColorRole.Button, QColor(53, 53, 53))
    dark_palette.setColor(QPalette.ColorRole.ButtonText, QColor(255, 255, 255))  # white
    dark_palette.setColor(QPalette.ColorRole.BrightText, QColor(255, 0, 0))  # red
    dark_palette.setColor(QPalette.ColorRole.Link, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
    dark_palette.setColor(QPalette.ColorRole.HighlightedText, QColor(0, 0, 0))  # black

    app = QApplication(sys.argv)
    app.setPalette(dark_palette)
    app.setFont(QFont("Segoe UI", 10))

    window = QWidget()
    window.setWindowTitle("RC & COC Generator")
    window.setGeometry(100, 100, 800, 600)

    # Initialize global file paths
    global NUME_FISIER_COMENZI, NUME_FISIER_TEHNOLOGII, WORK_DIR
    NUME_FISIER_COMENZI = get_saved_file_path('comenzi_file', "Planificare Elmet.xlsx")
    NUME_FISIER_TEHNOLOGII = get_saved_file_path('tehnologii_file', "Tehnologii.xlsx")
    
    # Initialize output directory from saved settings
    WORK_DIR = get_saved_file_path('output_dir', WORK_DIR)
    ensure_dir(WORK_DIR)

    main_layout = QVBoxLayout(window)
    
    # Create tab widget
    tabs = QTabWidget()
    main_layout.addWidget(tabs)
    
    # Main tab
    main_tab = QWidget()
    main_tab_layout = QVBoxLayout(main_tab)
    
    # Settings tab
    settings_tab = QWidget()
    settings_layout = QVBoxLayout(settings_tab)
    
    # Add file path settings
    settings_group = QGroupBox("Setări Fișiere")
    settings_group_layout = QGridLayout()
    
    # Orders file setting
    comenzi_label = QLabel("Fișier Comenzi:")
    comenzi_path = QLineEdit()
    comenzi_path.setText(NUME_FISIER_COMENZI)
    comenzi_path.setReadOnly(True)
    comenzi_btn = QPushButton("Modifică")
    
    def update_comenzi_path():
        global NUME_FISIER_COMENZI
        file_path = QFileDialog.getOpenFileName(
            window,
            "Selectați fișierul cu comenzi",
            os.path.dirname(NUME_FISIER_COMENZI) if NUME_FISIER_COMENZI else "",
            "Excel Files (*.xlsx *.xls)"
        )[0]
        if file_path:
            NUME_FISIER_COMENZI = file_path
            comenzi_path.setText(file_path)
            save_file_path('comenzi_file', file_path)
            load_orders()  # Reload orders list with new file
    
    comenzi_btn.clicked.connect(update_comenzi_path)
    
    # Technologies file setting
    tehn_label = QLabel("Fișier Tehnologii:")
    tehn_path = QLineEdit()
    tehn_path.setText(NUME_FISIER_TEHNOLOGII)
    tehn_path.setReadOnly(True)
    tehn_btn = QPushButton("Modifică")
    
    def update_tehn_path():
        global NUME_FISIER_TEHNOLOGII
        file_path = QFileDialog.getOpenFileName(
            window,
            "Selectați fișierul cu tehnologii",
            os.path.dirname(NUME_FISIER_TEHNOLOGII) if NUME_FISIER_TEHNOLOGII else "",
            "Excel Files (*.xlsx *.xls)"
        )[0]
        if file_path:
            NUME_FISIER_TEHNOLOGII = file_path
            tehn_path.setText(file_path)
            save_file_path('tehnologii_file', file_path)
    
    tehn_btn.clicked.connect(update_tehn_path)
    
    # Add widgets to settings layout
    settings_group_layout.addWidget(comenzi_label, 0, 0)
    settings_group_layout.addWidget(comenzi_path, 0, 1)
    settings_group_layout.addWidget(comenzi_btn, 0, 2)
    settings_group_layout.addWidget(tehn_label, 1, 0)
    settings_group_layout.addWidget(tehn_path, 1, 1)
    settings_group_layout.addWidget(tehn_btn, 1, 2)
    
    # Output folder setting
    output_label = QLabel("Director Salvare:")
    output_path = QLineEdit()
    # Load saved output directory or use default
    saved_output_dir = get_saved_file_path('output_dir', WORK_DIR)
    output_path.setText(saved_output_dir)
    output_path.setReadOnly(False)  # Make it editable
    output_btn = QPushButton("Selectează...")
    output_btn.clicked.connect(lambda: select_output_directory(output_path))
    
    # Add handler to save path when manually edited
    def on_output_path_changed():
        global WORK_DIR
        new_path = output_path.text()
        if os.path.isdir(new_path):
            save_file_path('output_dir', new_path)
            WORK_DIR = new_path
    
    output_path.textChanged.connect(lambda: on_output_path_changed())
    
    output_info = QLabel(f"Fișierele generate vor fi salvate în subdirectoare în această locație")
    output_info.setWordWrap(True)
    
    settings_group_layout.addWidget(output_label, 2, 0)
    settings_group_layout.addWidget(output_path, 2, 1)
    settings_group_layout.addWidget(output_btn, 2, 2)
    settings_group_layout.addWidget(output_info, 3, 0, 1, 3)
    
    settings_group.setLayout(settings_group_layout)
    settings_layout.addWidget(settings_group)
    settings_layout.addStretch()
    
    # Add the main controls to the main tab
    # Radio buttons for document type
    radio_layout = QHBoxLayout()
    radio_rc = QRadioButton("Route Card (RC)")
    radio_coc = QRadioButton("Declarație Conformitate (COC)")
    radio_rc.setChecked(True)
    radio_layout.addWidget(radio_rc)
    radio_layout.addWidget(radio_coc)
    main_tab_layout.addLayout(radio_layout)

    # Search bar
    search_input = QLineEdit()
    search_input.setPlaceholderText("Caută comenzi...")
    main_tab_layout.addWidget(search_input)

    # Orders list
    orders_list = QListWidget()
    orders_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
    main_tab_layout.addWidget(orders_list)
    
    def load_orders():
        """Load orders from the Excel file and populate the list."""
        orders_list.clear()
        try:
            df = pd.read_excel(NUME_FISIER_COMENZI, sheet_name='Comenzi', skiprows=[0])
            df.columns = df.columns.str.strip()
            comenzi = []
            if 'Comanda Interna' in df.columns:
                comenzi = df['Comanda Interna'].dropna().astype(str).str.strip().unique()
                for comanda in comenzi:
                    if comanda:
                        orders_list.addItem(QListWidgetItem(comanda))
            status_bar.setText(f"S-au încărcat {len(comenzi)} comenzi.")
        except Exception as e:
            QMessageBox.critical(window, "Eroare la încărcare", f"Nu s-a putut citi fișierul de comenzi:\n{e}")

    # Add the buttons layouts to main tab
    buttons_layout = QHBoxLayout()
    btn_generate = QPushButton("Generează")
    btn_email = QPushButton("Trimite Email")
    btn_email.setEnabled(False)  # Disabled initially
    buttons_layout.addWidget(btn_generate)
    buttons_layout.addWidget(btn_email)
    main_tab_layout.addLayout(buttons_layout)
    
    # Add refresh button in its own layout
    refresh_layout = QHBoxLayout()
    refresh_btn = QPushButton("Reîmprospătare Comenzi")
    refresh_btn.clicked.connect(load_orders)  # Connect to load_orders function
    refresh_layout.addWidget(refresh_btn)
    main_tab_layout.addLayout(refresh_layout)
    
    # Add status bar to main tab
    status_bar = QLabel("Status: Gata pentru generare documente")
    main_tab_layout.addWidget(status_bar)
    
    # Add tabs to main window
    tabs.addTab(main_tab, "Generator")
    tabs.addTab(settings_tab, "Setări")

    # Add status bar
    status_bar = QLabel("Status: Gata pentru generare documente")
    main_layout.addWidget(status_bar)

    generated_files_for_email = []

    def send_email_with_attachments():
        if not generated_files_for_email:
            QMessageBox.warning(window, "Eroare", "Nu există fișiere generate pentru a fi trimise.")
            return

        orders = sorted(list(set([f[0] for f in generated_files_for_email])))
        subject = f"Documente pentru comenzile: {', '.join(orders)}"
        body = f"Buna ziua,\n\nAtasate regasiti documentele pentru comenzile: {', '.join(orders)}.\n\nO zi buna!"
        
        file_paths = [os.path.abspath(f[1]) for f in generated_files_for_email]

        system = platform.system()

        if system == 'Darwin':  # macOS
            try:
                # AppleScript expects a list of POSIX paths in the format: {"/path/to/file1", "/path/to/file2"}
                posix_paths = '{' + ', '.join([f'"{p}"' for p in file_paths]) + '}'
                
                script = f'''
                tell application "Mail"
                    set newMessage to make new outgoing message with properties {{subject:"{subject}", content:"{body}" & return & return, visible:true}}
                    tell newMessage
                        make new to recipient at end of to recipients with properties {{address:""}}
                        repeat with aFile in {posix_paths}
                            make new attachment with properties {{file name:aFile as POSIX file}} at after the last paragraph
                        end repeat
                    end tell
                    activate
                end tell
                '''
                # Using subprocess.run to capture output for better debugging
                result = subprocess.run(['osascript', '-e', script], check=True, capture_output=True, text=True)
                if result.returncode == 0:
                    QMessageBox.information(window, "Succes", "E-mailul a fost creat în aplicația Mail.")
                else:
                    # If osascript returns an error, show it
                    QMessageBox.critical(window, "Eroare AppleScript", f"Nu s-a putut crea e-mailul:\n{result.stderr}")
            except subprocess.CalledProcessError as e:
                # This catches errors if the script itself fails to run
                QMessageBox.critical(window, "Eroare la Execuție", f"Eroare la rularea scriptului de e-mail:\n{e.stderr}")
            except Exception as e:
                # General fallback for other errors
                QMessageBox.critical(window, "Eroare Mail", f"Nu s-a putut crea e-mailul: {e}")

        elif system == 'Windows':
            try:
                import win32com.client as win32  # type: ignore
                outlook = win32.Dispatch('outlook.application')
                mail = outlook.CreateItem(0)
                mail.Subject = subject
                mail.Body = body
                for file_path in file_paths:
                    mail.Attachments.Add(file_path)
                mail.Display(True)
            except ImportError:
                QMessageBox.warning(window, "Funcționalitate Limitată",
                                    "Pentru a atașa fișiere automat pe Windows, biblioteca 'pywin32' este necesară. "
                                    "Se va deschide clientul de e-mail fără atașamente.")
                webbrowser.open(f"mailto:?subject={subject}&body={body}")
            except Exception as e:
                QMessageBox.critical(window, "Eroare Outlook", f"Nu s-a putut crea e-mailul: {e}")
        else: # Fallback for other OS (like Linux)
            webbrowser.open(f"mailto:?subject={subject}&body={body}")
            QMessageBox.information(window, "Acțiune Manuală Necesară",
                                    "E-mailul a fost deschis. Vă rugăm să atașați manual fișierele.")

    def filter_orders():
        """Filter the orders list based on the search input."""
        filter_text = search_input.text().lower()
        for i in range(orders_list.count()):
            item = orders_list.item(i)
            if item:  # Check if item exists
                item.setHidden(filter_text not in item.text().lower())

    def start_batch_generation():
        """Handle batch generation of selected orders"""
        selected_items = orders_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(window, "Eroare", "Selectați cel puțin o comandă internă pentru generare batch.")
            return

        selected_orders = [item.text() for item in selected_items]
        tip_document = 'RC' if radio_rc.isChecked() else 'COC'
        
        errors = []
        successes = []
        generated_files_for_email.clear() # Clear previous files
        
        lot_material_client = None
        nume_client = "Elmet International SRL"

        if tip_document == 'COC':
            lot_material_client, ok = QInputDialog.getText(window, 'Lot Material client', 'Introduceți Lot Material client pentru toate comenzile:')
            if not ok:
                return # User cancelled

        for comanda_interna in selected_orders:
            detalii, eroare_cautare = gaseste_detalii_comanda(comanda_interna)
            if eroare_cautare:
                errors.append(f"{comanda_interna}: {eroare_cautare}")
                continue

            folder_path, eroare_folder = get_or_create_document_folder(detalii)
            if eroare_folder:
                errors.append(f"{comanda_interna}: {eroare_folder}")
                continue

            if tip_document == 'RC':
                if detalii is None:
                    errors.append(f"{comanda_interna}: Missing details information")
                    continue
                reper = detalii.get('Reper', None)
                detalii_tehnologie, eroare_tehnologie = gaseste_detalii_tehnologie(reper)
                if eroare_tehnologie:
                    errors.append(f"{comanda_interna}: {eroare_tehnologie}")
                    continue
                
                tech_data = detalii_tehnologie if detalii_tehnologie else {}
                max_operations = 10
                operations_exist = False
                for i in range(1, max_operations + 1):
                    op_num_col = f'OP{i*10}'
                    operatie_text = tech_data.get(op_num_col, '')
                    if operatie_text and str(operatie_text).strip() not in ('', 'nan'):
                        operations_exist = True
                        break
                if not operations_exist:
                    errors.append(f"{comanda_interna}: Nu există operații pentru reperul în Tehnologii.xlsx.")
                    continue
                
                mesaj, succes = genereaza_route_card_excel(detalii, folder_path)
                if succes:
                    successes.append(f"{comanda_interna}")
                    # Extract file path from success message - handle both "salvat" and "salvată"
                    if "salvat în: " in mesaj:
                        file_path = mesaj.split("salvat în: ")[-1]
                    elif "salvată în: " in mesaj:
                        file_path = mesaj.split("salvată în: ")[-1]
                    else:
                        file_path = mesaj.split(": ")[-1]  # Fallback
                    generated_files_for_email.append((comanda_interna, file_path))
                else:
                    errors.append(f"{comanda_interna}: {mesaj}")
            
            elif tip_document == 'COC':
                defaults = build_coc_defaults(comanda_interna, detalii)
                date_suplimentare = defaults.copy()
                date_suplimentare['Lot Material client'] = lot_material_client if lot_material_client is not None else ''
                date_suplimentare['Nume Client'] = nume_client
                
                mesaj, succes = genereaza_declaratie_conformitate_excel(detalii, date_suplimentare, folder_path)
                if succes:
                    successes.append(f"{comanda_interna}")
                    # Extract file path from success message - handle both "salvat" and "salvată"
                    if "salvat în: " in mesaj:
                        file_path = mesaj.split("salvat în: ")[-1]
                    elif "salvată în: " in mesaj:
                        file_path = mesaj.split("salvată în: ")[-1]
                    else:
                        file_path = mesaj.split(": ")[-1]  # Fallback
                    generated_files_for_email.append((comanda_interna, file_path))
                else:
                    errors.append(f"{comanda_interna}: {mesaj}")

        # Show one message box at the end
        summary_parts = []
        if successes:
            summary_parts.append(f"Generare completă pentru {len(successes)} comenzi.")
            btn_email.setEnabled(True) # Enable email button
        else:
            btn_email.setEnabled(False) # Disable if no files were generated
        if errors:
            summary_parts.append(f"Erori la {len(errors)} comenzi:\n- " + "\n- ".join(errors))
        
        if summary_parts:
            QMessageBox.information(window, "Rezultat Generare Batch", "\n\n".join(summary_parts))
        else:
            QMessageBox.information(window, "Rezultat Generare Batch", "Nicio comandă nu a fost procesată.")

    def load_orders():
        """Load orders from the Excel file and populate the list."""
        orders_list.clear()
        try:
            df = pd.read_excel(NUME_FISIER_COMENZI, sheet_name='Comenzi', skiprows=[0])
            df.columns = df.columns.str.strip()
            comenzi = []
            if 'Comanda Interna' in df.columns:
                comenzi = df['Comanda Interna'].dropna().astype(str).str.strip().unique()
                for comanda in comenzi:
                    if comanda:
                        orders_list.addItem(QListWidgetItem(comanda))
            status_bar.setText(f"S-au încărcat {len(comenzi)} comenzi.")
        except Exception as e:
            QMessageBox.critical(window, "Eroare la încărcare", f"Nu s-a putut citi fișierul de comenzi:\n{e}")

    def load_orders():
        """Load orders from the Excel file and populate the list."""
        orders_list.clear()
        try:
            df = pd.read_excel(NUME_FISIER_COMENZI, sheet_name='Comenzi', skiprows=[0])
            df.columns = df.columns.str.strip()
            comenzi = []
            if 'Comanda Interna' in df.columns:
                comenzi = df['Comanda Interna'].dropna().astype(str).str.strip().unique()
                for comanda in comenzi:
                    if comanda:
                        orders_list.addItem(QListWidgetItem(comanda))
            status_bar.setText(f"S-au încărcat {len(comenzi)} comenzi.")
        except Exception as e:
            QMessageBox.critical(window, "Eroare la încărcare", f"Nu s-a putut citi fișierul de comenzi:\n{e}")

    # Connect signals
    search_input.textChanged.connect(filter_orders)
    refresh_btn.clicked.connect(load_orders)
    btn_generate.clicked.connect(start_batch_generation)
    btn_email.clicked.connect(send_email_with_attachments)
    
    # Initial load of orders
    load_orders()
    
    # Start the application
    window.show()
    app.exec()


def ruleaza_aplicatia_gui():
    # Import tkinter locally so the module can be imported headless; provide a clear
    # message and return if tkinter is not available.
    try:
        import tkinter as tk
        from tkinter import messagebox, simpledialog, filedialog
    except Exception:
        print("Tkinter nu este disponibil. Rulați în modul CLI cu --nogui sau instalați tkinter.")
        return
    """
    Launches the graphical interface (GUI) for the program using Tkinter.
    Allows the user to enter order details and generate documents with buttons and dialogs.
    """
    root = tk.Tk()
    root.title("RC & COC Generator")
    root.geometry("460x420")
    # Grey palette for tkinter fallback
    pastel_bg = '#F3F4F6'
    pastel_card = '#FFFFFF'
    pastel_btn = '#D1D5DB'
    text_color = '#1F2937'
    root.configure(bg=pastel_bg)

    def start_generation():
        text = entry_comanda.get().strip().upper()
        if not text:
            messagebox.showerror("Eroare", "Introduceți Comanda Internă.")
            return
        # Split by comma and clean up each order
        comenzi = [c.strip() for c in text.split(',') if c.strip()]
        if not comenzi:
            messagebox.showerror("Eroare", "Introduceți cel puțin o Comandă Internă validă.")
            return
        # Get document type once for all orders
        tip_document = var_tip_doc.get()
        # For COCs in batch, collect Lot Material client once and reuse
        coc_data = None
        if tip_document == "COC" and len(comenzi) > 0:
            # Show one dialog for Lot Material client
            resp = simpledialog.askstring("Lot Material client", "Introduceți Lot Material client pentru toate comenzile:", parent=root)
            if resp is not None:  # User clicked OK
                # First order's data as template
                first_order = comenzi[0]
                detalii_first, _ = gaseste_detalii_comanda(first_order)
                if detalii_first:
                    coc_data = build_coc_defaults(first_order, detalii_first)
                    coc_data['Lot Material client'] = resp
        # Process each order
        for comanda_interna in comenzi:
            # Get order details from Excel
            detalii, eroare_cautare = gaseste_detalii_comanda(comanda_interna)
            if eroare_cautare:
                messagebox.showerror("Eroare", f"Eroare pentru {comanda_interna}: {eroare_cautare}")
                continue
            if tip_document == "RC":
                # Check if operations exist before generating RC
                if detalii is None:
                    messagebox.showerror("Eroare", f"Missing details information for {comanda_interna}")
                    continue
                reper = detalii.get('Reper', None)
                detalii_tehnologie, eroare_tehnologie = gaseste_detalii_tehnologie(reper)
                if eroare_tehnologie:
                    # Message box already shown in gaseste_detalii_tehnologie
                    continue
                # Check if there are any operations
                tech_data = detalii_tehnologie if detalii_tehnologie else {}
                max_operations = 10
                operations_exist = False
                for i in range(1, max_operations + 1):
                    op_num_col = f'OP{i*10}'
                    operatie_text = tech_data.get(op_num_col, '')
                    if operatie_text and str(operatie_text).strip() not in ('', 'nan'):
                        operations_exist = True
                        break
                if not operations_exist:
                    messagebox.showerror("Eroare", f"Nu există operații pentru reperul din {comanda_interna} în Tehnologii.xlsx. Fișierul nu va fi generat.")
                    continue
                folder_path, eroare_folder = get_or_create_document_folder(detalii)
                if eroare_folder:
                    messagebox.showerror("Eroare", f"Eroare folder pentru {comanda_interna}: {eroare_folder}")
                    continue
                mesaj, succes = genereaza_route_card_excel(detalii, folder_path)
                if not succes:
                    messagebox.showerror("Eroare", f"Eroare generare pentru {comanda_interna}: {mesaj}")
                    continue
            elif tip_document == "COC":
                if coc_data is None:
                    # Each order will prompt separately if no common lot material client
                    defaults = build_coc_defaults(comanda_interna, detalii)
                    # Use simplified dialog with just Lot Material client field
                    lot_mat_client = simpledialog.askstring("Lot Material client", 
                                                        f"Introduceți Lot Material client pentru {comanda_interna}:", 
                                                        parent=root)
                    if lot_mat_client is None:  # User cancelled
                        continue
                    date_suplimentare = defaults.copy()
                    date_suplimentare['Lot Material client'] = lot_mat_client
                else:
                    # Use the shared COC data with updated certificate number for this order
                    date_suplimentare = coc_data.copy()
                    new_defaults = build_coc_defaults(comanda_interna, detalii)
                    date_suplimentare['Nr. Certificat'] = new_defaults['Nr. Certificat']
                    date_suplimentare['Lot Nr.'] = new_defaults['Lot Nr.']
                folder_path, eroare_folder = get_or_create_document_folder(detalii)
                if eroare_folder:
                    messagebox.showerror("Eroare", f"Eroare folder pentru {comanda_interna}: {eroare_folder}")
                    continue
                mesaj, succes = genereaza_declaratie_conformitate_excel(detalii, date_suplimentare, folder_path)
                if not succes:
                    messagebox.showerror("Eroare", f"Eroare generare pentru {comanda_interna}: {mesaj}")
                    continue
            else:
                messagebox.showerror("Eroare", "Selectați tipul de document.")
                return
            # Show success for each order
            if succes:
                messagebox.showinfo("Succes", f"{comanda_interna}: {mesaj}")

    # UI Elements for the main window
    # Card frame
    card = tk.Frame(root, bg=pastel_card, bd=0, relief='flat', padx=12, pady=12)
    card.pack(padx=16, pady=16, fill='both', expand=True)

    tk.Label(card, text="Comanda Internă:", bg=pastel_card, fg=text_color).pack(anchor='w', pady=(4,2))
    entry_comanda = tk.Entry(card, width=36)
    entry_comanda.pack()

    tk.Label(card, text="Tip Document:", bg=pastel_card, fg=text_color).pack(anchor='w', pady=(10,2))
    var_tip_doc = tk.StringVar(value="RC")
    radio_frame = tk.Frame(card, bg=pastel_card)
    radio_frame.pack(anchor='w')
    tk.Radiobutton(radio_frame, text="Route Card (RC)", variable=var_tip_doc, value="RC", bg=pastel_card, fg=text_color).pack(side='left', padx=6)
    tk.Radiobutton(radio_frame, text="Declarație Conformitate (COC)", variable=var_tip_doc, value="COC", bg=pastel_card, fg=text_color).pack(side='left', padx=6)

    btn = tk.Button(card, text="Generează Document", command=start_generation, bg=pastel_btn, fg=text_color, padx=8, pady=6)
    btn.pack(pady=14)

    # Orders controls: load orders from Excel and allow multi-select generation
    orders_container = tk.Frame(card, bg=pastel_card)
    orders_container.pack(fill='x', pady=(6,0))
    tk.Label(orders_container, text='Selectați comenzile (multi-select):', bg=pastel_card, fg=text_color).pack(anchor='w')
    orders_listbox = tk.Listbox(orders_container, selectmode='extended', height=6, width=60)
    orders_listbox.pack()

    def on_selection_change(event):
        selected = [orders_listbox.get(i) for i in orders_listbox.curselection()]
        if selected:
            entry_comanda.delete(0, tk.END)
            entry_comanda.insert(0, ', '.join(selected))

    orders_listbox.bind('<<ListboxSelect>>', on_selection_change)

    def load_orders_tk():
        try:
            path = actualizeaza_cale_fisier(NUME_FISIER_COMENZI)
            df = pd.read_excel(path, sheet_name='Comenzi', skiprows=[0])
            df.columns = df.columns.str.strip()
            if 'Comanda Interna' in df.columns:
                df['Comanda Interna'] = df['Comanda Interna'].fillna('').astype(str).str.strip()
                orders = [c for c in df['Comanda Interna'].unique() if c]
                orders_listbox.delete(0, tk.END)
                for o in orders:
                    orders_listbox.insert(tk.END, str(o))
                messagebox.showinfo('Info', f'Au fost încărcate {len(orders)} comenzi.')
        except Exception as e:
            messagebox.showerror('Eroare', f'Eroare la încărcarea comenzilor: {e}')

    def run_selected_tk():
        selected_indices = orders_listbox.curselection()
        if not selected_indices:
            messagebox.showerror('Eroare', 'Selectați cel puțin o comandă.')
            return
        tip = 'RC' if var_tip_doc.get()=='RC' else 'COC'
        for idx in selected_indices:
            com = orders_listbox.get(idx)
            if tip == 'COC':
                defaults = build_coc_defaults(com)
                lotmat = simpledialog.askstring('Lot Material client', f'Introduceți Lot Material client pentru {com} (lăsați gol pentru implicit):', parent=root)
                if not lotmat:
                    lotmat = ''
                date_supl = defaults.copy()
                date_supl['Lot Material client'] = lotmat
                # Provide date_suplimentare explicitly
                run_order(com, tip='COC', skip_prompts=True, date_suplimentare=date_supl)
            else:
                run_order(com, tip='RC', skip_prompts=False)
        refresh_logs_tk()

    tk.Button(orders_container, text='Încarcă Comenzi', command=load_orders_tk, bg=pastel_btn, fg=text_color).pack(side='left', padx=6, pady=6)
    tk.Button(orders_container, text='Generează pentru selecție', command=run_selected_tk, bg=pastel_btn, fg=text_color).pack(side='left', padx=6, pady=6)

    # Refresh sources button
    def refresh_sources_tk():
        global NUME_FISIER_COMENZI, NUME_FISIER_TEHNOLOGII
        NUME_FISIER_COMENZI = verifica_si_selecteaza_fisier(NUME_FISIER_COMENZI, 'comenzi') or NUME_FISIER_COMENZI
        NUME_FISIER_TEHNOLOGII = verifica_si_selecteaza_fisier(NUME_FISIER_TEHNOLOGII, 'tehnologii') or NUME_FISIER_TEHNOLOGII
        lbl_sources.config(text=f"Comenzi: {NUME_FISIER_COMENZI}    Tehnologii: {NUME_FISIER_TEHNOLOGII}")
    tk.Button(card, text='Refresh Sources', command=refresh_sources_tk).pack(pady=(6,0))

    tk.Label(card, text='Recent runs:', bg=pastel_card, fg=text_color).pack(anchor='w', pady=(8,0))
    listbox = tk.Listbox(card, height=6, width=80)
    listbox.pack()

    def refresh_logs_tk():
        listbox.delete(0, tk.END)
        entries = read_log_entries(10)
        for e in entries:
            txt = f"{e.get('ts_end','?')} - {e.get('order','?')} - {e.get('status','?')}"
            listbox.insert(tk.END, txt)

    lbl_sources = tk.Label(card, text=f"Comenzi: {NUME_FISIER_COMENZI}    Tehnologii: {NUME_FISIER_TEHNOLOGII}", bg=pastel_card, fg=text_color)
    lbl_sources.pack(anchor='w', pady=(6,0))

    refresh_logs_tk()

    root.mainloop()


#################################################################
# 5. PROGRAM ENTRY POINT
#
# === HOW TO RUN THIS PROGRAM ===
#
# 1. Create a virtual environment (if not already created):
#    python -m venv venv
#
# 2. Activate the virtual environment:
#    On Windows:
#        venv\Scripts\activate
#    On macOS/Linux:
#        source venv/bin/activate
#
# 3. Install required packages:
#    pip install pandas xlsxwriter openpyxl
#
# 4. Run the program:
#    python route_card_coc_app.py
# ===============================
#################################################################
# =============================================================
# 7. MAIN ENTRY POINT
# =============================================================

if __name__ == "__main__":
    # Entry point for the program
    # Checks if required packages are installed, then launches the GUI
    try:
        import pandas
        import xlsxwriter
        import openpyxl # openpyxl is needed by pandas for reading .xlsx files
    except ImportError as e:
        print("\nEROARE: Lipsesc dependințe esențiale (pandas/xlsxwriter/openpyxl).")
        print("Vă rugăm instalați pachetele necesare rulând în terminal:")
        print("pip install pandas xlsxwriter openpyxl")
        exit(1)
    
    # Apply UI contrast fixes before creating the real QApplication / launching GUI
    ensure_ui_contrast()
    
    # Default behavior: GUI. But allow CLI mode for headless use.
    parser = argparse.ArgumentParser(description="RC & COC Generator - GUI/CLI")
    parser.add_argument('--nogui', action='store_true', help='Run in non-GUI (CLI) mode')
    parser.add_argument('--comanda', type=str, help='Comanda Interna (order id) to process in CLI mode')
    parser.add_argument('--tip', type=str, choices=['RC', 'COC'], default='RC', help='Tip document to generate in CLI mode')
    parser.add_argument('--skip-prompts', action='store_true', help='Skip interactive prompts and use defaults for missing COC fields')
    parser.add_argument('--batch', type=str, help='Path to newline-delimited file with Comanda Interna values for batch processing')
    args = parser.parse_args()

    def ruleaza_aplicatia_cli(cli_args):
        """Run a non-GUI flow: fetch order, create folder, and generate RC or COC.
        Supports single --comanda or --batch file with multiple lines.
        Logs each run as JSON lines to rc_coc_runs.jsonl in the current directory.
        """
        # Prepare log file
        log_file = Path(os.getcwd()) / 'rc_coc_runs.jsonl'

        def log_run(entry: dict):
            try:
                with log_file.open('a', encoding='utf-8') as fh:
                    fh.write(json.dumps(entry, ensure_ascii=False) + "\n")
           

            except Exception:
                pass

        def process_one(comanda_interna: str):
            comanda_interna = comanda_interna.strip().upper()
            started = datetime.now().isoformat()
            detalii, eroare = gaseste_detalii_comanda(comanda_interna)
            if eroare:
                entry = {'order': comanda_interna, 'status': 'ERROR', 'message': eroare, 'ts_start': started, 'ts_end': datetime.now().isoformat()}
                print(f"EROARE: {eroare}")
                log_run(entry)
                return
            folder_path, eroare_folder = get_or_create_document_folder(detalii)
            if eroare_folder:
                entry = {'order': comanda_interna, 'status': 'ERROR', 'message': eroare_folder, 'ts_start': started, 'ts_end': datetime.now().isoformat()}
                print(f"EROARE: {eroare_folder}")
                log_run(entry)
                return
            if cli_args.tip == 'RC':
                mesaj, succes = genereaza_route_card_excel(detalii, folder_path)
                status = 'OK' if succes else 'ERROR'
                entry = {'order': comanda_interna, 'status': status, 'message': mesaj, 'ts_start': started, 'ts_end': datetime.now().isoformat()}
                print(mesaj)
                log_run(entry)
                return

            # COC path
            if cli_args.skip_prompts:
                date_suplimentare = {
                    'Nr. Certificat': f"DCIR{comanda_interna[3:]}",
                    'Lot Nr.': comanda_interna[-2:],
                    'Lot Material client': '',
                    'Revizie Desen': detalii.get('Revizie', 'N/A') if detalii is not None and isinstance(detalii.get('Revizie', None), str) else 'N/A',
                    'Nume Client': 'Elmet International SRL'
                }
            else:
                date_suplimentare = cere_date_suplimentare_coc(comanda_interna)
            mesaj, succes = genereaza_declaratie_conformitate_excel(detalii, date_suplimentare, folder_path)
            status = 'OK' if succes else 'ERROR'
            entry = {'order': comanda_interna, 'status': status, 'message': mesaj, 'ts_start': started, 'ts_end': datetime.now().isoformat()}
            print(mesaj)
            log_run(entry)

        # Execute batch or single
        if cli_args.batch:
            batch_path = Path(cli_args.batch)
            if not batch_path.exists():
                print(f"EROARE: Batch file '{batch_path}' nu a fost gasit.")
                return
            with batch_path.open('r', encoding='utf-8') as fh:
                lines = [l.strip() for l in fh if l.strip()]
            for line in lines:
                process_one(line)
            return

        if not cli_args.comanda:
            print("EROARE: În modul CLI trebuie specificată opțiunea --comanda COMANDA_INTERNA sau --batch BATCHFILE")
            return
        process_one(cli_args.comanda)

    if args.nogui:
        ruleaza_aplicatia_cli(args)
    else:
        ruleaza_aplicatia_pyqt()
