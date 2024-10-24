# Author @ju4ncaa (Juan Carlos Rodríguez)

# Librerías
import os
import sys
import subprocess
from shutil import copy
from time import sleep

# Paleta de colores ANSI
class Colors:
    """ ANSI color codes """
    RED = "\033[0;31m"
    GREEN = "\033[0;32m"
    CYAN = "\033[0;36m"
    YELLOW = "\033[1;33m"
    RESET = "\033[0m"

# Banner
def banner():
    print(f"""{Colors.YELLOW}
     _         _         ___   __  __ _            ____   ___ ____  _  _   
    / \  _   _| |_ ___  / _ \ / _|/ _(_) ___ ___  |___ \ / _ \___ \| || |  
   / _ \| | | | __/ _ \| | | | |_| |_| |/ __/ _ \   __) | | | |__) | || |_  {Colors.CYAN}(Author @ju4ncaa [Juan Carlos Rodríguez]){Colors.YELLOW}
  / ___ \ |_| | || (_) | |_| |  _|  _| | (_|  __/  / __/| |_| / __/|__   _|
 /_/   \_\__,_|\__\___/ \___/|_| |_| |_|\___\___| |_____|\___/_____|  |_|                                                                        
    {Colors.RESET}"""); sleep(3)

banner()

# Crear carpeta Office2024 en Documents
user_documents = os.path.join(os.environ["USERPROFILE"], "Documents")

office_folder = os.path.join(user_documents, "Office24")

if not os.path.exists(office_folder):
    print(f"\n{Colors.YELLOW}[+]{Colors.RESET} Creando el directorio para la instalación\n"); sleep(2)
    os.makedirs(office_folder)
    print(f"\n{Colors.GREEN}[+]{Colors.RESET} Carpeta {Colors.CYAN}'{office_folder}'{Colors.RESET} creada.\n"); sleep(2)
else:
    print(f"\n{Colors.YELLOW}[!]{Colors.RESET} La carpeta {Colors.CYAN}'{office_folder}'{Colors.RESET} ya existe.\n")
    print (f"{Colors.RED}[!]{Colors.RESET} Saliendo...\n")
    sys.exit(1)

# Copiar archivos a la carpeta Office2024
script_folder = os.path.dirname(os.path.abspath(__file__)) 
files_to_copy = ["configuration.xml", "setup.exe"]

for file_name in files_to_copy:
    src_path = os.path.join(script_folder, file_name)
    dst_path = os.path.join(office_folder, file_name)
    
    if os.path.exists(src_path):
        copy(src_path, dst_path)
        print(f"{Colors.GREEN}[+]{Colors.RESET} Archivo {Colors.CYAN}'{file_name}'{Colors.RESET} copiado a {Colors.CYAN}'{office_folder}'{Colors.RESET}.\n"); sleep(2)
    else:
        print(f"{Colors.RED}[!]{Colors.RESET} Archivo {Colors.CYAN}'{file_name}'{Colors.RESET} no encontrado en la carpeta del script.\n")
        print (f"{Colors.RED}[!]{Colors.RESET} Saliendo...\n")
        sys.exit(1)

# Instalación Office
def office_install(command, folder):
    print (f"{Colors.GREEN}[+]{Colors.RESET} Comenzando instalación...\n"); sleep(2)
    # Verificar si el script ya está siendo ejecutado como administrador
    if os.name == 'nt':
        try:
            cmd_cd = f'cd /d "{folder}"'
            cmd_install = command
            
            subprocess.run(["powershell", "-Command", f"Start-Process cmd.exe -ArgumentList '/C {cmd_cd} && {cmd_install}' -Verb runAs"], check=True)
        except subprocess.CalledProcessError as e:
            print(f"\n{Colors.RED}[!]{Colors.RESET} Error al ejecutar el comando como administrador: {e}")
            sys.exit(1)
    else:
        print(f"\n{Colors.YELLOW}[!]{Colors.RESET} Este script solo funciona en Windows.")
        sys.exit(0)

command = 'setup /configure configuration.xml'
folder = office_folder

office_install(command, folder)
