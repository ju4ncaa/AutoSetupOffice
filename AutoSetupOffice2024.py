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

# Crear carpeta Office2024 en Documents
user_documents = os.path.join(os.environ["USERPROFILE"], "Documents")

office_folder = os.path.join(user_documents, "OfficePrueba")

if not os.path.exists(office_folder):
    os.makedirs(office_folder)
    print(f"\n{Colors.GREEN}[+]{Colors.RESET} Carpeta {Colors.CYAN}'{office_folder}'{Colors.RESET} creada.\n")
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
        print(f"{Colors.GREEN}[+]{Colors.RESET} Archivo {Colors.CYAN}'{file_name}'{Colors.RESET} copiado a {Colors.CYAN}'{office_folder}'{Colors.RESET}.\n")
    else:
        print(f"{Colors.RED}[!]{Colors.RESET} Archivo {Colors.CYAN}'{file_name}'{Colors.RESET} no encontrado en la carpeta del script.\n")
        print (f"{Colors.RED}[!]{Colors.RESET} Saliendo...\n")
        sys.exit(1)

# Instalación Office
def office_install(command, folder):
    
    cmd = f'cmd.exe /K "cd /d {folder} && {command}"'

    # Verificar si el script ya está siendo ejecutado como administrador
    if os.name == 'nt':
        try:
            # Ejecutar el comando como administrador en PowerShell usando cmd.exe
            subprocess.run(["powershell", "Start-Process", "cmd.exe", "-ArgumentList", f'/C "cd /d {folder} && {command}"', "-Verb", "runAs"])
        except Exception as e:
            print(f"\n{Colors.RED}[!]{Colors.RESET} Error al ejecutar el comando como administrador: {e}")
            print(f"{Colors.RED}[!]{Colors.RESET} Saliendo...\n")
            sys.exit(1)
    else:
        print(f"\n{Colors.YELLOW}[!]{Colors.RESET} Este script solo funciona en Windows.")
        print(f"{Colors.RED}[!]{Colors.RESET} Saliendo...\n")
        sys.exit(0)

# Ejemplo de uso:
command = 'setup /configure configuration.xml'
folder = 'C:\\Users\\admin\\Documents\\OfficePrueba'

office_install(command, folder)
