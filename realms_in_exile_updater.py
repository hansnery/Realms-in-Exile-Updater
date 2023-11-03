import ctypes
import os
import sys
import requests
from tqdm import tqdm
import zipfile
import winreg
import win32com.client

# Constants for file names and URLs
BASE_URL = "https://storage.googleapis.com/realms-in-exile/updater/"
VERSION_FILE = "version.txt"
MOD_FILE = "files.zip"
ICON_FILE = "aotr_fs.ico"
GAME_EXE = "lotrbfme2ep1.exe"

# Helper function to check if the script is running with admin privileges
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False
    
# Helper function to request admin privileges
def run_as_admin(argv=None, debug=False):
    shell32 = ctypes.windll.shell32
    if argv is None and shell32.IsUserAnAdmin():
        return True  # Already an admin
    if argv is None:
        argv = sys.argv
    if hasattr(sys, '_MEIPASS'):
        # Support pyinstaller wrapped program.
        arguments = argv[1:]
    else:
        arguments = argv
    argument_line = u' '.join(arguments)
    executable = sys.executable
    if debug:
        print('Command line:', executable, argument_line)
    ret = shell32.ShellExecuteW(None, u'runas', executable, argument_line, None, 1)
    if ret <= 32:
        return False
    return None

# Determine if we're running as a bundled application or as a normal script
if getattr(sys, 'frozen', False):
    # If it's a bundled application, use the sys.executable path
    application_path = os.path.dirname(sys.executable)
else:
    # If it's a normal script, use the __file__ path
    application_path = os.path.dirname(__file__)

# Helper functions
def download_file(filename):
    """Download a file from the BASE_URL with progress indication."""
    file_path = os.path.join(application_path, filename)
    url = BASE_URL + filename
    response = requests.get(url, stream=True)
    total_size = int(response.headers.get('content-length', 0))
    
    print(f"Downloading {filename} to {file_path}")
    
    with tqdm(total=total_size, unit='B', unit_scale=True, desc=filename) as t:
        with open(file_path, 'wb') as f:
            for data in response.iter_content(1024):
                t.update(len(data))
                f.write(data)

def get_registry_key_value(key_path, value_name):
    """Retrieve a value from the Windows Registry."""
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            return winreg.QueryValueEx(key, value_name)[0]
    except FileNotFoundError:
        print("The specified registry key or value was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

def create_shortcut(target_path, shortcut_path, arguments="", icon_path=None):
    """Create a shortcut with optional arguments and icon."""
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.TargetPath = target_path
    shortcut.Arguments = arguments
    if icon_path:
        shortcut.IconLocation = icon_path
    shortcut.WindowStyle = 1  # Normal window
    shortcut.save()

def extract_zip_with_progress(zip_path, extract_to_directory):
    """Extract a ZIP file with progress indication."""
    extract_path = os.path.join(application_path, extract_to_directory)
    with zipfile.ZipFile(zip_path, 'r') as zf, tqdm(total=len(zf.infolist()), unit='file', desc="Extracting") as t:
        for member in zf.infolist():
            zf.extract(member, extract_path)
            t.update(1)

# Main functions
def get_versions():
    """Get the local and online versions of the mod."""
    local_version_path = os.path.join(application_path, VERSION_FILE)
    local_version = open(local_version_path).read().strip() if os.path.exists(local_version_path) else None
    try:
        online_version = requests.get(BASE_URL + VERSION_FILE).text.strip()
    except requests.RequestException as e:
        print(f"Error fetching online version: {e}")
        online_version = None
    return local_version, online_version

def update_or_install(local_version, online_version):
    """Handle installation or updating of the mod."""
    if local_version and online_version and local_version != online_version:
        print(f"Updating from 'Age of the Ring: Realms in Exile' version {local_version} to version {online_version}...")
        download_file(VERSION_FILE)
        download_file(MOD_FILE)
        extract_zip_with_progress(os.path.join(application_path, MOD_FILE), application_path)
        os.remove(os.path.join(application_path, MOD_FILE))
        print(f"Update to 'Age of the Ring: Realms in Exile' version {online_version} complete.")
    elif not local_version and online_version:
        print(f"Installing 'Age of the Ring: Realms in Exile' version {online_version}...")
        download_file(VERSION_FILE)
        download_file(MOD_FILE)
        extract_zip_with_progress(os.path.join(application_path, MOD_FILE), application_path)
        os.remove(os.path.join(application_path, MOD_FILE))
        print(f"Installation of 'Age of the Ring: Realms in Exile' version {online_version} complete.")
    else:
        print("Your mod is up to date, no update or installation needed.")


def create_desktop_shortcuts():
    """Create desktop shortcuts for the mod and the updater."""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    install_path = get_registry_key_value(r"SOFTWARE\WOW6432Node\Electronic Arts\Electronic Arts\The Lord of the Rings, The Rise of the Witch-king", "InstallPath")
    
    if install_path and os.path.exists(os.path.join(install_path, GAME_EXE)):
        # Use the directory of the executable if the script is frozen, otherwise use the script directory.
        mod_directory = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else application_path
        game_exe_path = os.path.join(install_path, GAME_EXE)
        local_version, _ = get_versions()
        mod_shortcut_path = os.path.join(desktop_path, f"Realms in Exile ({local_version}).lnk")
        create_shortcut(game_exe_path, mod_shortcut_path, f"-mod \"{mod_directory}\"", os.path.join(mod_directory, ICON_FILE))
        print("Mod shortcut created on desktop.")
        
        updater_exe_path = sys.executable
        updater_shortcut_path = os.path.join(desktop_path, f"{os.path.basename(updater_exe_path)}.lnk")
        create_shortcut(updater_exe_path, updater_shortcut_path, "", os.path.join(mod_directory, ICON_FILE))
        print("Updater shortcut created on desktop.")
    else:
        print("Game installation path not found or game executable is missing.")

def main():
    """Main function to orchestrate the mod update or installation process."""
    if not is_admin():
        result = run_as_admin()
        if result is False:
            print("Failed to gain administrative privileges.")
        return  # Exit the script if not admin or after requesting admin privileges
    
    local_version, online_version = get_versions()
    if online_version is None:
        print("Unable to check for updates at this time.")
    else:
        update_or_install(local_version, online_version)
        create_desktop_shortcuts()
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
