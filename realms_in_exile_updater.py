import os
import sys
import requests
from tqdm import tqdm
import zipfile
import winreg
import win32com.client

BASE_URL = "https://storage.googleapis.com/realms-in-exile/updater/"
VERSION_FILE = "version.txt"
MOD_FILE = "files.zip"

# Check if running as a PyInstaller bundle
if getattr(sys, 'frozen', False):
    # If bundle, set the base directory to the executable's directory
    base_dir = sys._MEIPASS
else:
    # Otherwise, use the script's directory
    base_dir = os.path.dirname(os.path.abspath(__file__))

def download_file(filename):
    file_path = os.path.join(script_directory, filename)
    url = BASE_URL + filename
    response = requests.get(url, stream=True)
    total_size = int(response.headers.get('content-length', 0))
    block_size = 1024
    t = tqdm(total=total_size, unit='B', unit_scale=True, desc=filename)
    with open(file_path, 'wb') as f:
        for data in response.iter_content(block_size):
            t.update(len(data))
            f.write(data)
    t.close()
    
def get_lotr_install_path():
    key_path = r"SOFTWARE\WOW6432Node\Electronic Arts\Electronic Arts\The Lord of the Rings, The Rise of the Witch-king"
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
            install_path, _ = winreg.QueryValueEx(key, "InstallPath")
            return install_path
    except FileNotFoundError:
        print("The specified registry key or value was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")
    return None

def create_shortcut_with_flag(target_path, shortcut_path, flag, icon_path=None):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(shortcut_path)
    shortcut.TargetPath = target_path
    shortcut.Arguments = flag
    if icon_path:
        shortcut.IconLocation = icon_path
    shortcut.WindowStyle = 3  # 3 means "Maximized", 7 means "Minimized". 1 is "Normal"
    shortcut.save()

def create_shortcut_on_desktop():
    # After the mod is installed/updated, create the shortcut
    install_path = get_lotr_install_path()
    if install_path:
        print(f"Installation path of The Rise of the Witch-king is: {install_path}")
        exe_path = os.path.join(install_path, "lotrbfme2ep1.exe")
        if os.path.exists(exe_path):
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            local_version = get_local_version()
            shortcut_path = os.path.join(desktop, f"Realms in Exile ({local_version}).lnk")
            icon_path = os.path.join(script_directory, "aotr_fs.ico")
            create_shortcut_with_flag(exe_path, shortcut_path, f"-mod \"{script_directory}\"", icon_path)
            print(f"Shortcut created successfully!")
        else:
            print(f"'lotrbfme2ep1.exe' not found in {install_path}")
    else:
        print("Failed to retrieve the installation path.")

def create_updater_shortcut_on_desktop():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    updater_exe_path = sys.executable  # This gives the full path to the running executable
    updater_executable_name = os.path.basename(updater_exe_path)  # Extracts the file name from the full path
    updater_shortcut_path = os.path.join(desktop, f"{updater_executable_name}.lnk")
    icon_path = os.path.join(script_directory, "aotr_fs.ico")
    
    if os.path.exists(updater_exe_path):
        create_shortcut_with_flag(updater_exe_path, updater_shortcut_path, "", icon_path)
        print(f"Updater shortcut created successfully on the desktop: {updater_shortcut_path}")
    else:
        print(f"Updater executable not found: {updater_exe_path}")

def extract_with_progress(zip_path, extract_path):
    with zipfile.ZipFile(zip_path, 'r') as zf:
        # Get the total number of entries within the ZIP file
        total_files = len(zf.infolist())
        with tqdm(total=total_files, unit='file', desc="Extracting") as t:
            for member in zf.infolist():
                zf.extract(member, extract_path)
                t.update(1)

def get_local_version():
    with open(os.path.join(script_directory, VERSION_FILE), 'r') as f:
        return f.read().strip()

def get_online_version():
    try:
        response = requests.get(BASE_URL + VERSION_FILE)
        response.raise_for_status()  # This will raise an error for 4xx and 5xx status codes
        return response.text.strip()
    except requests.RequestException as e:
        print(f"Error fetching online version: {e}\n")
        return None

def cleanup_and_retry():
    mod_file_path = os.path.join(script_directory, MOD_FILE)
    if os.path.exists(mod_file_path):
        os.remove(mod_file_path)
    choice = input("An error occurred. Would you like to start over? (Y/N): \n").upper()
    if choice == 'Y':
        main()

def main():
    global script_directory
    if getattr(sys, 'frozen', False):
        # If the script is run from a bundled executable, use the executable's directory
        script_directory = os.path.dirname(sys.executable)
    else:
        # Otherwise, use the script's directory
        script_directory = os.path.dirname(os.path.abspath(__file__))

    try:
        local_version = get_local_version() if os.path.exists(os.path.join(script_directory, VERSION_FILE)) else None
        online_version = get_online_version()

        if not local_version:
            if online_version is None:
                print("Unable to check for updates at this time.")
                return
            print(f"'Age of the Ring: Realms in Exile' is not installed.\n")
            print(f"Installing version {online_version} in: {script_directory}\n")
            download_file(VERSION_FILE)
            download_file(MOD_FILE)
            print("Installing files...")
            extract_with_progress(os.path.join(script_directory, MOD_FILE), script_directory)
            os.remove(os.path.join(script_directory, MOD_FILE))
            print(f"'Age of the Ring: Realms in Exile' version {online_version} was installed successfully!\n")
            create_shortcut_on_desktop()
            create_updater_shortcut_on_desktop()
            input("Press Enter to exit...")
            return

        if online_version is None:
            print("Unable to check for updates at this time.")
            return

        if local_version != online_version:
            print(f"Your version of 'Age of the Ring: Realms in Exile' is {local_version}.\nUpdating to version {online_version}...\n")
            download_file(VERSION_FILE)
            download_file(MOD_FILE)
            print("Updating files...")
            extract_with_progress(os.path.join(script_directory, MOD_FILE), script_directory)
            os.remove(os.path.join(script_directory, MOD_FILE))
            print(f"'Age of the Ring: Realms in Exile' was updated to version {online_version} successfully!\n")
            create_shortcut_on_desktop()
            create_updater_shortcut_on_desktop()
            input("Press Enter to exit...")
        else:
            print(f"You have the latest version of 'Age of the Ring: Realms in Exile' ({local_version}).\n")
            create_shortcut_on_desktop()
            create_updater_shortcut_on_desktop()
            input("Press Enter to exit...")

    except Exception as e:
        print(f"Error: {e}")
        cleanup_and_retry()

if __name__ == "__main__":
    main()
