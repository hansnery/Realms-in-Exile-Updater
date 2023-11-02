import os
import sys
import requests
from tqdm import tqdm
import zipfile

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
    os.remove(os.path.join(script_directory, MOD_FILE))
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
            return

        if online_version is None:
            print("Unable to check for updates at this time.")
            return

        if local_version != online_version:
            print(f"Your version of 'Age of the Ring: Realms in Exile' is {local_version}. Updating to version {online_version}...\n")
            download_file(VERSION_FILE)
            download_file(MOD_FILE)
            print("Updating files...")
            extract_with_progress(os.path.join(script_directory, MOD_FILE), script_directory)
            os.remove(os.path.join(script_directory, MOD_FILE))
            print(f"'Age of the Ring: Realms in Exile' was updated to version {online_version} successfully!\n")
        else:
            print(f"You have the latest version of 'Age of the Ring: Realms in Exile' ({local_version}).\n")

    except Exception as e:
        print(f"Error: {e}")
        cleanup_and_retry()

if __name__ == "__main__":
    main()
