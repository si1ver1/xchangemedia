import os
import zipfile

try:
    import httplib
except:
    import http.client as httplib

def internet_on():
    # Checks network connection by initiating a HEAD request to google.com
    conn = httplib.HTTPConnection("www.google.com", timeout=5)
    try:
        conn.request("HEAD", "/")
        conn.close()
        return True
    except:
        conn.close()
        return False

def extract_zip():
    # Extract data.zip file
    zip_src = 'data.zip'
    zip_target = './'
    with open('InstallerLog.txt', 'a') as f:
        print('Extracting zip file...', file=f)
    zip_ref = zipfile.ZipFile(zip_src, 'r')
    zip_ref.extractall(zip_target)
    zip_ref.close()

def copy_files():
    # Copy XchangeMedia folder
    from distutils.dir_util import copy_tree
    fromDirectory = './XchangeMedia'
    toDirectory = 'C:\\/XchangeMedia'
    copy_tree(fromDirectory, toDirectory)
    with open('InstallerLog.txt', 'a') as f:
        print('Moving XchangeMedia folder to C: drive', file=f)

    # Copy XchangeMediaLauncher folder
    fromDirectory2 = './XchangeMediaLauncher'
    toDirectory2 = 'C:\\XchangeMediaLauncher'
    copy_tree(fromDirectory2, toDirectory2)
    with open('InstallerLog.txt', 'a') as f:
        print('Moving XchangeMediaLaunch folder to C: drive', file=f)

def power_settings():
    import subprocess
    # Set power settings to high performance
    # Then set computer and display to never sleep
    subprocess.call(['C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe', 'powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'])
    subprocess.call(['C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe', 'powercfg -setacvalueindex 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c 238c9fa8-0aad-41ed-83f4-97be242c8f20 29f6c1db-86da-48c5-9fdb-f2b67b1f44da 0'])
    subprocess.call(['C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe', 'powercfg -setacvalueindex 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c 7516b95f-f776-4464-8c53-06167f40cc99 3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e 0'])
    # Create subdirectory for user folder
    # Copy Group Policy settings
    subprocess.call(['C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe', '-ExecutionPolicy', 'Unrestricted', '-ExecutionPolicy', 'remotesigned', 'new-item -itemtype directory -path C:\\Windows\\System32\\GroupPolicy\\User\\'])
    subprocess.call(['C:\\WINDOWS\\system32\\WindowsPowerShell\\v1.0\\powershell.exe', '-ExecutionPolicy', 'Unrestricted', '-ExecutionPolicy', 'remotesigned', 'copy-item C:\\Users\\adamw\\Desktop\\xchangemedia\\Policies\\Registry.pol C:\\Windows\\System32\\GroupPolicy\\User\\'])

def run_ninite():
    # Install ninite bundle: Chrome, .Net 4, Teamviewer
    os.startfile('NiniteInstaller.exe')
    with open('InstallerLog.txt', 'a') as f:
        print('Running Ninite installer...installing: Chrome, .Net 4, and TeamViewer.', file=f)

def log():
    with open('InstallerLog.txt', 'a') as f:
        print('Starting automated setup for XchangeMedia', file=f)
        print('Make sure you are connected to the internet for the next steps.', file=f)
        print('Checking network connection:', file=f)
        print(internet_on(), file=f)

log()
extract_zip()
copy_files()
run_ninite()
power_settings()
