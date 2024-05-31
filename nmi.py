import sys, py7zr, os, json, tempfile, subprocess, argparse, sys, winreg
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, QMessageBox, QProgressBar
from PyQt5.QtCore import Qt

def installNMI(file, minimal=False):
    print("Parsing...")

    name = os.path.basename(file)
    name = name.replace('.nmi', '')
    print(f'File = {file}')

    appname = os.path.basename(file)
    appname = appname.replace('.nmi', '')
    print(f'NMI Name = {appname}')
    
    if not minimal:
        window.bar.setValue(20)
    
    # NOTICE. The password for NMI files are stored in plain text.
    with py7zr.SevenZipFile(f"{file}", 'r', password=r"m1FC%0%0", header_encryption=True) as archive:
        archive.extractall(path=f"{os.getenv('LOCALAPPDATA')}\\")
    
    print("Extraction has finished.")
    
    if not minimal:
        window.bar.setValue(80)
    
    with open(f"{os.getenv('LOCALAPPDATA')}\\{appname}\\nmi.iinfo", 'r') as file:
        data = json.load(file)
    execname = data["execname"]
    path_to_exec = f"{os.getenv('LOCALAPPDATA')}\\{appname}\\{execname}"
    shortcut_linker(appname, path_to_exec)
    print("Shortcut created!")

    # Copy to start menu
    home_loc = os.path.expanduser('~')
    desktop_loc = os.path.expanduser('~\\Desktop')
    shortcut_src = f"{desktop_loc}\{appname}.lnk"
    shortcut_dest = f"{home_loc}\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"

    os.system(f"copy \"{shortcut_src}\" \"{shortcut_dest}\"")

    if not minimal:
        window.tutorial.setText("Tutorial: Installed successfully!")

        window.bar.setValue(100)

        informer(text=f"Installation of '{appname}' has finished.", title="NMI")

def shortcut_linker(name: str, targetpath: str):
    SCRIPT_TEMPLATE = """
    Set oWS = WScript.CreateObject("WScript.Shell") 
    sLinkFile = "{}"
    Set oLink = oWS.CreateShortcut(sLinkFile)
    oLink.TargetPath = "{}"
    oLink.Save
    """

    desktop_loc = os.path.expanduser("~\\Desktop")
    name_path = os.path.join(desktop_loc, f"{name}.lnk")

    with tempfile.TemporaryDirectory() as tmpdir:
        vbs_file = os.path.join(tmpdir, "CreateShortcut.vbs")
        with open(vbs_file, "w") as fout:
            fout.write(SCRIPT_TEMPLATE.format(name_path, targetpath))
        try:
            subprocess.call(["cscript", vbs_file], cwd=tmpdir)
        except Exception as ex:
            print("Error creating shortcut")
            print(ex)

def informer(icon="info", text="info", title="title"):
    informer = QMessageBox() 
    if icon == "info":
        informer.setIcon(QMessageBox.Information) 
    informer.setText(text) 
    informer.setWindowTitle(title)
    informer.setStandardButtons(QMessageBox.Ok) 
    retval = informer.exec_() 

class FileInstallerApp(QMainWindow):
    def __init__(self):
        global bar
        super().__init__()

        self.setWindowTitle("NMI")
        self.setGeometry(100, 100, 400, 200)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        self.tutorial = QLabel("Tutorial: Click the 'Browse' button and navigate to the '.nmi' file.")
        self.layout.addWidget(self.tutorial)

        self.file_label = QLabel("No file selected")
        self.layout.addWidget(self.file_label)

        self.browse_button = QPushButton("Browse")
        self.browse_button.clicked.connect(self.browse_file)
        self.layout.addWidget(self.browse_button)

        self.install_button = QPushButton("Install")
        self.install_button.clicked.connect(self.install_file)
        self.install_button.setEnabled(False)
        self.layout.addWidget(self.install_button)

        self.bar = QProgressBar(self)
        self.bar.setAlignment(Qt.AlignCenter) 
        self.layout.addWidget(self.bar)
        self.bar.setValue(0)

        self.selected_file_path = None

    def browse_file(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setNameFilter("NMI files (*.nmi)")
        if file_dialog.exec_():
            self.selected_file_path = file_dialog.selectedFiles()[0]
            self.file_label.setText(f"Selected file: {self.selected_file_path}")
            self.install_button.setEnabled(True)
        window.bar.setValue(0)
        window.tutorial.setText("Tutorial: Click the install button to install the currently selected '.nmi' file.")

    def install_file(self):
        if self.selected_file_path:
            self.install_button.setEnabled(False)
            window.tutorial.setText("Tutorial: Installing...")
            installNMI(self.selected_file_path)

def searchPath(query):
    path_env = os.getenv('PATH')

    paths = path_env.split(os.pathsep)

    found_paths = []
    for path in paths:
        if query in path:
            found_paths.append(path)

    return found_paths

def add_to_user_path(new_path):
    try:
        user_env_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Environment', 0, winreg.KEY_ALL_ACCESS)
        
        current_path, _ = winreg.QueryValueEx(user_env_key, 'PATH')
        
        if new_path not in current_path:
            new_path = new_path.rstrip(os.pathsep)
            new_path = current_path + os.pathsep + new_path
            
            winreg.SetValueEx(user_env_key, 'PATH', 0, winreg.REG_EXPAND_SZ, new_path)

        winreg.CloseKey(user_env_key)
    except Exception as e:
        print("Error:", e)

def set_file_association():
    executable_path = "nmi"

    script_path = os.path.abspath(sys.argv[0])

    command = f'"{executable_path}" --file "%1"'

    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 'Software\Classes', 0, winreg.KEY_WRITE)

        winreg.SetValueEx(key, '.nmi', 0, winreg.REG_SZ, 'NMIFile')

        key_nmi = winreg.CreateKey(key, 'NMIFile')
        winreg.SetValueEx(key_nmi, '', 0, winreg.REG_SZ, 'NMI File')
        
        key_shell = winreg.CreateKey(key_nmi, 'shell')

        key_open = winreg.CreateKey(key_shell, 'open')

        winreg.SetValueEx(key_open, '', 0, winreg.REG_SZ, None)

        key_command = winreg.CreateKey(key_open, 'command')
        winreg.SetValueEx(key_command, '', 0, winreg.REG_SZ, command)

        winreg.CloseKey(key)
        print("File association set successfully.")

    except Exception as e:
        print("Error occurred while setting file association:", e)

if __name__ == "__main__":

    print(sys.executable)

    # Add NMI to path.
    query = sys.executable
    result = searchPath(query)
    if result:
        for path in result:
            if path != query:
                print("NOTICE: Adding NMI to your user Path variable. If you have moved NMI many times your path variable could be filled with many directorys. This is for file associations to work.")
                add_to_user_path(os.getcwd())
                set_file_association()

    parser = argparse.ArgumentParser(description="Install NMI file")
    parser.add_argument("--file", help="NMI file to install", required=False)
    args = parser.parse_args()
    if args.file:
        installNMI(args.file, True)
        exit()
    try:
        app = QApplication(sys.argv)
        window = FileInstallerApp()
        window.show()
    except Exception as e:
        print(f"An error occurred: {e}")
    sys.exit(app.exec_())
