import sys, py7zr, os, json, tempfile, subprocess
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, QMessageBox, QProgressBar
from PyQt5.QtCore import Qt

def installNMI(file):
    name = os.path.basename(file)
    name = name.replace('.nmi', '')
    print(f'File = {file}')

    appname = os.path.basename(file)
    appname = appname.replace('.nmi', '')
    print(f'Appname = {appname}')
    
    window.bar.setValue(20)
    
    with py7zr.SevenZipFile(f"{file}", 'r', password=r"m1FC%0%0", header_encryption=True) as archive:
        archive.extractall(path=f"{os.getenv('LOCALAPPDATA')}\\")
    
    print("Extraction has finished.")
    
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
            # Add your installation logic here
            #print(f"Installing file: {self.selected_file_path}")
            # For demonstration purposes, just printing the selected file path
            self.install_button.setEnabled(False)
            window.tutorial.setText("Tutorial: Installing...")
            installNMI(self.selected_file_path)


if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = FileInstallerApp()
        window.show()
    except Exception as e:
        print(f"An error occurred: {e}")
    sys.exit(app.exec_())