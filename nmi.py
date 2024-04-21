import sys, py7zr, os
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, QMessageBox, QProgressBar

def installNMI(file):
    print(file)
    with py7zr.SevenZipFile(f"{file}", 'r', password=r"m1FC%0%0", header_encryption=True) as archive:
        archive.extractall(path=f"{os.getenv('LOCALAPPDATA')}\\")
    print("Installation has finished.")
    window.bar.setValue(100)
    informer(text=f"Installation of '{os.path.basename(file)}' has finished.", title="NMI")
    

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
        self.layout.addWidget(self.bar)

        self.selected_file_path = None
    
    def browse_file(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setNameFilter("NMI files (*.nmi)")
        if file_dialog.exec_():
            self.selected_file_path = file_dialog.selectedFiles()[0]
            self.file_label.setText(f"Selected file: {self.selected_file_path}")
            self.install_button.setEnabled(True)

    def install_file(self):
        if self.selected_file_path:
            # Add your installation logic here
            #print(f"Installing file: {self.selected_file_path}")
            # For demonstration purposes, just printing the selected file path
            self.install_button.setEnabled(False)
            installNMI(self.selected_file_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileInstallerApp()
    window.show()
    sys.exit(app.exec_())