import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QTextEdit, QProgressBar
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
import requests
import os

class DownloadThread(QThread):
    progress_update = pyqtSignal(int)
    download_finished = pyqtSignal()

    def __init__(self, df, output_folder):
        super().__init__()
        self.df = df
        self.output_folder = output_folder

    def run(self):
        total_images = len(self.df)
        for index, row in self.df.iterrows():
            matricule = str(row['MATRICULE'])
            url = row['URL']

            try:
                # Télécharger l'image
                response = requests.get(url)

                # Construire le chemin complet pour sauvegarder l'image
                image_filename = os.path.join(self.output_folder, os.path.basename(url))

                # Sauvegarder l'image dans le dossier de sortie
                with open(image_filename, 'wb') as img_file:
                    img_file.write(response.content)

                self.progress_update.emit((index + 1) * 100 // total_images)  # Mettre à jour la barre de progression
                self.output_box.append(f"Téléchargée : {image_filename}")
            except Exception as e:
                self.output_box.append(f"Erreur pour {matricule} : {e}")

        self.download_finished.emit()

class ImageDownloaderApp(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Image Downloader')
        self.setGeometry(100, 100, 600, 400)

        self.setWindowIcon(QIcon('logo.png'))

        layout = QVBoxLayout()

        self.file_label = QLabel('Aucun fichier sélectionné')
        layout.addWidget(self.file_label)

        self.progress_bar = QProgressBar(self)
        layout.addWidget(self.progress_bar)

        self.output_box = QTextEdit()
        layout.addWidget(self.output_box)

        self.download_button = QPushButton('Télécharger Images')
        self.download_button.clicked.connect(self.start_download)
        layout.addWidget(self.download_button)

        self.setLayout(layout)

    def start_download(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setNameFilter("Fichiers Excel (*.xlsx)")
        if file_dialog.exec_():
            file_path = file_dialog.selectedFiles()[0]
            self.file_label.setText(f'Fichier sélectionné: {file_path}')

            # Charger le fichier Excel ici
            df = pd.read_excel(file_path)

            self.download_thread = DownloadThread(df, "photos")
            self.download_thread.progress_update.connect(self.progress_bar.setValue)
            self.download_thread.download_finished.connect(self.download_finished)
            self.download_thread.output_box = self.output_box  # Passer la référence à output_box
            self.download_thread.start()

    def download_finished(self):
        self.output_box.append("Téléchargement terminé.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    downloader_app = ImageDownloaderApp()
    downloader_app.show()
    sys.exit(app.exec_())
