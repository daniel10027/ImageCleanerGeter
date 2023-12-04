import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QProgressBar, QFileDialog, QTextEdit, QHBoxLayout, QSizePolicy
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
import requests
from io import BytesIO
from PIL import Image
import os

class DownloadThread(QThread):
    progress_update = pyqtSignal(int)
    image_downloaded = pyqtSignal(str, str, bytes, str)  # Ajout d'un signal pour le chemin de l'image
    download_finished = pyqtSignal()

    def __init__(self, df, output_folder):
        super().__init__()
        self.df = df
        self.output_folder = output_folder

    def run(self):
        for index, row in self.df.iterrows():
            nom = str(row['NOM'])
            prenom = str(row['PRENOMS'])
            matricule = str(row['MATRICULE'])
            contact1 = row['CONTACT 1']
            url = row['URL']

            response = requests.get(url, stream=True)
            total_size = int(response.headers.get('content-length', 0))

            img_data = BytesIO()
            progress = 0

            for chunk in response.iter_content(chunk_size=1024):
                img_data.write(chunk)
                progress += len(chunk)
                self.progress_update.emit(progress * 100 / total_size)

            # Émettre le signal indiquant que l'image a été téléchargée
            self.image_downloaded.emit(nom, prenom, img_data.getvalue(), url)

        # Émettre le signal indiquant que le téléchargement est terminé
        self.download_finished.emit()


class ImageDownloaderApp(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('Image Downloader')
        self.setGeometry(100, 100, 800, 400)

        self.setWindowIcon(QIcon('logo.png'))

        layout = QVBoxLayout()

        left_layout = QVBoxLayout()

        self.file_label = QLabel('Aucun fichier sélectionné')
        left_layout.addWidget(self.file_label)

        self.destination_label = QLabel('Aucun dossier sélectionné')
        left_layout.addWidget(self.destination_label)

        self.progress_bar = QProgressBar()
        left_layout.addWidget(self.progress_bar)

        self.download_button = QPushButton('Télécharger')
        self.download_button.clicked.connect(self.start_download)
        left_layout.addWidget(self.download_button)

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setFixedSize(225, 250)
        image_layout = QHBoxLayout()
        image_layout.addWidget(self.image_label)

        self.output_box = QTextEdit()

        main_layout = QHBoxLayout()
        main_layout.addLayout(left_layout)
        main_layout.addLayout(image_layout)
        main_layout.addWidget(self.output_box)

        self.setLayout(main_layout)

    def start_download(self):
        self.output_box.clear()

        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.AnyFile)
        file_dialog.setNameFilter("Fichiers Excel (*.xlsx)")
        if file_dialog.exec_():
            file_path = file_dialog.selectedFiles()[0]
            self.file_label.setText(f'Fichier sélectionné: {file_path}')

            df = pd.read_excel(file_path)

            destination_folder = QFileDialog.getExistingDirectory(self, "Sélectionner le dossier de destination")
            if destination_folder:
                self.destination_label.setText(f'Dossier de destination sélectionné: {destination_folder}')

                self.download_thread = DownloadThread(df, destination_folder)
                self.download_thread.progress_update.connect(self.update_progress)
                self.download_thread.image_downloaded.connect(self.image_downloaded)
                self.download_thread.download_finished.connect(self.download_finished)
                self.download_thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def image_downloaded(self, nom, prenoms, image_data, filepath):
        pixmap = QPixmap()
        pixmap.loadFromData(image_data)
        pixmap = pixmap.scaledToWidth(250)
        self.image_label.setPixmap(pixmap)
        self.output_box.append(f"Image téléchargée : {nom} {prenoms}")

        # Afficher le nom de fichier renommé
        self.output_box.append(f"Image renommée : {os.path.basename(filepath)}")

    def download_finished(self):
        self.output_box.append("Téléchargement terminé.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    downloader_app = ImageDownloaderApp()
    downloader_app.show()
    sys.exit(app.exec_())
