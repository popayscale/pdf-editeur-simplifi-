import sys
import fitz  # PyMuPDF
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QFileDialog, QScrollArea, QCheckBox, QMessageBox
from PyQt5.QtCore import Qt, QMimeData, QTimer
from PyQt5.QtGui import QDrag, QPixmap, QImage, QCursor
from PyPDF2 import PdfReader, PdfWriter

class PDFPage(QWidget):
    def __init__(self, page_number, pixmap, pdf_path, is_copy=False, parent=None):
        super().__init__(parent)
        self.page_number = page_number
        self.pdf_path = pdf_path
        self.pixmap = pixmap
        self.is_copy = is_copy
        self.original_column = None  # Ajout d'une référence à la colonne d'origine

        layout = QVBoxLayout(self)
        self.checkbox = QCheckBox()
        layout.addWidget(self.checkbox)

        self.label = QLabel()
        self.label.setPixmap(pixmap)
        self.label.setFixedSize(200, 280)  # Adjust size as needed
        self.label.setScaledContents(True)
        self.label.setFrameStyle(QLabel.Panel | QLabel.Raised)
        layout.addWidget(self.label)

        self.setFixedSize(220, 330)  # Adjust size to accommodate checkbox

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.drag_start_position = e.pos()

    def mouseMoveEvent(self, e):
        if not (e.buttons() & Qt.LeftButton):
            return
        if (e.pos() - self.drag_start_position).manhattanLength() < QApplication.startDragDistance():
            return

        drag = QDrag(self)
        mime_data = QMimeData()
        mime_data.setText(f"{self.pdf_path}|{self.page_number}|{int(self.is_copy)}|{id(self.original_column)}")
        drag.setMimeData(mime_data)
        drag.setPixmap(self.pixmap.scaled(100, 140, Qt.KeepAspectRatio, Qt.SmoothTransformation))

        column = self.parent().parent().parent()
        if isinstance(column, PDFColumn):
            column.dragEnterEvent(e)

        drag.exec_(Qt.MoveAction)

        if isinstance(column, PDFColumn):
            column.dragLeaveEvent(e)

class PDFColumn(QWidget):
    def __init__(self, pdf_path, is_destination=False, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.is_destination = is_destination
        self.pages = []
        self.is_scrolling = False
        self.scroll_speed = 0
        self.scroll_acceleration = 0.5
        self.max_scroll_speed = 20

        layout = QVBoxLayout(self)

        self.scroll_up_button = QPushButton("▲")
        self.scroll_up_button.setStyleSheet("background-color: lightblue; font-size: 20px;")
        layout.addWidget(self.scroll_up_button)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        layout.addWidget(self.scroll_area)

        content_widget = QWidget()
        self.page_layout = QVBoxLayout(content_widget)
        self.scroll_area.setWidget(content_widget)

        self.scroll_down_button = QPushButton("▼")
        self.scroll_down_button.setStyleSheet("background-color: lightblue; font-size: 20px;")
        layout.addWidget(self.scroll_down_button)

        self.load_pdf_pages(pdf_path)

        button_layout = QHBoxLayout()
        save_button = QPushButton("Enregistrer cette colonne")
        save_button.setStyleSheet("background-color: green; color: white;")
        save_button.clicked.connect(self.save_column)
        button_layout.addWidget(save_button)

        delete_button = QPushButton("Supprimer la sélection")
        delete_button.setStyleSheet("background-color: red; color: white;")
        delete_button.clicked.connect(self.delete_selected_pages)
        button_layout.addWidget(delete_button)

        select_all_button = QPushButton("Sélectionner tout")
        select_all_button.setStyleSheet("background-color: yellow; color: black;")
        select_all_button.clicked.connect(self.select_all_pages)
        button_layout.addWidget(select_all_button)

        deselect_all_button = QPushButton("Désélectionner tout")
        deselect_all_button.setStyleSheet("background-color: orange; color: white;")
        deselect_all_button.clicked.connect(self.deselect_all_pages)
        button_layout.addWidget(deselect_all_button)

        layout.addLayout(button_layout)

        self.setAcceptDrops(True)

        self.scroll_timer = QTimer(self)
        self.scroll_timer.timeout.connect(self.auto_scroll)
        self.scroll_timer.setInterval(16)  # ~60 FPS

    def dragEnterEvent(self, event):
        event.accept()
        self.is_scrolling = True
        self.scroll_timer.start()

    def dragLeaveEvent(self, event):
        self.is_scrolling = False
        self.scroll_timer.stop()
        self.scroll_speed = 0

    def dragMoveEvent(self, event):
        cursor_pos = event.pos()
        scroll_area_rect = self.scroll_area.geometry()
        
        if cursor_pos.y() < scroll_area_rect.top() + 50:
            self.scroll_speed = max(-self.max_scroll_speed, self.scroll_speed - self.scroll_acceleration)
        elif cursor_pos.y() > scroll_area_rect.bottom() - 50:
            self.scroll_speed = min(self.max_scroll_speed, self.scroll_speed + self.scroll_acceleration)
        else:
            self.scroll_speed = 0

    def dropEvent(self, event):
        self.is_scrolling = False
        self.scroll_timer.stop()
        self.scroll_speed = 0

        pos = event.pos()
        mime_data = event.mimeData()

        if mime_data.hasText():
            pdf_path, page_number, is_copy, original_column_id = mime_data.text().split('|')
            page_number = int(page_number)
            is_copy = bool(int(is_copy))
            original_column_id = int(original_column_id)

            scroll_pos = self.scroll_area.mapFrom(self, pos)
            content_pos = self.scroll_area.widget().mapFrom(self.scroll_area, scroll_pos)

            if original_column_id == id(self):
                self.move_page(page_number, content_pos)
            else:
                self.copy_page(pdf_path, page_number, content_pos, is_copy)

    def auto_scroll(self):
        if self.is_scrolling:
            scrollbar = self.scroll_area.verticalScrollBar()
            scrollbar.setValue(int(scrollbar.value() + self.scroll_speed))

    def load_pdf_pages(self, pdf_path):
        pdf_document = fitz.open(pdf_path)
        for i in range(len(pdf_document)):
            page = pdf_document.load_page(i)
            pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
            qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
            pixmap = QPixmap.fromImage(qimg)
            pdf_page = PDFPage(i + 1, pixmap, pdf_path)
            pdf_page.original_column = self
            self.page_layout.addWidget(pdf_page)
            self.pages.append(pdf_page)

    def move_page(self, page_number, pos):
        moving_page = None
        for i, page in enumerate(self.pages):
            if page.page_number == page_number:
                moving_page = page
                self.page_layout.removeWidget(page)
                self.pages.remove(page)
                break

        if moving_page:
            insert_index = self.get_insert_index(pos)
            self.page_layout.insertWidget(insert_index, moving_page)
            self.pages.insert(insert_index, moving_page)

    def copy_page(self, pdf_path, page_number, pos, is_copy):
        pdf_document = fitz.open(pdf_path)
        page = pdf_document.load_page(page_number - 1)
        pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
        qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(qimg)
        new_page = PDFPage(page_number, pixmap, pdf_path, is_copy=is_copy)
        new_page.original_column = self  # Définir la colonne d'origine

        insert_index = self.get_insert_index(pos)
        self.page_layout.insertWidget(insert_index, new_page)
        self.pages.insert(insert_index, new_page)

    def get_insert_index(self, pos):
        for i, page in enumerate(self.pages):
            if page.geometry().contains(pos):
                return i
        return len(self.pages)

    def save_column(self):
        output_path, _ = QFileDialog.getSaveFileName(self, "Enregistrer le PDF de cette colonne", "", "PDF Files (*.pdf)")
        if output_path:
            writer = PdfWriter()
            for page in self.pages:
                reader = PdfReader(page.pdf_path)
                writer.add_page(reader.pages[page.page_number - 1])

            with open(output_path, "wb") as output_file:
                writer.write(output_file)

    def delete_selected_pages(self):
        pages_to_remove = [page for page in self.pages if page.checkbox.isChecked()]
        for page in pages_to_remove:
            self.page_layout.removeWidget(page)
            self.pages.remove(page)
            page.deleteLater()

    def select_all_pages(self):
        for page in self.pages:
            page.checkbox.setChecked(True)

    def deselect_all_pages(self):
        for page in self.pages:
            page.checkbox.setChecked(False)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Visual PDF Merger")
        self.pdf_columns = []

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QHBoxLayout()
        central_widget.setLayout(main_layout)

        control_layout = QVBoxLayout()
        main_layout.addLayout(control_layout)

        load_pdf_button = QPushButton("Charger PDF")
        load_pdf_button.setStyleSheet("background-color: blue; color: white;")
        load_pdf_button.clicked.connect(self.load_pdf)
        control_layout.addWidget(load_pdf_button)

        unload_pdf_button = QPushButton("Décharger PDF")
        unload_pdf_button.setStyleSheet("background-color: purple; color: white;")
        unload_pdf_button.clicked.connect(self.unload_pdf)
        control_layout.addWidget(unload_pdf_button)

        self.save_all_button = QPushButton("Enregistrer PDF Fusionné Global")
        self.save_all_button.setStyleSheet("background-color: green; color: white;")
        self.save_all_button.clicked.connect(self.save_merged_pdf)
        self.save_all_button.setEnabled(False)  # Disable by default
        control_layout.addWidget(self.save_all_button)

        self.pdf_layout = QHBoxLayout()
        main_layout.addLayout(self.pdf_layout)

    def load_pdf(self):
        file_dialog = QFileDialog()
        pdf_paths, _ = file_dialog.getOpenFileNames(self, "Sélectionner un ou plusieurs PDF", "", "PDF Files (*.pdf)")

        for pdf_path in pdf_paths:
            pdf_column = PDFColumn(pdf_path)
            self.pdf_columns.append(pdf_column)
            self.pdf_layout.addWidget(pdf_column)

        # Enable the save all button if at least 2 PDFs are loaded
        if len(self.pdf_columns) >= 2:
            self.save_all_button.setEnabled(True)

    def unload_pdf(self):
        if not self.pdf_columns:
            QMessageBox.warning(self, "Aucun PDF chargé", "Aucun PDF n'est actuellement chargé.")
            return

        selected_columns = []
        for column in self.pdf_columns:
            if any(page.checkbox.isChecked() for page in column.pages):
                selected_columns.append(column)

        if not selected_columns:
            QMessageBox.warning(self, "Aucune sélection", "Aucune page n'est sélectionnée pour être déchargée.")
            return

        for column in selected_columns:
            self.pdf_columns.remove(column)
            self.pdf_layout.removeWidget(column)
            column.deleteLater()

        # Disable the save all button if less than 2 PDFs are loaded
        if len(self.pdf_columns) < 2:
            self.save_all_button.setEnabled(False)

    def save_merged_pdf(self):
        output_path, _ = QFileDialog.getSaveFileName(self, "Enregistrer le PDF fusionné global", "", "PDF Files (*.pdf)")
        if output_path:
            writer = PdfWriter()
            for column in self.pdf_columns:
                for page in column.pages:
                    reader = PdfReader(page.pdf_path)
                    writer.add_page(reader.pages[page.page_number - 1])

            with open(output_path, "wb") as output_file:
                writer.write(output_file)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
