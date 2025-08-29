import sys
import pandas as pd
import os
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QPixmap, QPainter, QColor, QFont
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QPushButton,
    QFileDialog,
    QListWidget,
    QLabel,
    QComboBox,
    QProgressBar,
    QHBoxLayout,
    QScrollArea,
)


class MergeFilesThread(QThread):
    progress = pyqtSignal(int)
    result = pyqtSignal(str)

    def __init__(self, files, mode, output_filename):
        super().__init__()
        self.files = files
        self.mode = mode
        self.output_filename = output_filename

    def run(self):
        try:
            df_list = []
            for i, file in enumerate(self.files):
                df_list.append(pd.read_excel(file))
                self.progress.emit(int((i + 1) / len(self.files) * 100))

            download_path = os.path.join(os.path.expanduser("~"), "Downloads")
            output_file = os.path.join(download_path, self.output_filename)
            if self.mode == "Workbook":
                merged_df = pd.concat(df_list, ignore_index=True)
                merged_df.to_excel(output_file, index=False)
            elif self.mode == "Worksheet":
                with pd.ExcelWriter(output_file) as writer:
                    for i, df in enumerate(df_list):
                        df.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)
            self.result.emit(
                f"Excel files merged successfully. Merged data saved to {output_file}"
            )
        except Exception as e:
            self.result.emit(f"Error merging files: {e}")


class MultiFileOpener(QWidget):
    def __init__(self, mode):
        super().__init__()
        self.mode = mode
        self.file_paths = []
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.Label = QLabel("EXCEL FILE MERGER", self)
        font = QFont("Arial", 14, QFont.Bold)
        self.Label.setAlignment(Qt.AlignHCenter)  # type: ignore
        self.Label.setFont(font)
        layout.addWidget(self.Label)

        self.listWidget = QListWidget(self)
        layout.addWidget(self.listWidget)

        button_layout = QHBoxLayout()
        self.moveUpButton = QPushButton("Move Up", self)
        self.moveUpButton.setGeometry(90,230,71,17)
        self.moveUpButton.clicked.connect(self.moveUp)
        button_layout.addWidget(self.moveUpButton)

        self.moveDownButton = QPushButton("Move Down", self)
        self.moveDownButton.clicked.connect(self.moveDown)
        button_layout.addWidget(self.moveDownButton)

        self.deleteButton = QPushButton("Delete", self)
        self.deleteButton.clicked.connect(self.deleteItem)
        button_layout.addWidget(self.deleteButton)

        layout.addLayout(button_layout)

        self.button = QPushButton("Select Files", self)
        self.button.clicked.connect(self.openFiles)
        layout.addWidget(self.button)

        saveus_layout = QHBoxLayout()
        self.fix_window_button = QPushButton("Merge into")
        self.fix_window_button.clicked.connect(self.fix_window)
        self.selection_combo = QComboBox()
        self.selection_combo.addItem(" ")  # deselect option
        self.selection_combo.addItem("Workbook")
        self.selection_combo.addItem("Worksheet")
        saveus_layout.addWidget(self.fix_window_button)
        saveus_layout.addWidget(self.selection_combo)

        layout.addLayout(saveus_layout)

        Merger_layout = QHBoxLayout()
        self.window_button = QPushButton("Merge", self)
        self.window_button.clicked.connect(self.mergeFiles)

        self.progressBar = QProgressBar(self)
        self.progressBar.setAlignment(Qt.AlignCenter)  # type: ignore
        Merger_layout.addWidget(self.window_button)
        Merger_layout.addWidget(self.progressBar)

        layout.addLayout(Merger_layout)

        self.statusLabel = QLabel(" ", self)
        font1 = QFont("Arial", 10)
        self.statusLabel.setFont(font1)
        layout.addWidget(self.statusLabel)

        self.setLayout(layout)

    def openFiles(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly

        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Multiple Files",
            "",
            "All Files (*);;Excel Files (*.xlsx);;Python Files (*.py)",
            options=options,
        )
        if files:
            for file in files:
                self.file_paths.append(file)
                self.listWidget.addItem(os.path.basename(file))

    def fix_window(self):
        self.mode = self.selection_combo.currentText()

    def moveUp(self):
        currentRow = self.listWidget.currentRow()
        if currentRow > 0:
            currentItem = self.listWidget.takeItem(currentRow)
            self.listWidget.insertItem(currentRow - 1, currentItem)
            self.listWidget.setCurrentRow(currentRow - 1)
            self.file_paths.insert(currentRow - 1, self.file_paths.pop(currentRow))

    def moveDown(self):
        currentRow = self.listWidget.currentRow()
        if currentRow < self.listWidget.count() - 1:
            currentItem = self.listWidget.takeItem(currentRow)
            self.listWidget.insertItem(currentRow + 1, currentItem)
            self.listWidget.setCurrentRow(currentRow + 1)
            self.file_paths.insert(currentRow + 1, self.file_paths.pop(currentRow))

    def deleteItem(self):
        currentRow = self.listWidget.currentRow()
        if currentRow >= 0:
            self.listWidget.takeItem(currentRow)
            del self.file_paths[currentRow]

    def mergeFiles(self):
        self.fix_window()

        if not self.file_paths:
            self.statusLabel.setText("No files selected!")
            return

        output_filename = self.generate_output_filename()

        self.merge_thread = MergeFilesThread(
            self.file_paths, self.mode, output_filename
        )
        self.merge_thread.progress.connect(self.updateProgress)
        self.merge_thread.result.connect(self.showResult)
        self.merge_thread.start()

    def generate_output_filename(self):
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        base_names = [
            os.path.splitext(os.path.basename(path))[0] for path in self.file_paths
        ]

        # Join base names and ensure it is within 200 characters including timestamp and extension
        joined_base_names = "_".join(base_names)
        max_base_name_length = (
            80 - len(timestamp) - len(".xlsx") - 1
        )  # 1 for the underscore before the timestamp

        if len(joined_base_names) > max_base_name_length:
            joined_base_names = joined_base_names[:max_base_name_length]

        filename = f"{joined_base_names}_{timestamp}.xlsx"
        return filename

    def updateProgress(self, value):
        self.progressBar.setValue(value)

    def showResult(self, message):
        self.statusLabel.setText(message)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        main_widget = QWidget()
        self.multi_file_opener = MultiFileOpener("")

        scroll_layout = QVBoxLayout(main_widget)
        scroll_layout.addWidget(self.multi_file_opener)
        scroll_area.setWidget(main_widget)

        layout = QVBoxLayout()
        layout.addWidget(scroll_area)

        self.setWindowTitle(" Excel File Merger")  # linear-gradient(rgb(140,82,255), rgb(0,191,99));
        self.setGeometry(0, 0, 550, 600)
        self.adjustSize()

        self.setLayout(layout)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    main_window = MainWindow()
    main_window.show()

    sys.exit(app.exec_())
