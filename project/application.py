
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit,
                             QPushButton, QTableWidget, QTableWidgetItem, QFileDialog, QInputDialog)
from PyQt5.QtCore import Qt

from project.utils import *
from project.process import *



class App(QWidget):
    def __init__(self):
        super().__init__()
        self.title = 'File Processor'
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(100, 100, 800, 600)

        self.layout = QVBoxLayout()

        self.label = QLabel('Enter file name:', self)
        self.layout.addWidget(self.label)

        self.file_name_input = QLineEdit(self)
        self.layout.addWidget(self.file_name_input)

        self.browse_button = QPushButton('Browse', self)
        self.browse_button.clicked.connect(self.open_file_dialog)
        self.layout.addWidget(self.browse_button)

        self.process_button = QPushButton('Process Data', self)
        self.process_button.clicked.connect(self.process_file)
        self.layout.addWidget(self.process_button)


        self.process_button = QPushButton('Show Summary', self)
        self.process_button.clicked.connect(self.show_summary)
        self.layout.addWidget(self.process_button)


        self.save_button = QPushButton('Save to Exel', self)
        self.save_button.clicked.connect(self.save_exel)

        self.save_button.setEnabled(False)  # Disabled until a file is processed
        self.layout.addWidget(self.save_button)

        self.table = QTableWidget(self)
        self.layout.addWidget(self.table)

        self.setLayout(self.layout)
        self.df = None  # To store the processed DataFrame

    def open_file_dialog(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Open Text File", "", "Text Files (*.txt);;All Files (*)", options=options)
        if file_name:
            self.file_name_input.setText(file_name)

    def process_file(self):
        file_name = self.file_name_input.text()
        if file_name:
            try:
                self.df, self.summary = process_file(file_name)
                self.display_output(self.df)
                self.save_button.setEnabled(True)  # Enable the save button after processing
            except Exception as e:
                self.display_error(f"Error processing file: {str(e)}")

    def show_summary(self):
        # print(self.summary)
        if self.summary is not None:
            try:
                # self.new_df = work_summary(self.df)
                self.display_output(self.summary)
                self.save_button.setEnabled(True)  # Enable the save button after processing
            except Exception as e:
                self.display_error(f"Error processing file: {str(e)}")


    def display_output(self, df):
        self.table.setColumnCount(len(df.columns))
        self.table.setRowCount(len(df))
        self.table.setHorizontalHeaderLabels(df.columns)

        for i in range(len(df)):
            for j in range(len(df.columns)):
                item = QTableWidgetItem(str(df.iat[i, j]))
                item.setTextAlignment(Qt.AlignCenter)  # Center-align the text
                self.table.setItem(i, j, item)

                
    def save_exel(self):
        options = ['summary', 'data']
        text, ok = QInputDialog.getItem(self, "Save File", "Choose what to save:", options, 0, False)

        # text, ok = QInputDialog.getText(self, 'Save File', 'Enter summary or data:')
        if text == 'data':
            if self.df is not None:
                options = QFileDialog.Options()
                save_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
                if save_path:
                    if not save_path.endswith('.xlsx'):
                        save_path += '.xlsx'
                    self.df.to_excel(save_path, index=False)
                    modify_exel(save_path)
                    self.display_message(f"File saved successfully at {save_path}")
        elif text == 'summary':
            if self.summary is not None:
                options = QFileDialog.Options()
                save_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel Files (*.xlsx);;All Files (*)", options=options)
                if save_path:
                    if not save_path.endswith('.xlsx'):
                        save_path += '.xlsx'
                    self.summary.to_excel(save_path, index=False)
                    modify_exel(save_path)
                    self.display_message(f"File saved successfully at {save_path}")

    def display_error(self, message):
        self.table.clear()
        self.table.setRowCount(1)
        self.table.setColumnCount(1)
        self.table.setItem(0, 0, QTableWidgetItem(message))

    def display_message(self, message):
        self.table.clear()
        self.table.setRowCount(1)
        self.table.setColumnCount(1)
        self.table.setItem(0, 0, QTableWidgetItem(message))
