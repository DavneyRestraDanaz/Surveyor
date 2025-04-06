import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout, 
    QTableWidget, QTableWidgetItem, QLabel, QHBoxLayout, QLineEdit,
    QGridLayout, QGroupBox, QFormLayout, QHeaderView, QDialog, 
    QRadioButton, QDialogButtonBox, QCalendarWidget, QMessageBox,
    QComboBox, QScrollArea
)
from PyQt5.QtGui import (
    QFont, QTextDocument, QPageSize, QPageLayout
)
from PyQt5.QtCore import (
    Qt, QEventLoop, QSizeF, QMarginsF
)
from PyQt5.QtPrintSupport import QPrinter, QPrinterInfo, QPrintDialog, QPrintPreviewDialog
from PyQt5.QtWebEngineWidgets import QWebEngineView
import openpyxl
from openpyxl.writer.excel import ExcelWriter
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import os
import shutil
import copy
from openpyxl.utils.cell import get_column_letter
import re

class ExcelViewerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Aplikasi Laporan Psikologi")
        # Set window size to nearly fullscreen (90% of screen size)
        screen = QApplication.primaryScreen()
        screen_size = screen.availableGeometry()
        width = int(screen_size.width() * 0.9)
        height = int(screen_size.height() * 0.9)
        self.setGeometry(
            (screen_size.width() - width) // 2,  # Center horizontally
            (screen_size.height() - height) // 2, # Center vertically 
            width,
            height
        )
        # Updated columns based on the images provided
        self.columns = [
            "No", "No Tes", "TGL Lahir", "JK", "Nama Peserta", 
            "IQ", "Konkrit Praktis", "Verbal", "Flexibilitas Pikir", 
            "Daya Abstraksi Verbal", "Berpikir Praktis", "Berpikir Teoritis", 
            "Memori", "WA GE", "RA ZR", "KLASIFIKASI",
            "N", "G", "A", "L", "P", "I", "T", "V", "S", "B", "O", "X", 
            "C (Coding)", "D", "R", "Z", "E", "K", "F", "W", "CD", "TV", "BO", "SO", "BX"
        ]
        self.input_columns = self.columns.copy()
        self.initUI()
        self.excel_file_path = ""
        self.df = pd.DataFrame(columns=self.columns)

    def initUI(self):
        # Create scroll area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        # Create container widget for scroll area
        container = QWidget()
        
        # Create main layout
        main_layout = QVBoxLayout(container)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Set container as scroll area widget
        scroll.setWidget(container)
        
        # Create layout for the main window
        window_layout = QVBoxLayout(self)
        window_layout.addWidget(scroll)

        # File Selection Section
        file_group = QGroupBox("File Excel")
        file_group.setFont(QFont("Arial", 11, QFont.Bold))
        file_layout = QVBoxLayout()
        
        self.label = QLabel("Pilih file Excel (.xlsx)")
        self.label.setFont(QFont("Arial", 10))
        file_layout.addWidget(self.label)

        self.btn_select = QPushButton("Pilih File Excel") 
        self.btn_select.setFont(QFont("Arial", 10))
        self.btn_select.setFixedHeight(35)
        self.btn_select.clicked.connect(self.load_excel)
        # Remove automatic disable - now handled in load_excel() based on file selection
        file_layout.addWidget(self.btn_select)
        
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        # Group for Personal Information
        personal_group = QGroupBox("Personal Information")
        personal_group.setFont(QFont("Arial", 11, QFont.Bold))
        personal_layout = QGridLayout()  # Changed to grid layout for better organization
        
        personal_fields = ["No", "No Tes", "TGL Lahir", "JK", "Nama Peserta"]
        self.personal_inputs = []
        for i, placeholder in enumerate(personal_fields):
            label = QLabel(placeholder + ":")
            label.setFont(QFont("Arial", 10))
            
            # Special handling for JK and TGL Lahir
            if placeholder == "TGL Lahir":
                field = QPushButton("Pilih Tanggal")
                field.clicked.connect(self.show_calendar)
            elif placeholder == "JK":
                field = QPushButton("Pilih Jenis Kelamin")
                field.clicked.connect(self.show_gender_dialog)
            else:
                field = QLineEdit()
                
            field.setFixedHeight(30)
            self.personal_inputs.append(field)
            row = i // 2  # Arrange fields in rows of 2
            col = (i % 2) * 2
            personal_layout.addWidget(label, row, col)
            personal_layout.addWidget(field, row, col + 1)
        
        personal_group.setLayout(personal_layout)
        main_layout.addWidget(personal_group)

        # Group for IST
        ist_group = QGroupBox("IST")
        ist_group.setFont(QFont("Arial", 11, QFont.Bold))
        ist_layout = QGridLayout()  # Changed to grid layout
        
        # Remove WA GE, RA ZR, and IQ KLASIFIKASI from input fields
        ist_fields = ["IQ", "Konkrit Praktis", "Verbal", "Flexibilitas Pikir", 
                      "Daya Abstraksi Verbal", "Berpikir Praktis", "Berpikir Teoritis", 
                      "Memori"]
        self.ist_inputs = []
        for i, placeholder in enumerate(ist_fields):
            label = QLabel(placeholder + ":")
            label.setFont(QFont("Arial", 10))
            field = QLineEdit()
            field.setFixedHeight(30)
            self.ist_inputs.append(field)
            row = i // 3  # Arrange fields in rows of 3
            col = (i % 3) * 2
            ist_layout.addWidget(label, row, col)
            ist_layout.addWidget(field, row, col + 1)
        
        ist_group.setLayout(ist_layout)
        main_layout.addWidget(ist_group)

        # Group for PAPIKOSTICK
        papikostick_group = QGroupBox("PAPIKOSTICK (Numeric)")
        papikostick_group.setFont(QFont("Arial", 11, QFont.Bold))
        papikostick_layout = QGridLayout()  # Changed to grid layout
        
        # Include 'C' in input fields, remove 'C (Coding)', 'CD', 'TV', 'BO', 'SO', 'BX'
        papikostick_fields = ["N", "G", "A", "L", "P", "I", "T", "V", "S", "B", "O", "X", 
                              "C", "D", "R", "Z", "E", "K", "F", "W"]
        self.papikostick_inputs = []
        for i, placeholder in enumerate(papikostick_fields):
            label = QLabel(placeholder + ":")
            label.setFont(QFont("Arial", 10))
            field = QLineEdit()
            field.setFixedHeight(30)
            self.papikostick_inputs.append(field)
            row = i // 5  # Arrange fields in rows of 5
            col = (i % 5) * 2
            papikostick_layout.addWidget(label, row, col)
            papikostick_layout.addWidget(field, row, col + 1)
        
        papikostick_group.setLayout(papikostick_layout)
        main_layout.addWidget(papikostick_group)

        # Add buttons for adding and updating data
        self.btn_add = QPushButton("Tambah Data")
        self.btn_add.setFont(QFont("Arial", 10))
        self.btn_add.setFixedHeight(35)
        self.btn_add.setFixedWidth(200)
        self.btn_add.clicked.connect(lambda: self.add_or_update_row("add"))
        self.btn_add.setEnabled(False)  # Initially disabled

        self.btn_edit = QPushButton("Edit Data")
        self.btn_edit.setFont(QFont("Arial", 10))
        self.btn_edit.setFixedHeight(35)
        self.btn_edit.setFixedWidth(200)
        self.btn_edit.clicked.connect(lambda: self.add_or_update_row("edit"))
        self.btn_edit.setEnabled(False)  # Initially disabled

        # Right-align the buttons in the layout
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.addStretch()  # Push buttons to right
        button_layout.addWidget(self.btn_add)
        button_layout.addWidget(self.btn_edit)
        button_layout.setContentsMargins(0, 0, 20, 0)  # Add right margin
        main_layout.addWidget(button_container)

        # Enable buttons when file is loaded
        self.btn_select.clicked.connect(lambda: self.btn_add.setEnabled(True))
        self.btn_select.clicked.connect(lambda: self.btn_edit.setEnabled(True))

        # Add toggle buttons for each group
        toggle_layout = QHBoxLayout()
        toggle_personal = QPushButton("Personal Information")
        toggle_personal.setCheckable(True)
        toggle_personal.setChecked(False)  # Initially unchecked
        toggle_personal.setFont(QFont("Arial", 10))
        toggle_personal.setFixedHeight(35)
        toggle_personal.toggled.connect(lambda checked: personal_group.setVisible(checked))
        toggle_layout.addWidget(toggle_personal)
        
        toggle_ist = QPushButton("IST")
        toggle_ist.setCheckable(True)
        toggle_ist.setChecked(False)  # Initially unchecked
        toggle_ist.setFont(QFont("Arial", 10))
        toggle_ist.setFixedHeight(35)
        toggle_ist.toggled.connect(lambda checked: ist_group.setVisible(checked))
        toggle_layout.addWidget(toggle_ist)
        
        toggle_papikostick = QPushButton("PAPIKOSTICK")
        toggle_papikostick.setCheckable(True)
        toggle_papikostick.setChecked(False)  # Initially unchecked
        toggle_papikostick.setFont(QFont("Arial", 10))
        toggle_papikostick.setFixedHeight(35)
        toggle_papikostick.toggled.connect(lambda checked: papikostick_group.setVisible(checked))
        toggle_layout.addWidget(toggle_papikostick)
        
        # Initially hide all groups
        personal_group.setVisible(False)
        ist_group.setVisible(False)
        papikostick_group.setVisible(False)
        
        main_layout.addLayout(toggle_layout)

        # Search Section
        search_group = QGroupBox("Pencarian")
        search_group.setFont(QFont("Arial", 11, QFont.Bold))
        search_layout = QHBoxLayout()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Cari data...")
        self.search_input.setFont(QFont("Arial", 10))
        self.search_input.setFixedHeight(30)
        self.search_input.textChanged.connect(self.search_table)
        search_layout.addWidget(self.search_input)

        self.search_column = QComboBox()
        self.search_column.setFont(QFont("Arial", 10))
        self.search_column.setFixedHeight(30)
        self.search_column.addItems(["Semua Kolom"] + self.columns)
        search_layout.addWidget(self.search_column)

        search_group.setLayout(search_layout)
        main_layout.addWidget(search_group)

        # Table Section
        table_group = QGroupBox("Data Hasil")
        table_group.setFont(QFont("Arial", 11, QFont.Bold))
        table_layout = QVBoxLayout()
        
        self.table = QTableWidget()
        self.table.setFont(QFont("Arial", 9))
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # Stretch columns to fit
        self.table.itemSelectionChanged.connect(self.populate_fields_from_selection)
        table_layout.addWidget(self.table)

        # Buttons below table
        button_layout = QHBoxLayout()
        
        self.btn_delete = QPushButton("Delete Row")
        self.btn_delete.setFont(QFont("Arial", 10))
        self.btn_delete.setFixedHeight(35)
        self.btn_delete.clicked.connect(self.delete_selected_row)
        button_layout.addWidget(self.btn_delete)

        # Add Preview PDF button
        self.btn_preview_pdf = QPushButton("Preview PDF")
        self.btn_preview_pdf.setFont(QFont("Arial", 10))
        self.btn_preview_pdf.setFixedHeight(35)
        self.btn_preview_pdf.clicked.connect(self.preview_pdf)
        button_layout.addWidget(self.btn_preview_pdf)

        self.btn_save_excel = QPushButton("Save to Excel")
        self.btn_save_excel.setFont(QFont("Arial", 10))
        self.btn_save_excel.setFixedHeight(35)
        self.btn_save_excel.clicked.connect(self.save_to_excel)
        button_layout.addWidget(self.btn_save_excel)
                
        table_layout.addLayout(button_layout)
        table_group.setLayout(table_layout)
        main_layout.addWidget(table_group)

        self.setLayout(main_layout)

    def search_table(self):
        # Check if Excel file has been loaded
        if not hasattr(self, 'excel_file_path') or not self.excel_file_path:
            QMessageBox.warning(self, "Warning", "Please load an Excel file first!")
            return

        search_text = self.search_input.text().lower().strip()
        selected_column = self.search_column.currentText()

        for row in range(self.table.rowCount()):
            row_visible = False

            if selected_column == "Semua Kolom":
                # Search in all columns
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    if item and search_text in item.text().lower():
                        row_visible = True
                        break
            else:
                # Search in selected column
                col_idx = self.columns.index(selected_column)
                item = self.table.item(row, col_idx)
                if item and search_text in item.text().lower():
                    row_visible = True

            self.table.setRowHidden(row, not row_visible)

        # Debugging output
        print(f"Search text: '{search_text}' in column: '{selected_column}'")
        print(f"Total rows: {self.table.rowCount()}")
        for row in range(self.table.rowCount()):
            print(f"Row {row} visible: {not self.table.isRowHidden(row)}")

    def populate_fields_from_selection(self):
        selected_row = self.table.currentRow()
        if selected_row >= 0:
            row_data = {}
            for col, column_name in enumerate(self.columns):
                item = self.table.item(selected_row, col)
                row_data[column_name] = item.text() if item else ""

            # Debug prints
            print(f"Selected Row: {selected_row}")
            for key, value in row_data.items():
                print(f"{key}: {value}")

            # Populate personal inputs
            for i, field in enumerate(self.personal_inputs):
                if i < len(self.columns):
                    field.setText(row_data.get(self.columns[i], ""))

            # Populate IST inputs
            ist_start_idx = len(self.personal_inputs)
            for i, field in enumerate(self.ist_inputs):
                col_idx = ist_start_idx + i
                if col_idx < len(self.columns):
                    field.setText(row_data.get(self.columns[col_idx], ""))

            # Populate PAPIKOSTICK inputs
            papiko_start_idx = ist_start_idx + len(self.ist_inputs) + 3  # +3 for WA GE, RA ZR, IQ KLASIFIKASI
            for i, field in enumerate(self.papikostick_inputs):
                col_idx = papiko_start_idx + i
                if col_idx < len(self.columns):
                    field.setText(row_data.get(self.columns[col_idx], ""))

    def load_excel(self):
        
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Pilih File Excel", "", "Excel Files (*.xlsx);;All Files (*)", options=options)

        # Enable the button if user cancels file selection
        if not file_path:
            self.btn_select.setEnabled(True)
            return

        try:
            self.label.setText(f"File: {file_path}")
            self.excel_file_path = file_path
            self.btn_select.setEnabled(False)
            
            # **🔹 Membaca seluruh sheet dalam file Excel**
            self.excel_data = pd.read_excel(file_path, sheet_name=None)  # Baca semua sheet
            
            # **🔹 Pastikan Sheet1 dan Sheet2 terbaca**
            if "Sheet1" in self.excel_data:
                self.sheet1_data = self.excel_data["Sheet1"]
            else:
                print("Sheet1 tidak ditemukan!")
                self.btn_select.setEnabled(True)
                return

            if "Sheet2" in self.excel_data:
                self.sheet2_data = self.excel_data["Sheet2"]
            else:
                print("Sheet2 tidak ditemukan!")
                self.btn_select.setEnabled(True)
                return

            # **🔹 Proses data setelah membaca**
            self.process_excel(file_path)
            
        except Exception as e:
            print(f"Error loading file: {e}")
            self.btn_select.setEnabled(True)
            QMessageBox.critical(self, "Error", "Failed to load Excel file. Please try again.")
    def process_excel(self, file_path):
        try:
            # 🔹 Baca semua sheet dalam file Excel
            sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')

            # 🔹 Pastikan Sheet1 dan Sheet2 ada
            if "Sheet1" in sheets:
                df_sheet1 = sheets["Sheet1"]
            else:
                print("Sheet1 tidak ditemukan!")
                df_sheet1 = None
                self.btn_select.setEnabled(True)
                return

            if "Sheet2" in sheets:
                df_sheet2 = sheets["Sheet2"]
            else:
                print("Sheet2 tidak ditemukan!")
                df_sheet2 = None
                self.btn_select.setEnabled(True)
                return

            # 🔹 Proses Sheet1 (jika ada)
            if df_sheet1 is not None:
                print("Original columns (Sheet1):", df_sheet1.columns.tolist())

                # Mencari baris awal data berdasarkan keyword "No"
                start_row = None
                for idx, row in df_sheet1.iterrows():
                    if any(str(cell).strip().lower() == 'no' for cell in row):
                        start_row = idx
                        break

                if start_row is not None:
                    df_sheet1 = pd.read_excel(file_path, sheet_name="Sheet1", engine='openpyxl', skiprows=start_row+1)
                    new_df = df_sheet1.copy()

                    # Konversi kolom tertentu menjadi string
                    str_columns = ['No', 'No Tes', 'TGL Lahir', 'JK', 'Nama Peserta']
                    for col in str_columns:
                        if col in new_df.columns:
                            new_df[col] = new_df[col].astype(str)

                    # Konversi kolom angka ke numeric
                    numeric_columns = ['IQ', 'Konkrit Praktis', 'Verbal', 'Flexibilitas Pikir', 
                                    'Daya Abstraksi Verbal', 'Berpikir Praktis', 'Berpikir Teoritis', 
                                    'Memori']
                    for col in numeric_columns:
                        if col in new_df.columns:
                            new_df[col] = pd.to_numeric(new_df[col], errors='coerce').fillna(0)
                        else:
                            new_df[col] = 0

                    # Konversi angka ke string untuk tampilan
                    for col in numeric_columns:
                        if col in new_df.columns:
                            new_df[col] = new_df[col].astype(str)

                    self.df_sheet1 = new_df.fillna("")
                    self.columns = list(new_df.columns)
                    self.show_table(self.df_sheet1)
                else:
                    print("Could not find the start of data in Sheet1")
                    self.btn_select.setEnabled(True)
                    return

            # 🔹 Proses Sheet2 (jika ada)
            if df_sheet2 is not None:
                print("Original columns (Sheet2):", df_sheet2.columns.tolist())

                # Simpan data Sheet2 tanpa pemrosesan tambahan
                self.df_sheet2 = df_sheet2.fillna("")
        
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            import traceback
            traceback.print_exc()
            self.btn_select.setEnabled(True)
            QMessageBox.critical(self, "Error", "Failed to process Excel file. Please try again.")

    def show_table(self, df):
        try:
            self.table.setRowCount(len(df))
            self.table.setColumnCount(len(df.columns))
            self.table.setHorizontalHeaderLabels(df.columns)

            for row in range(len(df)):
                for col, column_name in enumerate(df.columns):
                    value = str(df.iloc[row][column_name])
                    self.table.setItem(row, col, QTableWidgetItem(value))

            # Change from Stretch to ResizeToContents for better visibility of column headers
            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        except Exception as e:
            print(f"Error showing table: {e}")
            import traceback
            traceback.print_exc()

    def show_calendar(self):
        calendar = QCalendarWidget(self)
        calendar.clicked.connect(lambda date: self.set_date(date))
        calendar.setWindowFlags(Qt.Popup)
        # Use personal_inputs instead of input_fields
        pos = self.personal_inputs[2].mapToGlobal(self.personal_inputs[2].rect().bottomLeft())
        calendar.move(pos)
        calendar.show()

    def set_date(self, date):
        # Update this method to use personal_inputs as well
        self.personal_inputs[2].setText(date.toString("dd/MM/yyyy"))

    def show_gender_dialog(self):
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QRadioButton, QDialogButtonBox
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Pilih Jenis Kelamin")
        layout = QVBoxLayout()
        
        radio_l = QRadioButton("L (Laki-laki)")
        radio_p = QRadioButton("P (Perempuan)")
        
        layout.addWidget(radio_l)
        layout.addWidget(radio_p)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        
        layout.addWidget(buttons)
        dialog.setLayout(layout)
        
        if dialog.exec_() == QDialog.Accepted:
            if radio_l.isChecked():
                self.personal_inputs[3].setText("L")
            elif radio_p.isChecked():
                self.personal_inputs[3].setText("P")

    def add_or_update_row(self, mode="add"):
        # Collect values from all three input groups
        values = []
        values.extend([field.text() for field in self.personal_inputs])
        values.extend([field.text() for field in self.ist_inputs])
        values.extend([field.text() for field in self.papikostick_inputs])
        
        # Debug information
        print(f"Number of values: {len(values)}")
        print(f"Number of columns: {len(self.columns)}")
        print(f"Columns: {self.columns}")  # Print column names for debugging
        
        # Check if any field is empty
        if any(not val.strip() for val in values):
            print("Please fill all fields")
            return
        
        # Ensure the values list has the correct number of elements
        expected_fields = len(self.personal_inputs) + len(self.ist_inputs) + len(self.papikostick_inputs)
        if len(values) < expected_fields:
            print(f"Expected {expected_fields} fields, got {len(values)}")
            return
        
        try:
            # Convert numeric values
            for i in range(5, len(values)):
                if i < len(values) and values[i].strip():  # Check if index is within bounds and value is not empty
                    try:
                        values[i] = float(values[i])
                    except ValueError:
                        print(f"Invalid numeric value: {values[i]} at position {i}")
                        return
        except ValueError as e:
            print(f"Please enter valid numeric values for numeric fields: {e}")
            return
        
        # Create a complete row with all columns
        row_data = {}
        
        # Add personal info
        for i, field in enumerate(self.personal_inputs):
            if i < len(self.columns):
                row_data[self.columns[i]] = field.text()
        
        # Add IST values
        ist_start_idx = len(self.personal_inputs)
        for i, field in enumerate(self.ist_inputs):
            col_idx = ist_start_idx + i
            if col_idx < len(self.columns):
                try:
                    row_data[self.columns[col_idx]] = float(field.text()) if field.text().strip() else 0
                except ValueError:
                    print(f"Invalid numeric value: {field.text()}")
                    return
        
        # Calculate derived values
        # Get the exact column names from self.columns
        wa_ge_col_name = 'Unnamed: 13'
        ra_zr_col_name = 'Unnamed: 14'
        iq_klas_col_name = 'KLASIFIKASI'
        
        # Calculate WA GE
        verbal_val = float(row_data.get("Verbal", 0))
        dav_val = float(row_data.get("Daya Abstraksi Verbal", 0))
        wa_ge = (verbal_val + dav_val) / 2
        row_data[wa_ge_col_name] = wa_ge
        
        # Calculate RA ZR
        bp_val = float(row_data.get("Berpikir Praktis", 0))
        bt_val = float(row_data.get("Berpikir Teoritis", 0))
        ra_zr = (bp_val + bt_val) / 2
        row_data[ra_zr_col_name] = ra_zr
        
        # Calculate IQ classification
        iq = float(row_data.get("IQ", 0))
        if iq < 79:
            iq_klasifikasi = "Rendah"
        elif 79 <= iq < 90:
            iq_klasifikasi = "Dibawah Rata-Rata"
        elif 90 <= iq < 110:
            iq_klasifikasi = "Rata-Rata"
        elif 110 <= iq < 120:
            iq_klasifikasi = "Diatas Rata-Rata"
        else:
            iq_klasifikasi = "Superior"
        
        row_data[iq_klas_col_name] = iq_klasifikasi
        
        # Add PAPIKOSTICK values
        papiko_start_idx = ist_start_idx + len(self.ist_inputs) + 3  # +3 for WA GE, RA ZR, IQ KLASIFIKASI
        data_idx = 0  # Indeks input field (mengabaikan "C (Coding)")

        for col_idx in range(papiko_start_idx, papiko_start_idx + len(self.papikostick_inputs) + 1):
            if col_idx >= len(self.columns):
                break  # Pastikan tidak melebihi jumlah kolom

            col_name = self.columns[col_idx]

            # Jika kolom adalah "C (Coding)", isi otomatis berdasarkan "C"
            if col_name == "C (Coding)":
                c_value = float(row_data.get("C", 0))
                row_data["C (Coding)"] = 10 - c_value if 1 <= c_value <= 9 else 0
                continue  # Lewati iterasi ini agar tidak mengganggu indeks input

            # Pastikan tidak melebihi jumlah input field
            if data_idx < len(self.papikostick_inputs):
                try:
                    row_data[col_name] = float(self.papikostick_inputs[data_idx].text()) if self.papikostick_inputs[data_idx].text().strip() else 0
                except ValueError:
                    print(f"Invalid numeric value: {self.papikostick_inputs[data_idx].text()}")
                    return

            data_idx += 1  # Hanya naikkan indeks input field jika bukan "C (Coding)"

        # Calculate C (Coding) based on the formula
        c_value = float(row_data.get("C", 0))
        c_coding = 0
        if c_value == 1:
            c_coding = 9
        elif c_value == 2:
            c_coding = 8
        elif c_value == 3:
            c_coding = 7
        elif c_value == 4:
            c_coding = 6
        elif c_value == 5:
            c_coding = 5
        elif c_value == 6:
            c_coding = 4
        elif c_value == 7:
            c_coding = 3
        elif c_value == 8:
            c_coding = 2
        elif c_value == 9:
            c_coding = 1
        
        row_data["C (Coding)"] = c_coding

         # Hitung kolom otomatis berdasarkan rumus
        try:
            row_data["CD"] = (row_data.get("C", 0) + row_data.get("D", 0)) / 2
            row_data["TV"] = (row_data.get("T", 0) + row_data.get("V", 0)) / 2
            row_data["BO"] = (row_data.get("B", 0) + row_data.get("O", 0)) / 2
            row_data["SO"] = (row_data.get("S", 0) + row_data.get("O", 0)) / 2
            row_data["BX"] = (row_data.get("B", 0) + row_data.get("X", 0)) / 2
        except TypeError as e:
            print(f"Error calculating derived columns: {e}")
        
        # Hitung Intelegensi Umum berdasarkan nilai IQ
        iq_value = row_data.get("IQ", 0)  # Ambil nilai IQ, default ke 0 jika kosong

        if iq_value < 90:
            row_data["Intelegensi Umum"] = "K"  # Kurang
        elif 90 <= iq_value <= 109:
            row_data["Intelegensi Umum"] = "C"  # Cukup
        else:
            row_data["Intelegensi Umum"] = "B"  # Baik

        # Hitung Daya Analisa/AN berdasarkan nilai Flexibilitas Pikir
        flex_pikir_value = row_data.get("Flexibilitas Pikir", 0)  # Ambil nilai, default ke 0 jika kosong

        if flex_pikir_value < 90:
            row_data["Daya Analisa/ AN"] = "K"  # Kurang
        elif flex_pikir_value < 110:
            row_data["Daya Analisa/ AN"] = "C"  # Cukup
        else:
            row_data["Daya Analisa/ AN"] = "B"  # Baik

        # Hitung Kemampuan Verbal/WA GE berdasarkan nilai Unnamed: 13
        unnamed_13_value = row_data.get("Unnamed: 13", 0)  # Ambil nilai, default ke 0 jika kosong

        if unnamed_13_value < 90:
            row_data["Kemampuan Verbal/WA GE"] = "K"  # Kurang
        elif unnamed_13_value < 110:
            row_data["Kemampuan Verbal/WA GE"] = "C"  # Cukup
        else:
            row_data["Kemampuan Verbal/WA GE"] = "B"  # Baik

        # Hitung Kemampuan Numerik/RA ZR berdasarkan nilai Unnamed: 14
        unnamed_14_value = row_data.get("Unnamed: 14", 0)  # Ambil nilai, default ke 0 jika kosong

        if unnamed_14_value < 90:
            row_data["Kemampuan Numerik/ RA ZR"] = "K"  # Kurang
        elif unnamed_14_value < 110:
            row_data["Kemampuan Numerik/ RA ZR"] = "C"  # Cukup
        else:
            row_data["Kemampuan Numerik/ RA ZR"] = "B"  # Baik
        
        # Perhitungan otomatis berdasarkan rumus yang diberikan

        # Daya Ingat/ME (dari kolom Memori)
        memori_value = row_data.get("Memori", 0)
        if memori_value < 90:
            row_data["Daya Ingat/ME"] = "K"
        elif memori_value < 110:
            row_data["Daya Ingat/ME"] = "C"
        else:
            row_data["Daya Ingat/ME"] = "B"

        # Fleksibilitas/TV (dari kolom TV)
        tv_value = row_data.get("TV", 0)
        if tv_value < 4:
            row_data["Fleksibilitas/ T V"] = "K"
        elif tv_value < 6:
            row_data["Fleksibilitas/ T V"] = "C"
        else:
            row_data["Fleksibilitas/ T V"] = "B"

        # Sistematika Kerja/CD (dari kolom CD)
        cd_value = row_data.get("CD", 0)
        if cd_value < 4:
            row_data["Sistematika Kerja/ cd"] = "K"
        elif cd_value < 6:
            row_data["Sistematika Kerja/ cd"] = "C"
        else:
            row_data["Sistematika Kerja/ cd"] = "B"

        # Inisiatif/W (dari kolom W)
        w_value = row_data.get("W", 0)
        if w_value < 4:
            row_data["Inisiatif/W"] = "B"
        elif w_value < 6:
            row_data["Inisiatif/W"] = "C"
        else:
            row_data["Inisiatif/W"] = "K"

        # Stabilitas Emosi/E (dari kolom E)
        e_value = row_data.get("E", 0)
        if e_value < 4:
            row_data["Stabilitas Emosi / E"] = "B"
        elif e_value < 6:
            row_data["Stabilitas Emosi / E"] = "C"
        else:
            row_data["Stabilitas Emosi / E"] = "K"

        # Komunikasi/BO (dari kolom BO)
        bo_value = row_data.get("BO", 0)
        if bo_value < 4:
            row_data["Komunikasi / B O"] = "K"
        elif bo_value < 6:
            row_data["Komunikasi / B O"] = "C"
        else:
            row_data["Komunikasi / B O"] = "B"

        # Keterampilan Interpersonal/SO (dari kolom SO)
        so_value = row_data.get("SO", 0)
        if so_value < 4:
            row_data["Keterampilan Interpersonal / S O"] = "K"
        elif so_value < 6:
            row_data["Keterampilan Interpersonal / S O"] = "C"
        else:
            row_data["Keterampilan Interpersonal / S O"] = "B"

        # Kerjasama/BX (dari kolom BX)
        bx_value = row_data.get("BX", 0)
        if bx_value < 4:
            row_data["Kerjasama / B X"] = "K"
        elif bx_value < 6:
            row_data["Kerjasama / B X"] = "C"
        else:
            row_data["Kerjasama / B X"] = "B"

        try:
            # Pastikan kolom 'D' ada atau cukup kolom dalam Sheet2
            if "D" in self.sheet2_data.columns or self.sheet2_data.shape[1] > 3:
                column_d = self.sheet2_data["D"] if "D" in self.sheet2_data.columns else self.sheet2_data.iloc[:, 3]

                # Intelegensi Umum
                intelegensi_umum_value = row_data.get("Intelegensi Umum", "")
                if len(self.sheet2_data) > 4:
                    if intelegensi_umum_value == "K":
                        row_data["Intelegensi Umum.1"] = column_d.iloc[3]  # D5
                    elif intelegensi_umum_value == "B":
                        row_data["Intelegensi Umum.1"] = column_d.iloc[1]  # D3
                    elif intelegensi_umum_value == "C":
                        row_data["Intelegensi Umum.1"] = column_d.iloc[2]  # D4

                # Daya Analisa
                daya_analisa_value = row_data.get("Daya Analisa/ AN", "")
                if len(self.sheet2_data) > 7:
                    if daya_analisa_value == "B":
                        row_data["Daya Analisa/ AN.1"] = column_d.iloc[4]  # D6
                    elif daya_analisa_value == "C":
                        row_data["Daya Analisa/ AN.1"] = column_d.iloc[5]  # D7
                    elif daya_analisa_value == "K":
                        row_data["Daya Analisa/ AN.1"] = column_d.iloc[6]  # D8

                # Kemampuan Verbal
                kemampuan_verbal_value = row_data.get("Kemampuan Verbal/WA GE", "")
                if len(self.sheet2_data) > 10:
                    if kemampuan_verbal_value == "B":
                        row_data["Kemampuan Verbal/WA GE.1"] = column_d.iloc[7]  # D9
                    elif kemampuan_verbal_value == "C":
                        row_data["Kemampuan Verbal/WA GE.1"] = column_d.iloc[8]  # D10
                    elif kemampuan_verbal_value == "K":
                        row_data["Kemampuan Verbal/WA GE.1"] = column_d.iloc[9]  # D11

                # Kemampuan Numerik
                kemampuan_numerik_value = row_data.get("Kemampuan Numerik/ RA ZR", "")
                if len(self.sheet2_data) > 13:
                    if kemampuan_numerik_value == "B":
                        row_data["Kemampuan Numerik/ RA ZR.1"] = column_d.iloc[10]  # D12
                    elif kemampuan_numerik_value == "C":
                        row_data["Kemampuan Numerik/ RA ZR.1"] = column_d.iloc[11]  # D13
                    elif kemampuan_numerik_value == "K":
                        row_data["Kemampuan Numerik/ RA ZR.1"] = column_d.iloc[12]  # D14

                # Daya Ingat
                daya_ingat_value = row_data.get("Daya Ingat/ME", "")
                if len(self.sheet2_data) > 16:
                    if daya_ingat_value == "B":
                        row_data["Daya Ingat/ME.1"] = column_d.iloc[13]  # D15
                    elif daya_ingat_value == "C":
                        row_data["Daya Ingat/ME.1"] = column_d.iloc[14]  # D16
                    elif daya_ingat_value == "K":
                        row_data["Daya Ingat/ME.1"] = column_d.iloc[15]  # D17

                # Fleksibilitas
                fleksibilitas_value = row_data.get("Fleksibilitas/ T V", "")
                if len(self.sheet2_data) > 19:
                    if fleksibilitas_value == "B":
                        row_data["Fleksibilitas"] = column_d.iloc[16]  # D18
                    elif fleksibilitas_value == "C":
                        row_data["Fleksibilitas"] = column_d.iloc[17]  # D19
                    elif fleksibilitas_value == "K":
                        row_data["Fleksibilitas"] = column_d.iloc[18]  # D20

                # Sistematika Kerja
                sistematika_kerja_value = row_data.get("Sistematika Kerja/ cd", "")
                if len(self.sheet2_data) > 22:
                    if sistematika_kerja_value == "B":
                        row_data["Sistematika Kerja/ cd.1"] = column_d.iloc[19]  # D21
                    elif sistematika_kerja_value == "C":
                        row_data["Sistematika Kerja/ cd.1"] = column_d.iloc[20]  # D22
                    elif sistematika_kerja_value == "K":
                        row_data["Sistematika Kerja/ cd.1"] = column_d.iloc[21]  # D23

                # Inisiatif
                inisiatif_value = row_data.get("Inisiatif/W", "")
                if len(self.sheet2_data) > 25:
                    if inisiatif_value == "B":
                        row_data["Inisiatif/W.1"] = column_d.iloc[22]  # D24
                    elif inisiatif_value == "C":
                        row_data["Inisiatif/W.1"] = column_d.iloc[23]  # D25
                    elif inisiatif_value == "K":
                        row_data["Inisiatif/W.1"] = column_d.iloc[24]  # D26

                # Stabilitas Emosi
                stabilitas_emosi_value = row_data.get("Stabilitas Emosi / E", "")
                if len(self.sheet2_data) > 28:
                    if stabilitas_emosi_value == "B":
                        row_data["Stabilitas Emosi / E.1"] = column_d.iloc[25]  # D27
                    elif stabilitas_emosi_value == "C":
                        row_data["Stabilitas Emosi / E.1"] = column_d.iloc[26]  # D28
                    elif stabilitas_emosi_value == "K":
                        row_data["Stabilitas Emosi / E.1"] = column_d.iloc[27]  # D29

                # Komunikasi
                komunikasi_value = row_data.get("Komunikasi / B O", "")
                if len(self.sheet2_data) > 31:
                    if komunikasi_value == "B":
                        row_data["Komunikasi / B O.1"] = column_d.iloc[28]  # D30
                    elif komunikasi_value == "C":
                        row_data["Komunikasi / B O.1"] = column_d.iloc[29]  # D31
                    elif komunikasi_value == "K":
                        row_data["Komunikasi / B O.1"] = column_d.iloc[30]  # D32

                # Keterampilan Sosial
                keterampilan_sosial_value = row_data.get("Keterampilan Interpersonal / S O", "")
                if len(self.sheet2_data) > 34:
                    if keterampilan_sosial_value == "B":
                        row_data["Keterampilan Sosial / X S"] = column_d.iloc[31]  # D33
                    elif keterampilan_sosial_value == "C":
                        row_data["Keterampilan Sosial / X S"] = column_d.iloc[32]  # D34
                    elif keterampilan_sosial_value == "K":
                        row_data["Keterampilan Sosial / X S"] = column_d.iloc[33]  # D35

                # Kerjasama
                kerjasama_value = row_data.get("Kerjasama / B X", "")

                # Cek apakah jumlah data cukup
                if len(self.sheet2_data) >= 37:
                    kerjasama_value = row_data.get("Kerjasama / B X", "")
                    
                    column_d = self.sheet2_data["D"] if "D" in self.sheet2_data.columns else self.sheet2_data.iloc[:, 3]
                    
                    if kerjasama_value == "B" and pd.notna(column_d.iloc[34]):
                        row_data["Kerjasama"] = column_d.iloc[34]  # D36
                    elif kerjasama_value == "C" and pd.notna(column_d.iloc[35]):
                        row_data["Kerjasama"] = column_d.iloc[35]  # D37
                    elif kerjasama_value == "K" and pd.notna(column_d.iloc[36]):
                        row_data["Kerjasama"] = column_d.iloc[36]  # D38


            else:
                print("Column 'D' not found in Sheet2")

        except Exception as e:
            print(f"Error calculating values: {e}")

        # Determine action based on mode
        if mode == "add":
            # Add new row to the table
            row = self.table.rowCount()
            self.table.insertRow(row)
            for col, column_name in enumerate(self.columns):
                value = row_data.get(column_name, "")
                self.table.setItem(row, col, QTableWidgetItem(str(value)))
        elif mode == "edit":
            # Update existing row
            selected_row = self.table.currentRow()
            if selected_row >= 0:
                for col, column_name in enumerate(self.columns):
                    value = row_data.get(column_name, "")
                    self.table.setItem(selected_row, col, QTableWidgetItem(str(value)))
            else:
                print("No row selected for editing")

        # Clear input fields after adding/updating
        for field in self.personal_inputs + self.ist_inputs + self.papikostick_inputs:
            if isinstance(field, QLineEdit):
                field.clear()
            elif isinstance(field, QPushButton):
                field.setText("Pilih Tanggal")

    def recalculate_values(self, row):
        try:
            iq = self.get_cell_value(row, 5)
            verbal = self.get_cell_value(row, 7)
            daya_abstraksi_verbal = self.get_cell_value(row, 9)
            berpikir_praktis = self.get_cell_value(row, 10)
            berpikir_teoritis = self.get_cell_value(row, 11)

            if verbal is not None and daya_abstraksi_verbal is not None:
                wa_ge = (verbal + daya_abstraksi_verbal) / 2
                self.table.setItem(row, 13, QTableWidgetItem(str(wa_ge)))

            if berpikir_praktis is not None and berpikir_teoritis is not None:
                ra_zr = (berpikir_praktis + berpikir_teoritis) / 2
                self.table.setItem(row, 14, QTableWidgetItem(str(ra_zr)))

            if iq is not None:
                if iq < 79:
                    iq_klasifikasi = "Rendah"
                elif 79 <= iq < 90:
                    iq_klasifikasi = "Dibawah Rata-Rata"
                elif 90 <= iq < 110:
                    iq_klasifikasi = "Rata-Rata"
                elif 110 <= iq < 120:
                    iq_klasifikasi = "Diatas Rata-Rata"
                else:
                    iq_klasifikasi = "Superior"
                self.table.setItem(row, 15, QTableWidgetItem(iq_klasifikasi))
        except Exception as e:
            print(f"Kesalahan dalam perhitungan ulang: {e}")

    def get_cell_value(self, row, col):
        item = self.table.item(row, col)
        if item and item.text().isdigit():
            return int(item.text())
        return None

    def delete_selected_row(self):
        selected_row = self.table.currentRow()
        if selected_row >= 0:
            self.table.removeRow(selected_row)
            self.table.resizeColumnsToContents()
            
    def print_selected_row(self):
        selected_row = self.table.currentRow()
        if selected_row >= 0:
            row_data = {}
            for col, column_name in enumerate(self.columns):
                item = self.table.item(selected_row, col)
                row_data[column_name] = item.text() if item else ""
            
            df_print = pd.DataFrame([row_data])
            temp_file = "temp_print.xlsx"
            df_print.to_excel(temp_file, index=False, engine="openpyxl")
            
            import os
            os.startfile(temp_file, "print")

    def save_to_excel(self):
        if not self.excel_file_path:
            return

        try:
            # Buat direktori backup jika belum ada
            import os
            import shutil
            import copy
            from datetime import datetime
            
            # Simpan jalur file asli
            original_file = self.excel_file_path
            dir_path = os.path.dirname(original_file)
            file_name = os.path.basename(original_file)
            file_name_without_ext, file_ext = os.path.splitext(file_name)
            
            # Buat folder backup
            backup_dir = os.path.join(dir_path, "backup")
            os.makedirs(backup_dir, exist_ok=True)
            
            # Buat backup file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(backup_dir, f"backup_{timestamp}_{file_name}")
            shutil.copy2(original_file, backup_file)
            
            # Buka file Excel langsung
            import openpyxl
            from openpyxl.styles import Border, Side, Alignment, Protection, Font
            from openpyxl.utils.cell import get_column_letter
            
            # Buka workbook - pastikan parameter yang benar
            # data_only=False untuk mempertahankan formula
            # keep_vba=False untuk menghindari masalah korupsi file
            wb = openpyxl.load_workbook(original_file, data_only=False, keep_vba=False)
            
            # Pilih sheet utama (biasanya Sheet1)
            if "Sheet1" in wb.sheetnames:
                sheet = wb["Sheet1"]
            else:
                sheet = wb.active
            
            print(f"File Excel dibuka: {original_file}")
            
            # Cari baris header yang berisi "No"
            header_row = None
            for row_idx in range(1, min(sheet.max_row + 1, 100)):
                for col_idx in range(1, min(sheet.max_column + 1, 20)):
                    cell_value = sheet.cell(row=row_idx, column=col_idx).value
                    if cell_value is not None and str(cell_value).strip() == "No":
                        header_row = row_idx
                        break
                if header_row:
                    break
            
            if not header_row:
                header_row = 3  # Default jika tidak ditemukan
                print(f"Header tidak ditemukan, menggunakan baris {header_row}")
            else:
                print(f"Header ditemukan di baris {header_row}")
            
            # Buat pemetaan antara nama kolom di file Excel dan indeks kolom
            col_mapping = {}
            
            # Simpan sel dengan formula agar bisa dipertahankan
            formula_cells = {}
            
            # Baris data sampel pertama (jika ada) untuk mendapatkan formula
            formula_cols_to_check = [43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55]  # Kolom yang perlu diperiksa khusus
            
            if sheet.max_row > header_row:
                sample_row = header_row + 1
                for col_idx in range(1, sheet.max_column + 1):
                    # Prioritaskan kolom yang perlu diperiksa khusus
                    is_prioritized = col_idx in formula_cols_to_check
                    
                    cell = sheet.cell(row=sample_row, column=col_idx)
                    if cell.value is not None and isinstance(cell.value, str) and cell.value.startswith('='):
                        formula = cell.value
                        
                        # Cek apakah ini formula IFS yang memerlukan penanganan khusus
                        if ('IFS(' in formula or '_xlfn.IFS(' in formula or '@IFS(' in formula):
                            # Simpan template formula IFS untuk digunakan nanti
                            formula_template = formula
                            # Ganti referensi baris dengan placeholder {row}
                            formula_parts = []
                            current_pos = 0
                            # Cari semua referensi sel dalam formula
                            import re
                            cell_refs = re.findall(r'[A-Z]+\d+', formula)
                            for ref in cell_refs:
                                # Pisahkan kolom dan baris
                                col_part = ''.join(c for c in ref if c.isalpha())
                                row_part = ''.join(c for c in ref if c.isdigit())
                                # Ganti dengan template
                                formula = formula.replace(ref, col_part + "{row}")
                            
                            formula_cells[col_idx] = {
                                'type': 'IFS',
                                'template': formula,
                                'original': cell.value
                            }
                            print(f"Formula IFS ditemukan di kolom {col_idx}: {cell.value}")
                            print(f"Template formula: {formula}")
                        else:
                            # Formula biasa
                            formula_cells[col_idx] = {
                                'type': 'normal',
                                'formula': formula
                            }
                            print(f"Formula normal ditemukan di kolom {col_idx}: {formula}")
                    elif is_prioritized:
                        # Jika tidak ditemukan formula di kolom prioritas, cek baris berikutnya
                        for next_row in range(header_row + 2, min(header_row + 6, sheet.max_row + 1)):
                            next_cell = sheet.cell(row=next_row, column=col_idx)
                            if next_cell.value is not None and isinstance(next_cell.value, str) and next_cell.value.startswith('='):
                                formula = next_cell.value
                                # Analisis dan simpan formula
                                if ('IFS(' in formula or '_xlfn.IFS(' in formula or '@IFS(' in formula):
                                    # Proses formula IFS
                                    import re
                                    # Cari semua referensi sel dalam formula
                                    cell_refs = re.findall(r'[A-Z]+\d+', formula)
                                    # Dapatkan baris asli
                                    original_row = next_row
                                    # Buat template dengan mengganti nomor baris dengan {row}
                                    template_formula = formula
                                    for ref in cell_refs:
                                        col_part = ''.join(c for c in ref if c.isalpha())
                                        row_part = ''.join(c for c in ref if c.isdigit())
                                        # Ganti hanya jika baris sama dengan baris formula
                                        if row_part == str(original_row):
                                            template_formula = template_formula.replace(ref, col_part + "{row}")
                                    
                                    formula_cells[col_idx] = {
                                        'type': 'IFS',
                                        'template': template_formula,
                                        'original': formula,
                                        'original_row': original_row
                                    }
                                    print(f"Formula IFS ditemukan di kolom {col_idx} (baris alternatif {next_row}): {formula}")
                                    print(f"Template formula: {template_formula}")
                                else:
                                    # Formula biasa
                                    formula_cells[col_idx] = {
                                        'type': 'normal',
                                        'formula': formula,
                                        'original_row': next_row
                                    }
                                    print(f"Formula normal ditemukan di kolom {col_idx} (baris alternatif {next_row}): {formula}")
                                break
                
                # Cek dan tambahkan formula yang umum digunakan di kolom psikogram jika belum ada
                missing_psikogram_formulas = {
                    43: "=IF(F{row}<90,\"K\",IF(F{row}<110,\"C\",\"B\"))",  # Kemampuan Numerik
                    44: "=IF(I{row}<90,\"K\",IF(I{row}<110,\"C\",\"B\"))",  # Daya Ingat
                    45: "=IF(N{row}<90,\"K\",IF(N{row}<110,\"C\",\"B\"))",  # Fleksibilitas
                    46: "=IF(O{row}<90,\"K\",IF(O{row}<110,\"C\",\"B\"))",  # Sistematika
                    47: "=IF(M{row}<90,\"K\",IF(M{row}<110,\"C\",\"B\"))",  # Inisiatif
                    48: "=IFS(AM{row}<4,\"K\",AM{row}<6,\"C\",AM{row}>5,\"B\")",  # Stabilitas Emosi
                    49: "=IFS(AL{row}<4,\"K\",AL{row}<6,\"C\",AL{row}>5,\"B\")",  # Komunikasi
                    50: "=IFS(AK{row}<4,\"B\",AK{row}<6,\"C\",AK{row}>5,\"K\")",  # Inisiatif/W
                    51: "=IFS(AH{row}<4,\"B\",AH{row}<6,\"C\",AH{row}>5,\"K\")",  # Stabilitas Emosi/E
                    52: "=IFS(AN{row}<4,\"K\",AN{row}<6,\"C\",AN{row}>5,\"B\")",  # Kolom 52 dengan formula yang benar
                    53: "=IFS(AO{row}<4,\"K\",AO{row}<6,\"C\",AO{row}>5,\"B\")",  # Kolom 53 dengan formula yang benar
                    54: "=IFS(AP{row}<4,\"K\",AP{row}<6,\"C\",AP{row}>5,\"B\")",  # Kolom 54 dengan formula yang benar
                    55: "=IFS(AQ{row}=\"B\",Sheet2!$D$3,AQ{row}=\"C\",Sheet2!$D$4,AQ{row}=\"K\",Sheet2!$D$5)",  # Kolom 55 dengan formula yang benar
                    56: "=IFS(AR{row}=\"B\",Sheet2!$D$6,AR{row}=\"C\",Sheet2!$D$7,AR{row}=\"K\",Sheet2!$D$8)",  # Daya Analisa/ AN.1
                    57: "=IFS(AS{row}=\"B\",Sheet2!$D$9,AS{row}=\"C\",Sheet2!$D$10,AS{row}=\"K\",Sheet2!$D$11)",  # Kemampuan Verbal/WA GE.1
                    58: "=IFS(AT{row}=\"B\",Sheet2!$D$12,AT{row}=\"C\",Sheet2!$D$13,AT{row}=\"K\",Sheet2!$D$14)",  # Kemampuan Numerik/ RA ZR.1
                    59: "=IFS(AU{row}=\"B\",Sheet2!$D$15,AU{row}=\"C\",Sheet2!$D$16,AU{row}=\"K\",Sheet2!$D$17)",  # Daya Ingat/ME.1
                    60: "=IFS(AV{row}=\"B\",Sheet2!$D$18,AV{row}=\"C\",Sheet2!$D$19,AV{row}=\"K\",Sheet2!$D$20)",  # Fleksibilitas
                    61: "=IFS(AW{row}=\"B\",Sheet2!$D$21,AW{row}=\"C\",Sheet2!$D$22,AW{row}=\"K\",Sheet2!$D$23)",  # Sistematika Kerja/ cd.1
                    62: "=IFS(AX{row}=\"B\",Sheet2!$D$24,AX{row}=\"C\",Sheet2!$D$25,AX{row}=\"K\",Sheet2!$D$26)",  # Inisiatif/W.1
                    63: "=IFS(AY{row}=\"B\",Sheet2!$D$27,AY{row}=\"C\",Sheet2!$D$28,AY{row}=\"K\",Sheet2!$D$29)",  # Stabilitas Emosi / E.1
                    64: "=IFS(AZ{row}=\"B\",Sheet2!$D$30,AZ{row}=\"C\",Sheet2!$D$31,AZ{row}=\"K\",Sheet2!$D$32)",  # Komunikasi / B O.1
                    65: "=IFS(BA{row}=\"B\",Sheet2!$D$33,BA{row}=\"C\",Sheet2!$D$34,BA{row}=\"K\",Sheet2!$D$35)",  # Keterampilan Sosial
                    66: "=IFS(BB{row}=\"B\",Sheet2!$D$36,BB{row}=\"C\",Sheet2!$D$37,BB{row}=\"K\",Sheet2!$D$38)"   # Kerjasama
                }
                
                # Tambahkan formula dari file asli jika ditemukan di kolom mana pun
                for row_idx in range(header_row + 1, min(sheet.max_row + 1, header_row + 10)):
                    for col_idx in range(52, 56):  # Cek khusus kolom 52-55
                        if col_idx not in formula_cells:
                            cell = sheet.cell(row=row_idx, column=col_idx)
                            if cell.value is not None and isinstance(cell.value, str) and cell.value.startswith('='):
                                formula = cell.value
                                
                                if ('IFS(' in formula or '_xlfn.IFS(' in formula or '@IFS(' in formula):
                                    # Proses formula IFS
                                    import re
                                    cell_refs = re.findall(r'[A-Z]+\d+', formula)
                                    template_formula = formula
                                    for ref in cell_refs:
                                        col_part = ''.join(c for c in ref if c.isalpha())
                                        row_part = ''.join(c for c in ref if c.isdigit())
                                        if row_part == str(row_idx):
                                            template_formula = template_formula.replace(ref, col_part + "{row}")
                                    
                                    formula_cells[col_idx] = {
                                        'type': 'IFS',
                                        'template': template_formula,
                                        'original': formula,
                                        'original_row': row_idx
                                    }
                                    print(f"Formula asli ditemukan di kolom {col_idx} (baris {row_idx}): {formula}")
                                    print(f"Template: {template_formula}")
                                else:
                                    formula_cells[col_idx] = {
                                        'type': 'normal',
                                        'formula': formula,
                                        'original_row': row_idx
                                    }
                                    print(f"Formula normal ditemukan di kolom {col_idx} (baris {row_idx}): {formula}")
                
                # Tambahkan formula yang hilang
                for col, formula_template in missing_psikogram_formulas.items():
                    if col not in formula_cells:
                        formula_cells[col] = {
                            'type': 'IFS',
                            'template': formula_template,
                            'original': formula_template.replace("{row}", str(header_row + 1)),
                            'is_generated': True
                        }
                        print(f"Menambahkan formula IFS yang hilang untuk kolom {col}: {formula_template}")
                
                # Buat dictionary untuk menyimpan hasil evaluasi formula
                formula_results = {}
                
                # Coba evaluasi formula IFS untuk mendapatkan hasilnya
                try:
                    # Hanya jika file sudah memiliki data
                    if sheet.max_row > header_row + 1:
                        # Ambil sampel baris pertama untuk evaluasi
                        sample_row = header_row + 1
                        
                        # Cek nilai-nilai di kolom yang digunakan dalam formula
                        col_values = {}
                        for col_letter in ['F', 'I', 'N', 'O', 'M', 'AK', 'AL', 'AF', 'AG', 'AH', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB']:
                            col_index = openpyxl.utils.column_index_from_string(col_letter)
                            cell = sheet.cell(row=sample_row, column=col_index)
                            # Simpan nilai numeriknya jika bisa
                            try:
                                col_values[col_letter] = float(cell.value) if cell.value else 0
                            except (ValueError, TypeError):
                                col_values[col_letter] = cell.value if cell.value else ""
                        
                        # Evaluasi formula untuk kolom 43-54
                        for col_idx in range(43, 55):
                            if col_idx in missing_psikogram_formulas:
                                formula_template = missing_psikogram_formulas[col_idx]
                                # Ekstrak kolom referensi (mis. AK, AL, dll.)
                                ref_col = None
                                
                                # Untuk formula IF
                                if 'IF(' in formula_template and not 'IFS(' in formula_template:
                                    ref_match = re.search(r'([A-Z]{1,2})\{row\}', formula_template)
                                    if ref_match:
                                        ref_col = ref_match.group(1)
                                        if ref_col and ref_col in col_values:
                                            value = col_values[ref_col]
                                            if isinstance(value, (int, float)):
                                                if value < 90:
                                                    formula_results[col_idx] = "K"
                                                elif value < 110:
                                                    formula_results[col_idx] = "C"
                                                else:
                                                    formula_results[col_idx] = "B"
                                                print(f"Hasil evaluasi formula IF kolom {col_idx}: {formula_results[col_idx]} (ref:{ref_col}={value})")
                                
                                # Untuk formula IFS (50-51 memiliki pola berbeda)
                                elif col_idx in [50, 51]:
                                    # Ekstrak kolom referensi (mis. AK, AH, dll.)
                                    ref_match = re.search(r'([A-Z]{1,2})\{row\}', formula_template)
                                    if ref_match:
                                        ref_col = ref_match.group(1)
                                        if ref_col and ref_col in col_values:
                                            value = col_values[ref_col]
                                            if isinstance(value, (int, float)):
                                                if value < 4:
                                                    formula_results[col_idx] = "B"
                                                elif value < 6:
                                                    formula_results[col_idx] = "C"
                                                elif value > 5:
                                                    formula_results[col_idx] = "K"
                                                else:
                                                    formula_results[col_idx] = ""
                                                print(f"Hasil evaluasi formula IFS terbalik kolom {col_idx}: {formula_results[col_idx]} (ref:{ref_col}={value})")
                                
                                # Untuk formula IFS standar
                                elif '<4' in formula_template and '<6' in formula_template:
                                    # Format formula standar: IFS(XX{row}<4,"K",XX{row}<6,"C",XX{row}>5,"B")
                                    ref_match = re.search(r'([A-Z]{1,2})\{row\}', formula_template)
                                    if ref_match:
                                        ref_col = ref_match.group(1)
                                        if ref_col and ref_col in col_values:
                                            value = col_values[ref_col]
                                            if isinstance(value, (int, float)):
                                                if value < 4:
                                                    formula_results[col_idx] = "K"
                                                elif value < 6:
                                                    formula_results[col_idx] = "C"
                                                elif value > 5:
                                                    formula_results[col_idx] = "B"
                                                else:
                                                    formula_results[col_idx] = ""
                                                print(f"Hasil evaluasi formula IFS standar kolom {col_idx}: {formula_results[col_idx]} (ref:{ref_col}={value})")
                        
                        # Evaluasi khusus untuk kolom 55-66 (referensi ke Sheet2)
                        for col_idx in range(55, 67):
                            ref_col_letter = None
                            
                            # Tentukan kolom referensi untuk setiap kolom special
                            if col_idx == 55:
                                ref_col_letter = 'AQ'
                            elif col_idx == 56:
                                ref_col_letter = 'AR'
                            elif col_idx == 57:
                                ref_col_letter = 'AS'
                            elif col_idx == 58:
                                ref_col_letter = 'AT'
                            elif col_idx == 59:
                                ref_col_letter = 'AU'
                            elif col_idx == 60:
                                ref_col_letter = 'AV'
                            elif col_idx == 61:
                                ref_col_letter = 'AW'
                            elif col_idx == 62:
                                ref_col_letter = 'AX'
                            elif col_idx == 63:
                                ref_col_letter = 'AY'
                            elif col_idx == 64:
                                ref_col_letter = 'AZ'
                            elif col_idx == 65:
                                ref_col_letter = 'BA'
                            elif col_idx == 66:
                                ref_col_letter = 'BB'
                            
                            if ref_col_letter and ref_col_letter in col_values:
                                ref_value = col_values.get(ref_col_letter)
                                if ref_value in ["B", "C", "K"]:
                                    formula_results[col_idx] = ref_value
                                    print(f"Hasil evaluasi formula Sheet2 kolom {col_idx}: {formula_results[col_idx]} (ref:{ref_col_letter}={ref_value})")
                
                except Exception as e:
                    print(f"Error saat mengevaluasi formula: {e}")
                
                # Cek baris-baris berikutnya untuk memverifikasi formula IFS yang berbeda
                for sample_offset in range(2, min(6, sheet.max_row - header_row + 1)):
                    check_row = header_row + sample_offset
                    for col_idx, formula_info in formula_cells.items():
                        if isinstance(formula_info, dict) and formula_info.get('type') == 'IFS':
                            cell = sheet.cell(row=check_row, column=col_idx)
                            if cell.value is not None and isinstance(cell.value, str) and cell.value.startswith('='):
                                print(f"Formula IFS di baris {check_row}, kolom {col_idx}: {cell.value}")
            
            # Deteksi kolom IQ dan Flexibilitas Pikir dengan spasi
            iq_column = None
            flex_column = None
            
            # Deteksi kolom Unnamed untuk pemetaan khusus
            unnamed_columns = {}
            
            # Kolom-kolom yang perlu pemetaan khusus
            special_columns = {
                "Unnamed: 13": 14,  # Perhatikan perbedaan indeks: kolom 14 di Excel = indeks 13 di array
                "Unnamed: 14": 15,
                "Intelegensi Umum.1": 55,
                "Daya Analisa/ AN.1": 56,
                "Kemampuan Verbal/WA GE.1": 57,
                "Kemampuan Numerik/ RA ZR.1": 58,
                "Daya Ingat/ME.1": 59,
                "Sistematika Kerja/ cd.1": 61,
                "Inisiatif/W.1": 62,
                "Stabilitas Emosi / E.1": 63,
                "Komunikasi / B O.1": 64
            }
            
            # Pemetaan untuk kolom Psikogram (bagian POTENSI INTELEKTUAL dan KEPRIBADIAN)
            psikogram_columns = {
                "Intelegensi Umum": 43,  
                "Daya Analisa/ AN": 44,
                "Kemampuan Verbal/WA GE": 45,
                "Kemampuan Numerik/ RA ZR": 46,
                "Daya Ingat/ME": 47,
                "Fleksibilitas/ T V": 48,
                "Sistematika Kerja/ cd": 49,
                "Inisiatif/W": 50,
                "Stabilitas Emosi / E": 51,
                "Komunikasi / B O": 52,
                "Keterampilan Interpersonal / S O": 53,
                "Kerjasama / B X": 54
            }
            
            # Pisahkan kolom untuk melacak secara eksplisit
            potensi_intelektual_columns = [
                "Intelegensi Umum", 
                "Daya Analisa/ AN", 
                "Kemampuan Verbal/WA GE", 
                "Kemampuan Numerik/ RA ZR", 
                "Daya Ingat/ME", 
                "Fleksibilitas/ T V"
            ]
            
            kepribadian_columns = [
                "Sistematika Kerja/ cd", 
                "Inisiatif/W", 
                "Stabilitas Emosi / E", 
                "Komunikasi / B O", 
                "Keterampilan Interpersonal / S O", 
                "Kerjasama / B X"
            ]
            
            for col_idx in range(1, sheet.max_column + 1):
                cell_value = sheet.cell(row=header_row, column=col_idx).value
                if cell_value is not None:
                    col_name = str(cell_value)
                    col_name_stripped = col_name.strip()
                    
                    # Simpan semua pemetaan kolom
                    col_mapping[col_name] = col_idx
                    
                    # Deteksi kolom Unnamed
                    if col_name.startswith("Unnamed:"):
                        unnamed_columns[col_name] = col_idx
                        print(f"Kolom Unnamed ditemukan: '{col_name}' di kolom {col_idx}")
                    
                    # Deteksi kolom IQ dengan spasi
                    if col_name_stripped == "IQ" and (col_name.startswith(" ") or col_name.endswith(" ")):
                        iq_column = col_idx
                        print(f"Kolom IQ (dengan spasi) ditemukan: '{col_name}' di kolom {col_idx}")
                    
                    # Deteksi kolom Flexibilitas Pikir dengan spasi
                    if col_name_stripped == "Flexibilitas Pikir" and (col_name.startswith(" ") or col_name.endswith(" ")):
                        flex_column = col_idx
                        print(f"Kolom Flexibilitas Pikir (dengan spasi) ditemukan: '{col_name}' di kolom {col_idx}")
                    
                    print(f"Pemetaan kolom: '{col_name}' -> {col_idx}")
            
            # Pemetaan kolom Unnamed secara manual jika tidak ditemukan
            if 14 not in unnamed_columns.values():
                unnamed_columns["Unnamed: 13"] = 14
                print(f"Memetakan secara manual: 'Unnamed: 13' -> 14")
            
            if 15 not in unnamed_columns.values():
                unnamed_columns["Unnamed: 14"] = 15
                print(f"Memetakan secara manual: 'Unnamed: 14' -> 15")
            
            # Tentukan kolom "No"
            no_col_idx = col_mapping.get("No", 1)
            
            # Tentukan kolom duplikat yang akan dihapus dari UI sebelum menyimpan
            duplicate_columns = []
            
            # Cari kolom IQ dan Flexibilitas Pikir tanpa spasi di akhir daftar kolom
            last_columns = list(self.columns)[-10:]  # Ambil 10 kolom terakhir
            
            for i, col_name in enumerate(last_columns):
                if col_name.strip() == "IQ" and col_name != "IQ ":
                    idx = len(self.columns) - 10 + i
                    duplicate_columns.append(idx)
                    print(f"Kolom duplikat IQ ditemukan di indeks {idx}: '{col_name}'")
                
                if col_name.strip() == "Flexibilitas Pikir" and col_name != " Flexibilitas Pikir":
                    idx = len(self.columns) - 10 + i
                    duplicate_columns.append(idx)
                    print(f"Kolom duplikat Flexibilitas Pikir ditemukan di indeks {idx}: '{col_name}'")
            
            # Simpan informasi tentang tinggi baris
            row_heights = {}
            for row_idx in range(1, min(sheet.max_row + 1, header_row + 10)):
                if row_idx in sheet.row_dimensions and sheet.row_dimensions[row_idx].height is not None:
                    row_heights[row_idx] = sheet.row_dimensions[row_idx].height
            
            # Simpan tinggi baris default dari baris data pertama (jika ada)
            default_row_height = None
            if header_row + 1 in sheet.row_dimensions and sheet.row_dimensions[header_row + 1].height is not None:
                default_row_height = sheet.row_dimensions[header_row + 1].height
                print(f"Tinggi baris default (dari baris data pertama): {default_row_height}")
            else:
                default_row_height = 15  # Tinggi default Excel biasanya sekitar 15
                print(f"Menggunakan tinggi baris default: {default_row_height}")
            
            # Simpan informasi tentang merged cells
            merged_ranges = list(sheet.merged_cells.ranges)
            print(f"Jumlah merged cells: {len(merged_ranges)}")
            
            # Hapus semua merged cells sebelum menghapus baris (untuk menghindari error)
            for merge_range in merged_ranges:
                sheet.unmerge_cells(str(merge_range))
            
            # Hapus semua data yang ada di bawah header
            # Ini akan memastikan kita tidak menduplikasi data
            data_rows_to_delete = []
            for row_idx in range(header_row + 1, sheet.max_row + 1):
                cell_value = sheet.cell(row=row_idx, column=no_col_idx).value
                if cell_value is not None:
                    data_rows_to_delete.append(row_idx)
            
            # Hapus dari baris terbawah ke atas untuk menghindari pergeseran indeks
            for row_idx in sorted(data_rows_to_delete, reverse=True):
                sheet.delete_rows(row_idx)
            
            print(f"Menghapus {len(data_rows_to_delete)} baris data lama")
            
            # Tambahan untuk menangani warna khusus pada file seperti yang terlihat di screenshot
            # Identifikasi area header utama (seperti IST, PAPIKOSTIK, PSIKOGRAM)
            special_headers = ["IST", "PAPIKOSTIK", "PSIKOGRAM", "POTESI INTELEKTUAL", "SIKAP DAN CARA KERJA"]
            header_colors = {}
            
            # Cari row dan kolom untuk header khusus
            for row_idx in range(1, header_row):
                for col_idx in range(1, sheet.max_column + 1):
                    cell_value = sheet.cell(row=row_idx, column=col_idx).value
                    if cell_value and any(header in str(cell_value).upper() for header in special_headers):
                        header_colors[str(cell_value)] = {
                            'row': row_idx,
                            'col': col_idx,
                            'fill': copy.copy(sheet.cell(row=row_idx, column=col_idx).fill)
                        }
                        print(f"Header khusus ditemukan: {cell_value} di baris {row_idx}, kolom {col_idx}")
            
            # Simpan informasi ini untuk digunakan nanti (jika perlu)
            self.header_colors = header_colors
            
            # Simpan untuk sel data pada baris pertama (sebagai referensi)
            special_data_styles = {}
            if sheet.max_row > header_row:
                first_data_row = header_row + 1
                for col_idx in range(1, sheet.max_column + 1):
                    cell = sheet.cell(row=first_data_row, column=col_idx)
                    if cell.has_style and cell.fill and cell.fill.start_color and cell.fill.start_color.index:
                        special_data_styles[col_idx] = copy.copy(cell.fill)
            
            # Simpan informasi tentang sel-sel data sampel
            # Untuk setiap baris di bawah header (jika ada), simpan sebagai template
            template_rows = {}
            sample_row_count = min(5, sheet.max_row - header_row)  # Maksimal 5 baris template
            for offset in range(1, sample_row_count + 1):
                sample_row = header_row + offset
                template_rows[offset] = {}
                
                for col_idx in range(1, sheet.max_column + 1):
                    source_cell = sheet.cell(row=sample_row, column=col_idx)
                    if source_cell.has_style:
                        template_rows[offset][col_idx] = {
                            'font': copy.copy(source_cell.font),
                            'border': copy.copy(source_cell.border),
                            'fill': copy.copy(source_cell.fill),
                            'number_format': source_cell.number_format,
                            'alignment': copy.copy(source_cell.alignment),
                            'value_type': type(source_cell.value)
                        }
            
            # Jika tidak ada baris sampel, gunakan style header sebagai default
            if not template_rows:
                col_styles = {}
                for col_idx in range(1, sheet.max_column + 1):
                    source_cell = sheet.cell(row=header_row, column=col_idx)
                    if source_cell.has_style:
                        col_styles[col_idx] = {
                            'font': copy.copy(source_cell.font),
                            'border': copy.copy(source_cell.border),
                            'fill': copy.copy(source_cell.fill),
                            'number_format': source_cell.number_format,
                            'alignment': copy.copy(source_cell.alignment)
                        }
                template_rows[1] = col_styles
            
            # Kumpulkan data dari tabel UI (hapus kolom duplikat)
            ui_data = []
            for row in range(self.table.rowCount()):
                row_data = {}
                for col, column_name in enumerate(self.columns):
                    # Lewati kolom duplikat
                    if col in duplicate_columns:
                        continue
                    
                    item = self.table.item(row, col)
                    value = item.text() if item else ""
                    row_data[column_name] = value
                    
                    # Cetak nilai dari kolom psikogram untuk debugging
                    if column_name in potensi_intelektual_columns or column_name in kepribadian_columns:
                        print(f"Kolom Psikogram: {column_name} = '{value}'")
                
                ui_data.append(row_data)
            
            # Tambahkan data baru ke worksheet
            for i, row_data in enumerate(ui_data):
                # Target baris adalah baris header + 1 + indeks data
                target_row = header_row + 1 + i
                template_idx = (i % len(template_rows)) + 1 if template_rows else 1
                print(f"Menambahkan data baru di baris {target_row} menggunakan template {template_idx}")
                
                # Salin style dari template baris ke baris baru
                for col_idx in range(1, sheet.max_column + 1):
                    target_cell = sheet.cell(row=target_row, column=col_idx)
                    
                    # Determine style to use from template rows or default to None
                    style = None
                    if template_idx in template_rows and col_idx in template_rows[template_idx]:
                        style = template_rows[template_idx][col_idx]
                    else:
                        # Try to get style from header row if available
                        header_cell = sheet.cell(row=header_row, column=col_idx)
                        if header_cell.has_style:
                            style = {
                                'font': copy.copy(header_cell.font),
                                'border': copy.copy(header_cell.border), 
                                'fill': copy.copy(header_cell.fill),
                                'number_format': header_cell.number_format,
                                'alignment': copy.copy(header_cell.alignment)
                            }
                    
                    if not style:
                        continue
                    
                    # Terapkan style tapi hindari bold
                    if style:
                        try:
                            # Font - salin semua properti kecuali bold
                            if hasattr(style['font'], 'name') and style['font'].name:
                                new_font = copy.copy(style['font'])
                                # Set bold ke False untuk data (bukan header)
                                if target_row > header_row:  # Gunakan target_row bukan row_idx
                                    new_font.bold = False
                                target_cell.font = new_font
                            
                            # Border - pastikan semua border disalin dengan benar
                            if style['border']:
                                target_cell.border = copy.copy(style['border'])
                            
                            # Fill - pastikan warna latar belakang disalin dengan benar
                            if style['fill'] and style['fill'].fill_type:
                                target_cell.fill = copy.copy(style['fill'])
                            
                            # Format angka
                            if style['number_format']:
                                target_cell.number_format = style['number_format']
                            
                            # Alignment - pastikan orientasi teks benar (horizontal)
                            if style['alignment']:
                                alignment = copy.copy(style['alignment'])
                                # Pastikan orientasi selalu horizontal untuk data
                                if target_row > header_row:
                                    alignment.text_rotation = 0  # 0 derajat = horizontal
                                    alignment.vertical = 'center'  # Tengah secara vertikal
                                target_cell.alignment = alignment
                        except Exception as e:
                            print(f"Error applying style: {e}")
                
                # Isi data dari UI ke baris target
                for col_name, value in row_data.items():
                    col_name_stripped = col_name.strip()
                    col_idx = None
                    
                    # Handling khusus untuk kolom-kolom tertentu berdasarkan nama
                    if col_name in special_columns:
                        col_idx = special_columns[col_name]
                        print(f"Menggunakan pemetaan khusus untuk kolom '{col_name}' -> {col_idx}")
                    # Gunakan kolom yang benar untuk IQ dan Flexibilitas Pikir
                    elif col_name_stripped == "IQ" and iq_column:
                        col_idx = iq_column
                    elif col_name_stripped == "Flexibilitas Pikir" and flex_column:
                        col_idx = flex_column
                    # Khusus untuk kolom psikogram
                    elif col_name in psikogram_columns:
                        col_idx = psikogram_columns[col_name]
                        print(f"Menggunakan pemetaan psikogram untuk '{col_name}' -> {col_idx}")
                    # Khusus untuk kolom Unnamed
                    elif col_name.startswith("Unnamed:"):
                        # Coba dari unnamed_columns
                        if col_name in unnamed_columns:
                            col_idx = unnamed_columns[col_name]
                        else:
                            # Coba cari angka di belakang Unnamed:
                            try:
                                num = int(col_name.split(":")[-1].strip())
                                col_idx = num + 1  # Angka kolom Excel biasanya +1 dari indeks
                                print(f"Memetakan berdasarkan angka: '{col_name}' -> {col_idx}")
                            except:
                                pass
                    elif col_name in col_mapping:
                        col_idx = col_mapping[col_name]
                    else:
                        # Coba cari dengan nama yang sudah di-strip
                        for name, idx in col_mapping.items():
                            if name.strip() == col_name_stripped:
                                col_idx = idx
                                break
                    
                    if not col_idx:
                        # Khusus untuk kolom dengan akhiran .1 (duplicate columns)
                        if ".1" in col_name:
                            base_name = col_name.replace(".1", "")
                            if base_name in col_mapping:
                                base_col = col_mapping[base_name]
                                # Cari kolom setelah kolom dasar dengan nama yang mirip
                                for idx in range(base_col + 1, sheet.max_column + 1):
                                    cell_val = sheet.cell(row=header_row, column=idx).value
                                    if cell_val is not None and base_name in str(cell_val):
                                        col_idx = idx
                                        print(f"Memetakan kolom duplikat '{col_name}' ke '{cell_val}' di kolom {col_idx}")
                                        break
                        
                        if not col_idx:
                            print(f"Tidak dapat menemukan kolom untuk '{col_name}'")
                            continue
                    
                    target_cell = sheet.cell(row=target_row, column=col_idx)
                    
                    # Periksa apakah kolom ini memiliki formula
                    if col_idx in formula_cells:
                        # Gunakan formula, ganti nilai ref baris jika perlu
                        try:
                            formula_info = formula_cells[col_idx]
                            
                            if isinstance(formula_info, dict):
                                # Penanganan formula berdasarkan tipe
                                if formula_info.get('type') == 'IFS':
                                    # Untuk formula IFS, kita perlu buat formula yang benar untuk baris ini
                                    template = formula_info.get('template')
                                    
                                    # Ganti {row} dengan nomor baris saat ini
                                    formula = template.replace("{row}", str(target_row))
                                    
                                    # Perbaiki format formula IFS (hilangkan _xlfn. atau @ jika perlu)
                                    # Gunakan format yang kompatibel dengan Excel
                                    if '_xlfn.IFS' in formula:
                                        # Excel modern menggunakan IFS langsung
                                        formula = formula.replace('_xlfn.IFS', 'IFS')
                                    
                                    # Jika ada tanda @ di awal formula, hapus
                                    if formula.startswith('=@'):
                                        formula = '=' + formula[2:]
                                    
                                    # Cek apakah kolom ini punya hasil evaluasi formula
                                    if col_idx in formula_results:
                                        # Gunakan hasil langsung alih-alih formula untuk menghindari #NAME?
                                        # Kolom 50-54 harus selalu menggunakan formula, bukan hasil evaluasi
                                        if col_idx >= 50 and col_idx <= 54:
                                            # Gunakan formula untuk kolom 50-54
                                            target_cell.value = formula
                                            print(f"Menetapkan formula IFS ke baris {target_row}, kolom {col_idx}: {formula}")
                                        else:
                                            target_cell.value = formula_results[col_idx]
                                            print(f"Menggunakan hasil evaluasi {formula_results[col_idx]} untuk kolom {col_idx} (IFS)")
                                    else:
                                        # Gunakan formula jika tidak ada hasil evaluasi
                                        target_cell.value = formula
                                        print(f"Menetapkan formula IFS ke baris {target_row}, kolom {col_idx}: {formula}")
                                else:
                                    # Formula normal
                                    formula = formula_info.get('formula', '')
                                    # Perbarui referensi baris
                                    import re
                                    
                                    # Cari referensi sel dan perbarui nomornya
                                    def update_cell_reference(match):
                                        ref = match.group(0)
                                        col_part = ''.join(c for c in ref if c.isalpha())
                                        row_part = ''.join(c for c in ref if c.isdigit())
                                        if row_part == str(header_row + 1):
                                            return col_part + str(target_row)
                                        return ref
                                    
                                    formula = re.sub(r'[A-Z]+\d+', update_cell_reference, formula)
                                    target_cell.value = formula
                                    print(f"Menetapkan formula normal ke baris {target_row}, kolom {col_idx}: {formula}")
                            else:
                                # Format lama - jika formula_cells adalah string langsung bukan dict
                                formula = formula_info
                                formula = formula.replace(str(header_row + 1), str(target_row))
                                
                                # Beberapa formula Excel mungkin bermasalah (#NAME?), coba tampilkan hasil langsung
                                if col_idx in formula_results:
                                    # Berikan nilai langsung daripada formula
                                    target_cell.value = formula_results[col_idx]
                                    print(f"Menggunakan hasil evaluasi {formula_results[col_idx]} untuk kolom {col_idx} daripada formula")
                                else:
                                    # Gunakan formula jika tidak ada hasil evaluasi
                                    target_cell.value = formula
                                    print(f"Menetapkan formula format lama ke baris {target_row}, kolom {col_idx}: {formula}")
                                
                        except Exception as e:
                            print(f"Error applying formula to row {target_row}, column {col_idx}: {e}")
                            # Jika gagal menerapkan formula, gunakan nilai biasa
                            # Pastikan tidak ada #N/A dengan memastikan nilai tidak kosong
                            if value and value.strip():
                                target_cell.value = value
                            else:
                                target_cell.value = 0  # Default ke 0 jika kosong
                    else:
                        # Periksa kategori kolom untuk pemrosesan khusus
                        is_psikogram_column = col_name in potensi_intelektual_columns or col_name in kepribadian_columns
                        
                        # Handle nilai khusus untuk mencegah #N/A
                        if not value or value.strip() == "" or value.strip().lower() == "#n/a":
                            if col_name.startswith("Unnamed:"):
                                # Untuk kolom unnamed, gunakan nilai asli dari input jika ada
                                # Jika tidak ada, gunakan 0 sebagai default
                                try:
                                    # Coba konversi ke nilai numerik jika mungkin
                                    if value and value.strip():
                                        numeric_value = float(value)
                                        if numeric_value == int(numeric_value):
                                            target_cell.value = int(numeric_value)
                                        else:
                                            target_cell.value = numeric_value
                                    else:
                                        target_cell.value = 0
                                except (ValueError, TypeError):
                                    # Jika tidak bisa dikonversi ke angka, gunakan nilai asli
                                    target_cell.value = value if value else 0
                            elif is_psikogram_column:
                                # Untuk kolom psikogram yang kosong, gunakan string kosong
                                target_cell.value = ""
                                print(f"Menyimpan nilai kosong untuk psikogram di kolom {col_idx} ('{col_name}')")
                            else:
                                # Untuk kolom lain, kosongkan saja
                                target_cell.value = ""
                        else:
                            # Kolom Unnamed yang sudah memiliki nilai
                            if col_name.startswith("Unnamed:") or col_idx in [14, 15]:
                                try:
                                    # Coba konversi ke angka jika mungkin
                                    numeric_value = float(value)
                                    if numeric_value == int(numeric_value):
                                        target_cell.value = int(numeric_value)
                                    else:
                                        target_cell.value = numeric_value
                                except (ValueError, TypeError):
                                    # Jika tidak bisa dikonversi, gunakan nilai asli
                                    target_cell.value = value
                            # Kolom Psikogram
                            elif is_psikogram_column:
                                # Cek apakah ada hasil formula yang sudah dievaluasi
                                if col_idx in formula_results:
                                    # Kolom 47 (Daya Ingat/ME), 50-54 dan 59-66 (kolom referensi ke Sheet2) harus selalu menggunakan formula, bukan hasil evaluasi
                                    if (col_idx == 47 or (col_idx >= 50 and col_idx <= 66)) and col_idx in formula_cells:
                                        # Ambil formula dari formula_cells
                                        try:
                                            formula_info = formula_cells[col_idx]
                                            if isinstance(formula_info, dict):
                                                if formula_info.get('type') == 'IFS':
                                                    template = formula_info.get('template')
                                                    
                                                    # Ganti {row} dengan nomor baris saat ini
                                                    formula = template.replace("{row}", str(target_row))
                                                    
                                                    # Perbaiki format formula IFS (hilangkan _xlfn. atau @ jika perlu)
                                                    if '_xlfn.IFS' in formula:
                                                        formula = formula.replace('_xlfn.IFS', 'IFS')
                                                    
                                                    # Jika ada tanda @ di awal formula, hapus
                                                    if formula.startswith('=@'):
                                                        formula = '=' + formula[2:]
                                                    
                                                    target_cell.value = formula
                                                    print(f"Menetapkan formula IFS ke baris {target_row}, kolom {col_idx}: {formula}")
                                                elif formula_info.get('type') == 'normal':
                                                    formula = formula_info.get('formula', '')
                                                    # Perbarui referensi baris
                                                    def update_cell_reference(match):
                                                        ref = match.group(0)
                                                        col_part = ''.join(c for c in ref if c.isalpha())
                                                        row_part = ''.join(c for c in ref if c.isdigit())
                                                        if row_part == str(header_row + 1):
                                                            return col_part + str(target_row)
                                                        return ref
                                                    
                                                    formula = re.sub(r'[A-Z]+\d+', update_cell_reference, formula)
                                                    target_cell.value = formula
                                                    print(f"Menetapkan formula normal ke baris {target_row}, kolom {col_idx}: {formula}")
                                            else:
                                                # Jika bukan formula terstruktur, gunakan nilai default
                                                target_cell.value = value.strip() or "C"  # Default ke "C" jika kosong
                                                print(f"Menggunakan nilai default untuk kolom {col_idx}: {target_cell.value}")
                                        except Exception as e:
                                            print(f"Error applying formula for column {col_idx}: {e}")
                                            target_cell.value = value.strip() or formula_results.get(col_idx, "")
                                    else:
                                        target_cell.value = formula_results[col_idx]
                                        print(f"Menggunakan hasil evaluasi {formula_results[col_idx]} untuk kolom {col_idx}")
                                else:
                                    # Jika tidak ada hasil evaluasi, gunakan nilai inputan
                                    target_cell.value = value.strip()
                                    print(f"Menyimpan nilai psikogram '{value.strip()}' ke kolom {col_idx} ('{col_name}')")
                            # Konversi nilai ke integer jika memungkinkan
                            elif value and value.replace('.', '', 1).replace('-', '', 1).isdigit():
                                try:
                                    # Coba konversi ke integer jika nilai merupakan angka bulat
                                    if float(value).is_integer():
                                        target_cell.value = int(float(value))
                                    else:
                                        # Jika ada desimal, tetap gunakan float
                                        target_cell.value = float(value)
                                except ValueError:
                                    target_cell.value = value
                            else:
                                target_cell.value = value
            
            # Mengembalikan merged cells
            # Perlu disesuaikan dengan jumlah baris baru
            row_offset = len(data_rows_to_delete) - len(ui_data)
            successfully_merged = 0
            for original_range in merged_ranges:
                try:
                    # Cek apakah range berada di atas header (tetap tidak berubah)
                    if original_range.min_row <= header_row:
                        sheet.merge_cells(str(original_range))
                        successfully_merged += 1
                    else:
                        # Sesuaikan range untuk baris di bawah header
                        # Gunakan offset antara data lama dan baru
                        if row_offset != 0:
                            new_min_row = original_range.min_row - row_offset
                            new_max_row = original_range.max_row - row_offset
                            
                            # Pastikan baris tetap berada dalam rentang yang valid
                            if new_max_row > header_row:
                                min_col_letter = get_column_letter(original_range.min_col)
                                max_col_letter = get_column_letter(original_range.max_col)
                                new_range = f"{min_col_letter}{new_min_row}:{max_col_letter}{new_max_row}"
                                
                                try:
                                    sheet.merge_cells(new_range)
                                    successfully_merged += 1
                                    print(f"Merged cells: {new_range}")
                                except Exception as e:
                                    print(f"Gagal merge cells {new_range}: {e}")
                except Exception as e:
                    print(f"Error processing merged range: {e}")
                    
            print(f"Berhasil mengembalikan {successfully_merged} dari {len(merged_ranges)} merged cells")
            
            # Sesuaikan lebar kolom agar sama dengan file asli
            for i, column in enumerate(sheet.columns, 1):
                try:
                    # Simpan lebar kolom asli
                    col_letter = get_column_letter(i)
                    if col_letter in sheet.column_dimensions:
                        width = sheet.column_dimensions[col_letter].width
                        if width is not None:
                            sheet.column_dimensions[col_letter].width = width
                except Exception as e:
                    print(f"Error setting column width for column {i}: {e}")
            
            # Sesuaikan tinggi baris
            for row_idx, height in row_heights.items():
                try:
                    if row_idx <= header_row:  # Baris header ke atas
                        sheet.row_dimensions[row_idx].height = height
                    else:  # Baris data
                        # Untuk baris data, gunakan tinggi default yang sama
                        for data_row_idx in range(header_row + 1, header_row + 1 + len(ui_data)):
                            sheet.row_dimensions[data_row_idx].height = default_row_height
                except Exception as e:
                    print(f"Error setting row height for row {row_idx}: {e}")
            
            # Buat temporary file untuk save dan copy
            temp_file = os.path.join(dir_path, f"temp_{timestamp}.xlsx")
            
            # Simpan ke file sementara dulu
            try:
                wb.save(temp_file)
                wb.close()
                print(f"Berhasil menyimpan ke file sementara: {temp_file}")
                
                # Tutup workbook asli jika masih terbuka
                import gc
                gc.collect()
                
                # Salin file sementara ke file asli
                shutil.copy2(temp_file, original_file)
                
                # Hapus file sementara
                try:
                    os.remove(temp_file)
                except:
                    pass
                
                print(f"Berhasil menyimpan ke file asli: {original_file}")
                
                # Tampilkan dialog informasi setelah berhasil menyimpan
                QMessageBox.information(self, "Sukses", f"Data berhasil disimpan ke {original_file}")
            except Exception as save_error:
                print(f"Error saving to temp file: {save_error}")
                raise save_error
            
        except Exception as e:
            print(f"Error saving to Excel: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Error", f"Gagal menyimpan file Excel: {e}")
            
            # Jika terjadi error, coba kembalikan dari backup
            try:
                latest_backup = max(
                    [f for f in os.listdir(backup_dir) if f.startswith("backup_") and f.endswith(os.path.basename(self.excel_file_path))],
                    key=lambda x: os.path.getmtime(os.path.join(backup_dir, x))
                )
                backup_path = os.path.join(backup_dir, latest_backup)
                restore_path = self.excel_file_path
                
                # Coba kembalikan file
                shutil.copy2(backup_path, restore_path)
                QMessageBox.information(self, "Pemulihan", f"File dikembalikan dari backup: {latest_backup}")
            except Exception as restore_error:
                print(f"Error restoring from backup: {restore_error}")

    def get_column_index(self, column_name):
        # Search for exact match first
        for i in range(self.table.columnCount()):
            header = self.table.horizontalHeaderItem(i)
            if header and header.text() == column_name:
                return i
                
        # If no exact match, try partial match
        for i in range(self.table.columnCount()):
            header = self.table.horizontalHeaderItem(i)
            if header and column_name.strip() in header.text():
                return i
                
        # If no match found, return -1
        print(f"Column '{column_name}' not found")
        return -1

    def preview_pdf(self):
        try:
            # Check if Excel file has been loaded
            if not hasattr(self, 'excel_file_path') or not self.excel_file_path:
                QMessageBox.warning(self, "Warning", "Please load an Excel file first!")
                return

            # Check if a row is selected
            selected_row = self.table.currentRow()
            if selected_row < 0:
                QMessageBox.warning(self, "Warning", "Please select a row to preview first!")
                return

            # Get column indices and data
            iq_col = self.get_column_index("IQ ")
            nama_col = self.get_column_index("Nama Peserta")
            tgl_lahir_col = self.get_column_index("TGL Lahir")

            # Get data from selected row
            iq_val = self.table.item(selected_row, iq_col)
            nama_val = self.table.item(selected_row, nama_col)
            tgl_lahir_val = self.table.item(selected_row, tgl_lahir_col)

            # Initialize row_data from the selected row
            row_data = {}
            for col, column_name in enumerate(self.columns):
                item = self.table.item(selected_row, col)
                row_data[column_name] = item.text() if item else ""

            # Debug prints
            print(f"Selected Row: {selected_row}")
            for key, value in row_data.items():
                print(f"{key}: {value}")

            nama = nama_val.text() if nama_val else ""
            tgl_lahir = tgl_lahir_val.text() if tgl_lahir_val else ""

            # Ensure iq_val is converted to float correctly
            try:
                iq_value_numeric = float(iq_val.text()) if iq_val and iq_val.text().strip() else 0.0
            except ValueError:
                iq_value_numeric = 0.0

            # For debugging
            print(f"IQ Value from table: {iq_value_numeric}")

            # Create HTML content
            html_content = """
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    @page {
                        size: A4;
                        margin: 1cm;
                    }
                    body { 
                        font-family: Arial, sans-serif;
                        padding: 20px;
                        width: 21cm;
                        min-height: 29.7cm;
                        margin: 0 auto;
                        background: white;
                    }
                    @media print {
                        body {
                            width: auto;
                            height: auto;
                            margin: 0;
                            padding: 0;
                        }
                        .page {
                            width: 21cm;
                            min-height: 29.7cm;
                            padding: 1cm;
                            margin: 0;
                            page-break-after: always;
                        }
                        .page:last-child {
                            page-break-after: avoid;
                        }
                        .page-break {
                            page-break-before: always;
                            margin: 0;
                            padding: 0;
                        }
                    }
                    .header { text-align: center; margin-bottom: 20px; }
                    .title { font-size: 16px; font-weight: bold; margin: 10px 0; }
                    .info-table { width: 100%; margin-bottom: 15px; border-spacing: 0; table-layout:fixed;  }
                    .info-table td { padding: 3px; vertical-align: top; width: 25%;  }
                    .main-table, .psikogram { 
                        width: 100%; 
                        border-collapse: collapse; 
                        margin-bottom: 20px;
                    }
                    .main-table th, .main-table td, .psikogram th, .psikogram td { 
                        border: 1px solid black; 
                        padding: 8px; 
                        text-align: center;
                    }
                    .psikogram td { text-align: left; }
                    .main-table th, .psikogram th { background-color: #f2f2f2; }
                    .center-text { text-align: center; }
                    .footer { text-align: center; font-style: italic; margin-top: 20px; }
                    .psikogram th, .psikogram td { 
                        padding: 10px; 
                        font-size: 12px; 
                        vertical-align: middle;
                    }
                </style>
            </head>
            """

            # Add personal info
            # Dapatkan tanggal saat ini
            tanggal_tes = datetime.now().strftime("%d %B %Y")

            # Konversi dan format tanggal lahir
            try:
                # Coba parse dengan format DD/MM/YYYY
                tgl_lahir_obj = datetime.strptime(tgl_lahir, "%d/%m/%Y")
            except ValueError:
                try:
                    # Jika gagal, coba parse dengan format YYYY-MM-DD
                    tgl_lahir_obj = datetime.strptime(tgl_lahir, "%Y-%m-%d")
                except ValueError:
                    # Jika masih gagal, gunakan tanggal hari ini sebagai fallback
                    print(f"Error: Format tanggal '{tgl_lahir}' tidak dikenali. Menggunakan tanggal saat ini.")
                    tgl_lahir_obj = datetime.now()
                    
            tgl_lahir_formatted = tgl_lahir_obj.strftime("%d %B %Y")

            html_content += f"""
            <div style="width: 100%; margin: 0 auto;">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                    <div style="width: 80px;">
                        <img src="E:/Ney/Kuliah/Project/Surveyor/logo1.jpg" alt="Logo" style="width: 100%; height: auto;">
                        <div style="color: #1f4e79; font-weight: bold; text-align: center; font-size: 14px; margin-top: 5px;">BEHAVYOURS</div>
                    </div>
                    <div style="text-align: center; flex-grow: 1;">
                        <div style="font-size: 14px; font-weight: bold; color: #1f4e79;">HASIL PEMERIKSAAN PSIKOLOGIS</div>
                        <div style="font-size: 12px; color: #1f4e79;">(Asesmen Intelegensi, Kepribadian dan Minat)</div>
                    </div>
                    <div style="text-align: right; font-size: 12px;">
                        <div style="font-weight: bold; color: #1f4e79;">RAHASIA</div>
                        <div style="color: #1f4e79;">No. {row_data.get('No', '')} / {row_data.get('No Tes', '')}</div>
                    </div>
                </div>
                <table class="info-table" style="margin-bottom: 20px; border-spacing: 0; font-size: 12px;">
                    <tr>
                        <td width="15%" style="padding: 4px 0; color: #c45911; font-weight: bold;">NAMA</td>
                        <td width="35%" style="padding: 4px 0; color: #c45911; font-weight: bold;">: {nama}</td>
                        <td width="15%" style="padding: 4px 0; color: #c45911; font-weight: bold;">PERUSAHAAN</td>
                        <td width="35%" style="padding: 4px 0; color: #c45911; font-weight: bold;">: PT. BAM</td>
                    </tr>
                    <tr>
                        <td style="padding: 4px 0; color: #c45911; font-weight: bold;">TANGGAL LAHIR</td>
                        <td style="padding: 4px 0; color: #c45911; font-weight: bold;">: {tgl_lahir_formatted}</td>
                        <td style="padding: 4px 0; color: #c45911; font-weight: bold;">TANGGAL TES</td>
                        <td style="padding: 4px 0; color: #c45911; font-weight: bold;">: {tanggal_tes}</td>
                    </tr>
                    <tr>
                        <td style="padding: 4px 0; color: #c45911; font-weight: bold;">PEMERIKSA</td>
                        <td style="padding: 4px 0; color: #c45911; font-weight: bold;">: Chitra Ananda Mulia, M.Psi., Psikolog</td>
                        <td style="padding: 4px 0; color: #c45911; font-weight: bold;">LEMBAGA</td>
                        <td style="padding: 4px 0; color: #c45911; font-weight: bold;">: BEHAVYOURS</td>
                    </tr>
                    <tr>
                        <td style="padding: 4px 0; color: #c45911; font-weight: bold;">ALAMAT LEMBAGA</td>
                        <td colspan="3" style="padding: 4px 0; color: #c45911; font-weight: bold;">: Jl. Patal Senayan No.01</td>
                    </tr>
                </table>
            </div>
            """
            # Pastikan iq_value adalah numerik
            try:
                iq_value_numeric = float(iq_val.text())
            except ValueError:
                iq_value_numeric = 0 
            # Add IQ Classification table
            html_content += f"""
                    <div style="width: 100%; margin: 0 auto;">
                        <table style="width: 100%; border-collapse: separate; border-spacing: 0 0;">
                            <tr>
                                <td style="width: 25%; padding-right: 15px; vertical-align: top;">
                                    <table style="width: 100%; border-collapse: collapse; border: 1px solid black;">
                                        <tr>
                                            <th style="border-bottom: 1px solid black; padding: 8px; text-align: center; background-color: #f7caac;">KECERDASAN UMUM</th>
                                        </tr>
                                        <tr>
                                            <td>
                                                <table style="width: 100%; border-collapse: collapse;">
                                                    <tr>
                                                        <td style="border-right: 1px solid black; padding: 8px; text-align: left;">Taraf<br>Kecerdasan<br>IQ</td>
                                                        <td style="padding: 8px; text-align: center;">{iq_val.text()}</td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                                <td style="width: 75%; vertical-align: top;">
                                    <table style="width: 100%; border-collapse: collapse;">
                                        <tr>
                                            <th colspan="5" style="border: 1px solid black; padding: 4px; text-align: center; background-color: #ffceb4;">KLASIFIKASI KECERDASAN IQ</th>
                                        </tr>
                                        <tr>
                                            <td style="border: 1px solid black; padding: 4px; text-align: center;">Rendah</td>
                                            <td style="border: 1px solid black; padding: 4px; text-align: center;">Dibawah Rata-Rata</td>
                                            <td style="border: 1px solid black; padding: 4px; text-align: center;">Rata-Rata</td>
                                            <td style="border: 1px solid black; padding: 4px; text-align: center;">Diatas Rata-Rata</td>
                                            <td style="border: 1px solid black; padding: 4px; text-align: center;">Superior</td>
                                        </tr>
                                        <tr>
                                            <td style="border: 1px solid black; padding: 4px; text-align: center;">&lt; 79</td>
                                            <td style="border: 1px solid black; padding: 4px; text-align: center;">80 - 89</td>
                                            <td style="border: 1px solid black; padding: 4px; text-align: center;">90 - 109</td>
                                            <td style="border: 1px solid black; padding: 4px; text-align: center;">110 - 119</td>
                                            <td style="border: 1px solid black; padding: 4px; text-align: center;">&gt; 120</td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>   
                    </div>     
                        """

            # Add IQ data
            iq_val = self.table.item(selected_row, self.get_column_index("IQ "))
            iq_class = self.table.item(selected_row, self.get_column_index("KLASIFIKASI"))
            if iq_val and iq_class:

                # Define a function to determine the position of 'X'
                def get_x_position(value):
                    positions = {
                        "R": ["X", "", "", "", ""],
                        "K": ["", "X", "", "", ""],
                        "C": ["", "", "X", "", ""],
                        "B": ["", "", "", "X", ""],
                        "T": ["", "", "", "", "X"]
                    }
                    return positions.get(value, ["", "", "", "", ""])

                # Get values for each psychological aspect
                aspects = {
                    "daya_analisa": row_data.get("Daya Analisa/ AN", ""),
                    "kemampuan_numerik": row_data.get("Kemampuan Numerik/ RA ZR", ""),
                    "kemampuan_verbal": row_data.get("Kemampuan Verbal/WA GE", ""),
                    "fleksibilitas": row_data.get("Flexibilitas/ T V", ""),
                    "sistematika_kerja": row_data.get("Sistematika Kerja/ cd", ""),
                    "inisiatif": row_data.get("Inisiatif/W", ""),
                    "stabilitas_emosi": row_data.get("Stabilitas Emosi / E", ""),
                    "komunikasi": row_data.get("Komunikasi / B O", ""),
                    "keterampilan_interpersonal": row_data.get("Keterampilan Interpersonal / S O", ""),
                    "kerjasama": row_data.get("Kerjasama / B X", "")
                }

                # Define common cell styles
                header_style = 'text-align: center; padding: 8px; background-color: #deeaf6; border: 1px solid black;'
                cell_style = 'text-align: center; font-weight: bold;'
                section_header_style = 'background-color: #fbe4d5; text-align: center; font-weight: bold;'
                
                # Add psikogram table with dynamic 'X' positions
                html_content += f"""
                    <table class="psikogram" style="width: 100%; margin-top: 20px; border-collapse: collapse; border: 1px solid black;">
                        <tr>
                            <th colspan="8" style="{header_style}">PSIKOGRAM</th>
                        </tr>
                        <tr>
                            <th style="width: 5%; {header_style}">NO</th>
                            <th style="width: 15%; {header_style}">ASPEK<br>PSIKOLOGIS</th>
                            <th style="width: 40%; {header_style}">DEFINISI</th>
                            <th style="width: 8%; {header_style}">R</th>
                            <th style="width: 8%; {header_style}">K</th>
                            <th style="width: 8%; {header_style}">C</th>
                            <th style="width: 8%; {header_style}">B</th>
                            <th style="width: 8%; {header_style}">T</th>
                        </tr>

                        <tr><td colspan="8" style="{section_header_style}">KEMAMPUAN INTELEKTUAL</td></tr>
                        
                        <!-- Logika Berpikir -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">1.</td>
                            <td style="font-weight: bold;">Logika Berpikir</td>
                            <td>Kemampuan untuk berpikir secara logis dan sistematis.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["daya_analisa"]))}
                        </tr>

                        <!-- Daya Analisa -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">2.</td>
                            <td style="font-weight: bold;">Daya Analisa</td>
                            <td>Kemampuan untuk melihat permasalahan dan memahami hubungan sebab akibat permasalahan.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["daya_analisa"]))}
                        </tr>

                        <!-- Kemampuan Numerikal -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">3.</td>
                            <td style="font-weight: bold;">Kemampuan Numerikal</td>
                            <td>Kemampuan untuk berpikir praktis dalam memahami konsep angka dan hitungan.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["kemampuan_numerik"]))}
                        </tr>

                        <!-- Kemampuan Verbal -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">4.</td>
                            <td style="font-weight: bold;">Kemampuan Verbal</td>
                            <td>Kemampuan untuk memahami konsep dan pola dalam bentuk kata dan mengekspresikan gagasan secara verbal.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["kemampuan_verbal"]))}
                        </tr>

                        <tr><td colspan="8" style="{section_header_style}">SIKAP DAN CARA KERJA</td></tr>

                        <!-- Orientasi Hasil -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">5.</td>
                            <td style="font-weight: bold;">Orientasi Hasil</td>
                            <td>Kemampuan untuk mempertahankan komitmen untuk menyelesaikan tugas secara bertanggung jawab dan memperhatikan keterhubungan antara perencanaan dan hasil kerja.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["sistematika_kerja"]))}
                        </tr>

                        <!-- Fleksibilitas -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">6.</td>
                            <td style="font-weight: bold;">Fleksibilitas</td>
                            <td>Kemampuan untuk menyesuaikan diri dalam menghadapi permasalahan.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["fleksibilitas"]))}
                        </tr>

                        <!-- Sistematika Kerja -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">7.</td>
                            <td style="font-weight: bold;">Sistematika Kerja</td>
                            <td>Kemampuan untuk merencanakan hingga mengorganisasikan cara kerja dalam proses penyelesaian pekerjaannya.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["sistematika_kerja"]))}
                        </tr>

                        <tr><td colspan="8" style="{section_header_style}">KEPRIBADIAN</td></tr>

                        <!-- Motivasi Berprestasi -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">8.</td>
                            <td style="font-weight: bold;">Motivasi Berprestasi</td>
                            <td>Kemampuan untuk menunjukkan prestasi dan mencapai target.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["inisiatif"]))}
                        </tr>

                        <!-- Kerjasama -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">9.</td>
                            <td style="font-weight: bold;">Kerjasama</td>
                            <td>Kemampuan untuk menjalin, membina dan mengoptimalkan hubungan kerja yang efektif demi tercapainya tujuan bersama.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["kerjasama"]))}
                        </tr>

                        <!-- Keterampilan Interpersonal -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">10.</td>
                            <td style="font-weight: bold;">Keterampilan Interpersonal</td>
                            <td>Kemampuan untuk menjalin hubungan sosial dan mampu memahami kebutuhan orang lain.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["keterampilan_interpersonal"]))}
                        </tr>

                        <!-- Stabilitas Emosi -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">11.</td>
                            <td style="font-weight: bold;">Stabilitas Emosi</td>
                            <td>Kemampuan untuk memahami dan mengontrol emosi.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["stabilitas_emosi"]))}
                        </tr>

                        <tr><td colspan="8" style="{section_header_style}">KEMAMPUAN BELAJAR</td></tr>

                        <!-- Pengembangan Diri -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">12.</td>
                            <td style="font-weight: bold;">Pengembangan Diri</td>
                            <td>Kemampuan untuk meningkatkan pengetahuan dan menyempurnakan keterampilan diri.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["inisiatif"]))}
                        </tr>

                        <!-- Mengelola Perubahan -->
                        <tr>
                            <td style="text-align: center; background-color: #deeaf6; font-weight: bold;">13.</td>
                            <td style="font-weight: bold;">Mengelola Perubahan</td>
                            <td>Kemampuan dalam menyesuaikan diri dengan situasi yang baru.</td>
                            {' '.join(f'<td style="{cell_style}">{x}</td>' for x in get_x_position(aspects["fleksibilitas"]))}
                        </tr>

                        <!-- Legend -->
                        <tr style="border-top: 1px solid black;">
                            <td colspan="8" style="text-align: center; padding: 2px; font-family: Arial; font-size: 11px; background-color: #deeaf6;">
                                <div style="display: inline-block; width: 100%; font-weight: bold;">
                                    T : Tinggi&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                    B : Baik&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                    C : Cukup&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                    K : Kurang&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                    R : Rendah
                                </div>
                            </td>
                        </tr>
                    </table>
                """

            # Add page break and second page content
            html_content += f"""
                <div class="page-break"></div>
                <div class="page" style="padding: 1cm; font-family: Arial;">
                    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px;">
                        <div style="width: 80px;">
                            <img src="E:/Ney/Kuliah/Project/Surveyor/logo1.jpg" alt="Logo" style="width: 100%; height: auto;">
                            <div style="color: #1f4e79; font-weight: bold; text-align: center; font-size: 14px; margin-top: 5px;">BEHAVYOURS</div>
                        </div>
                        <div style="text-align: center; flex-grow: 1;">
                            <div style="font-size: 14px; font-weight: bold; color: #1f4e79;">HASIL PEMERIKSAAN PSIKOLOGIS</div>
                            <div style="font-size: 12px; color: #1f4e79;">(Asesmen Intelegensi, Kepribadian dan Minat)</div>
                        </div>
                        <div style="text-align: right; font-size: 12px;">
                            <div style="font-weight: bold; color: #1f4e79;">RAHASIA</div>
                            <div style="color: #1f4e79;">No. {row_data.get('No', '')} / {row_data.get('No Tes', '')}</div>
                        </div>
                    </div>

                    <table class="psikogram" style="width: 100%; border-collapse: collapse; margin-top: 20px; font-family: Arial, sans-serif;">
                        <tr>
                            <th colspan="2" style="text-align: center; padding: 8px; background-color: #fbe4d5; border: 1px solid black;">KESIMPULAN</th>
                        </tr>
                        <tr>
                            <td style="width: 20%; padding: 8px; vertical-align: top; border: 1px solid black; font-weight: bold;">KEMAMPUAN INTELEKTUAL</td>
                            <td style="width: 80%; padding: 8px; text-align: justify; border: 1px solid black;">
                                Berdasarkan pemeriksaan kemampuan intelektual, diketahui bahwa Sdr. {nama} menunjukkan kapasitas intelektual yang cukup memadai. Ia cukup mampu untuk menganalisis dan membuat kesimpulan yang tidak sepenuhnya didukung oleh bukti. Ia menunjukkan ketelitian dalam mengidentifikasi komponen-komponen penting dari suatu masalah, sehingga pemahaman terhadap hubungan sebab-akibat menjadi terbatas. Selain itu, menunjukkan pemahaman yang kurang memadai terhadap konsep matematis dasar, yang mempengaruhi kemampuan analisa. Namun, menunjukkan pemahaman yang baik terhadap makna makna dalam bahasa, yang mempengaruhi kemampuan ekspresi dalam berkomunikasi secara efektif.
                            </td>
                        </tr>
                        <tr>
                            <td style="padding: 8px; vertical-align: top; border: 1px solid black; font-weight: bold;">SIKAP DAN CARA KERJA</td>
                            <td style="padding: 8px; text-align: justify; border: 1px solid black;">
                                Berdasarkan pemeriksaan sikap dan cara kerja, diketahui bahwa Sdr. {nama} mampu menyelesaikan tugas-tugas yang diberikan dengan cukup baik meskipun terkadang membutuhkan waktu yang lebih lama. Ia menunjukkan kesulitan dalam beradaptasi dan menyesuaikan diri dengan perubahan, terkadang merasa tidak nyaman dengan hal-hal baru. Kemudian, mampu membuat rencana kerja yang cukup terstruktur, meskipun terkadang membutuhkan pengawasan tambahan.
                            </td>
                        </tr>
                        <tr>
                            <td style="padding: 8px; vertical-align: top; border: 1px solid black; font-weight: bold;">KEPRIBADIAN</td>
                            <td style="padding: 8px; text-align: justify; border: 1px solid black;">
                                Berdasarkan pemeriksaan kepribadian, diketahui bahwa Sdr. {nama} menunjukkan motivasi yang kuat untuk mencapai target yang ditetapkan, selalu berusaha untuk mencapai ekspektasi. Tak hanya itu, juga menunjukkan usaha yang kurang maksimal dalam berkontribusi pada kelompok, terkadang mengalami tanggung jawab yang diberikan. Ia mampu membina dan mempertahankan hubungan sosial yang cukup baik, meskipun terkadang dalam situasi tertentu. Selain itu, menunjukkan emosi yang cukup stabil dalam menghadapi situasi yang menantang, meskipun terkadang membutuhkan waktu untuk beradaptasi dengan perubahan.
                            </td>
                        </tr>
                        <tr>
                            <td style="padding: 8px; vertical-align: top; border: 1px solid black; font-weight: bold;">KEMAMPUAN BELAJAR</td>
                            <td style="padding: 8px; text-align: justify; border: 1px solid black;">
                                Berdasarkan pemeriksaan kemampuan belajar, diketahui bahwa Sdr. {nama} menunjukkan inisiatif yang kuat dalam memiliki pengetahuan dan keterampilan diri, giat dalam mencapai hal-hal baru dan berusaha untuk meningkatkan pengetahuan dan keterampilan diri. Namun, terkadang membutuhkan waktu untuk beradaptasi dengan perubahan, yang menunjukkan perlunya perhatian khusus dalam beradaptasi.
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" style="padding: 8px;">
                                <div style="font-weight: bold; margin: 10px 0; font-size: 12px;">PENGEMBANGAN</div>
                                <div style="text-align: justify; border: 1px solid black; padding: 8px;">
                                    Sdr. {nama} masih membutuhkan pengembangan dalam mengidentifikasi pola logis, menarik kesimpulan, dan memperdalami penguasaan hubungan sebab akibat menjadi lebih baik dalam konsep matematis dan butuh memperdalam penggunaan bahasa lebih lanjut. Kemudian, butuh kefokusan agar lebih mudah dalam beradaptasi, dan butuh kontribusi lebih, dan koordinasi dengan kelompok agar mencapai tujuan bersama. Serta mudah diri dengan pada pada hal-hal baru.
                                </div>
                            </td>
                        </tr>
                    </table>

                    <table class="psikogram" style="width: 100%; border-collapse: collapse; margin-top: 20px; font-family: Arial, sans-serif;">
                        <tr>
                            <th colspan="2" style="text-align: center; padding: 8px; background-color: #fbe4d5; border: 1px solid black;">Kategori Hasil Screening</th>
                        </tr>
                        <tr>
                            <td style="width: 5%; text-align: center; border: 1px solid black; padding: 8px;">X</td>
                            <td style="padding: 8px; border: 1px solid black;">Tahapan Normal<br><span style="font-size: 10px; color: #666;">Individu menunjukkan adaptasi gejala gangguan mental yang mengganggu fungsi sehari-hari</span></td>
                        </tr>
                        <tr>
                            <td style="text-align: center; border: 1px solid black; padding: 8px;"></td>
                            <td style="padding: 8px; border: 1px solid black;">Kecenderungan Stress dalam Tekanan<br><span style="font-size: 10px; color: #666;">Dalam situasi yg menimbulkan tekanan dapat berdampak pada kondisi individu & respon emosional yg ditampilkan</span></td>
                        </tr>
                        <tr>
                            <td style="text-align: center; border: 1px solid black; padding: 8px;"></td>
                            <td style="padding: 8px; border: 1px solid black;">Gangguan<br><span style="font-size: 10px; color: #666;">Individu menunjukkan gejala-gejala gangguan yang dapat mengganggu fungsi sehari-hari</span></td>
                        </tr>
                    </table>
                    
                    <table class="psikogram" style="width: 100%; border-collapse: collapse; margin-top: 20px; font-family: Arial, sans-serif;">
                        <tr>
                            <th colspan="2" style="text-align: center; padding: 8px; background-color: #fbe4d5; border: 1px solid black;">Kesimpulan Keseluruhan</th>
                        </tr>
                        <tr>
                            <td style="width: 8%; text-align: center; border: 1px solid black; padding: 8px;"></td>
                            <td style="padding: 8px; border: 1px solid black;">LAYAK DIREKOMENDASIKAN</td>
                        </tr>
                        <tr>
                            <td style="width: 8%; text-align: center; border: 1px solid black; padding: 8px;">X</td>
                            <td style="padding: 8px; border: 1px solid black;">LAYAK DIPERTIMBANGKAN</td>
                        </tr>
                        <tr>
                            <td style="text-align: center; border: 1px solid black; padding: 8px;"></td>
                            <td style="padding: 8px; border: 1px solid black;">TIDAK DISARANKAN</td>
                        </tr>
                    </table>    
                </div>
            """

            # Add page break and third page content
            html_content += f"""
                <div class="page-break"></div>
                <div class="page" style="padding: 1cm; font-family: Arial;">
                    <div style="display: flex; align-items: center; margin-bottom: 20px;">
                        <img src="behanyours.png" alt="Logo" style="width: 80px; height: auto; margin-right: 20px;">
                        <div style="flex-grow: 1; text-align: center;">
                            <div style="font-size: 14px; font-weight: bold; color: #1f4e79;">HASIL PEMERIKSAAN PSIKOLOGIS</div>
                            <div style="font-size: 12px; color: #1f4e79;">(Asesmen Intelegensi, Kepribadian dan Minat)</div>
                        </div>
                        <div style="text-align: right; font-size: 12px;">
                            <div style="font-weight: bold; color: #1f4e79;">RAHASIA</div>
                            <div style="color: #1f4e79;">No. {row_data.get('No', '')} / {row_data.get('No Tes', '')}</div>
                        </div>
                    </div>

                    <div style="margin-bottom: 20px;">
                        <div style="margin-bottom: 15px;">
                            <div>
                                <span style="display: inline-block; width: 120px;">Tanggal</span>
                                <span>: 26 Februari 2025</span>
                            </div>
                            <div style="font-style: italic; font-size: 11px; color: #666;">Date</div>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <div>
                                <span style="display: inline-block; width: 120px;">Tanda Tangan</span>
                            </div>
                            <div style="font-style: italic; font-size: 11px; color: #666;">Signature</div>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <div>
                                <span style="display: inline-block; width: 120px;">Nama Psikolog</span>
                                <span>: Chitra Ananda Mulia, M.Psi., Psikolog</span>
                            </div>
                            <div style="font-style: italic; font-size: 11px; color: #666;">Psychologist Name</div>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <div>
                                <span style="display: inline-block; width: 120px;">Nomor STR/SIK</span>
                                <span>:</span>
                            </div>
                            <div style="font-style: italic; font-size: 11px; color: #666;">Registration Number</div>
                        </div>
                        
                        <div style="margin-bottom: 15px;">
                            <div>
                                <span style="display: inline-block; width: 120px;">Nomor SIPP/SIPPK</span>
                                <span>: 1564-19-2-2</span>
                            </div>
                            <div style="font-style: italic; font-size: 11px; color: #666;">Licence Number</div>
                        </div>
                    </div>
                </div>
            """

            # Close tables and add footer
            html_content += """
                </table>
                <div class="footer">
                    Laporan ini bersifat confidential dan diketahui oleh Psikolog
                </div>
            </body>
            </html>
            """

            # Create and show preview dialog
            preview_dialog = QDialog(self)
            preview_dialog.setWindowTitle("Preview PDF")
            
            # Get screen size
            screen = QApplication.primaryScreen()
            screen_size = screen.availableGeometry()
            
            # Set dialog size to 90% of screen size
            dialog_width = int(screen_size.width() * 0.9)
            dialog_height = int(screen_size.height() * 0.9)
            preview_dialog.setFixedSize(dialog_width, dialog_height)
            
            # Center the dialog on screen
            preview_dialog.move(
                (screen_size.width() - dialog_width) // 2,
                (screen_size.height() - dialog_height) // 2
            )
            
            preview_dialog.setWindowFlags(Qt.Window | Qt.WindowMaximizeButtonHint | Qt.WindowCloseButtonHint)
            
            # Create main vertical layout
            main_layout = QVBoxLayout(preview_dialog)
            main_layout.setContentsMargins(10, 10, 10, 10)
            main_layout.setSpacing(10)
            
            # Create horizontal layout for preview pages
            preview_layout = QHBoxLayout()
            
            # Create web view for page 1
            web_view1 = QWebEngineView(preview_dialog)
            web_view1.setZoomFactor(0.6)
            web_view1.setFixedWidth(int(dialog_width * 0.3))  # Adjust width to 30% for 3 pages
            # Split HTML content at page break
            pages = html_content.split('<div class="page-break"></div>')
            web_view1.setHtml(pages[0])
            preview_layout.addWidget(web_view1)
            
            # Create web view for page 2
            web_view2 = QWebEngineView(preview_dialog)
            web_view2.setZoomFactor(0.6)
            web_view2.setFixedWidth(int(dialog_width * 0.3))  # Adjust width to 30% for 3 pages
            if len(pages) > 1:
                web_view2.setHtml(pages[1])
            preview_layout.addWidget(web_view2)

            # Create web view for page 3
            web_view3 = QWebEngineView(preview_dialog)
            web_view3.setZoomFactor(0.8)
            web_view3.setFixedWidth(int(dialog_width * 0.3))  # Adjust width to 30% for 3 pages
            if len(pages) > 2:
                web_view3.setHtml(pages[2])
            preview_layout.addWidget(web_view3)
            
            # Add preview layout to main layout
            main_layout.addLayout(preview_layout)
            
            # Create button layout for centering
            button_layout = QHBoxLayout()
            button_layout.addStretch()
            
            # Add save PDF button
            save_button = QPushButton("Save PDF", preview_dialog)
            save_button.setFixedHeight(30)
            save_button.setFixedWidth(200)  # Set fixed width for button
            save_button.clicked.connect(lambda: self.save_as_pdf(html_content))
            button_layout.addWidget(save_button)
            
            # Add direct print button without preview
            print_button = QPushButton("Print PDF", preview_dialog)
            print_button.setFixedHeight(30)
            print_button.setFixedWidth(200)
            print_button.clicked.connect(lambda: self.print_pdf(html_content))
            button_layout.addWidget(print_button)
            
            button_layout.addStretch()
            main_layout.addLayout(button_layout)
            
            preview_dialog.exec_()
            
        except Exception as e:
            print(f"Error saat preview: {e}")
            
    def print_pdf(self, html_content):
        try:
            # Create printer with A4 settings
            printer = QPrinter(QPrinter.HighResolution)
            printer.setPageSize(QPageSize(QPageSize.A4))
            # Set margins using individual float values
            printer.setPageMargins(10, 10, 10, 10, QPrinter.Millimeter)
            
            # Create web view and load content
            web_view = QWebEngineView()
            web_view.setHtml(html_content)
            
            # Wait for page load
            loop = QEventLoop()
            web_view.loadFinished.connect(loop.quit)
            loop.exec_()
            
            # Show print dialog
            print_dialog = QPrintDialog(printer, self)
            if print_dialog.exec_() == QPrintDialog.Accepted:
                def print_finished(success):
                    if success:
                        QMessageBox.information(self, "Success", "Document printed successfully")
                    else:
                        QMessageBox.warning(self, "Warning", "Print job failed")
                
                # Direct print without preview
                web_view.page().print(printer, print_finished)
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error printing document: {e}")
            print(f"Print error: {e}")

    def save_as_pdf(self, html_content):
        try:
            file_name, _ = QFileDialog.getSaveFileName(
                self,
                "Save PDF",
                "",
                "PDF Files (*.pdf)"
            )
            
            if file_name:
                # Create printer with A4 settings
                printer = QPrinter(QPrinter.HighResolution)
                printer.setOutputFormat(QPrinter.PdfFormat)
                printer.setOutputFileName(file_name)
                printer.setPageSize(QPageSize(QPageSize.A4))
                
                # Create web view with A4 dimensions
                web_view = QWebEngineView()
                
                # Update the styles to ensure content fits on one page
                html_content = html_content.replace('</head>',
                    '''
                    <style>
                        @page {
                            size: A4;
                            margin: 1cm;
                        }
                        @media print {
                            body {
                                width: 210mm;
                                height: 297mm;
                                margin: 0;
                                padding: 1cm;
                            }
                            .page {
                                page-break-after: always;
                            }
                            .page:last-child {
                                page-break-after: avoid;
                            }
                        }
                        body {
                            margin: 0;
                            padding: 1cm;
                            width: 210mm;
                            height: 297mm;
                            font-family: Arial, sans-serif;
                            font-size: 11px;
                            position: relative;
                        }
                        .header {
                            text-align: center;
                            margin-bottom: 15px;
                        }
                        .header .title {
                            font-size: 14px;
                            font-weight: bold;
                            margin-bottom: 5px;
                        }
                        .info-table {
                            width: 100%;
                            margin-bottom: 15px;
                            border-spacing: 0;
                        }
                        .info-table td {
                            padding: 3px;
                            vertical-align: top;
                        }
                        table {
                            border-collapse: collapse;
                            width: 100%;
                        }
                        .psikogram {
                            margin-top: 15px;
                        }
                        .psikogram th, .psikogram td {
                            border: 1px solid black;
                            padding: 4px;
                            font-size: 11px;
                        }
                        .psikogram th {
                            background-color: #f2f2f2;
                            text-align: center;
                        }
                        .category-header {
                            background-color: #f8d7da !important;
                            text-align: center;
                            font-weight: bold;
                        }
                        .footer {
                            position: fixed;
                            bottom: 1cm;
                            left: 1cm;
                            right: 1cm;
                            text-align: center;
                            font-style: italic;
                            font-size: 10px;
                            background: white;
                            padding: 5px;
                        }
                        .legend-row td {
                            text-align: center;
                            padding: 2px;
                            font-size: 11px;
                            border: none;
                        }
                        .page-footer {
                            position: fixed;
                            bottom: 1cm;
                            left: 1cm;
                            right: 1cm;
                            text-align: center;
                            font-style: italic;
                            font-size: 10px;
                            background: white;
                            padding: 5px;
                        }
                    </style>
                    </head>
                    ''')
                web_view.setHtml(html_content)
                
                # Wait for page to load
                loop = QEventLoop()
                web_view.loadFinished.connect(loop.quit)
                loop.exec_()
                
                # Print to PDF
                web_view.page().printToPdf(file_name)
                
                QMessageBox.information(self, "Success", "PDF saved successfully!")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error saving PDF: {e}")
            print(f"Error saving PDF: {e}")
    
    def save_pdf_file(self, pdf_data, file_name):
        try:
            with open(file_name, 'wb') as f:
                f.write(pdf_data)
        except Exception as e:
            print(f"Error writing PDF file: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelViewerApp()
    window.show()
    sys.exit(app.exec_())
