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
    Qt, QEventLoop, QSizeF, QMarginsF, QUrl
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
import base64

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
        # Updated columns based on the data from log
        self.columns = [
            "No", "No Tes", "Tgl Test", "TGL Lahir", "Nama PT", "JK", "SDR/SDRI", "Nama Peserta", 
            "PHQ", "Keterangan PHQ", "IQ ", "SE / Konkrit Praktis", "WA/ Verbal", " AN / Flexibilitas Pikir",
            "GE / Daya Abstraksi Verbal", "RA / Berpikir Praktis", "Unnamed: 16", "KLASIFIKASI",
            "N", "G", "A", "L", "P", "I", "T", "V", "S", "B", "O", "X", "C", "C (Coding)", "D", "R", "Z", "E", "K", "F", "W", 
            "NG", "CDR", "TV", "PI", "BS", "ZK",
            "Logika Berpikir 1", "Daya Analisa 3", "Kemampuan Verbal 2 dam 4", "Kemampuan Numerik 5", 
            "Sistematika Kerja/ C D R", "Orientasi Hasil/ N G", "Fleksibilitas/ T V", "Motivasi Berprestasi/ A", 
            "Kerjasama/ P I", "Keterampilan Interpersonal/ B S", "Stabilitas Emosi/ E PHQ", "Pegembangan Diri/ W", 
            "Mengelola Perubahan/ Z K",
            "Logika Berpikir 1.1", "Daya Analisa 3.1", "Kemampuan Verbal 2 dam 4.1", "Kemampuan Numerik 5.1", 
            "Sistematika Kerja/ C D R.1", "Orientasi Hasil/ N G.1", "Fleksibilitas/ T V.1", "Motivasi Berprestasi/ A.1", 
            "Kerjasama/ P I.1", "Keterampilan Interpersonal/ B S.1", "Stabilitas Emosi/ E PHQ.1", "Pegembangan Diri/ W.1", 
            "Mengelola Perubahan/ Z K.1"
        ]
        # Daftar kolom yang harus dihapus dari tampilan
        self.columns_to_hide = [
            'Unnamed: 13', 'Unnamed: 14', 'Intelegensi Umum.1', 'Daya Analisa/ AN.1', 
            'Kemampuan Verbal/WA GE.1', 'Kemampuan Numerik/ RA ZR.1', 'Daya Ingat/ME.1', 'Fleksibilitas',
            'Sistematika Kerja/ cd.1', 'Inisiatif/W.1', 'Stabilitas Emosi / E.1', 'Komunikasi / B O.1', 
            'Keterampilan Sosial / X S', 'Kerjasama'
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
        
        personal_fields = ["No", "No Tes", "Tgl Test", "TGL Lahir", "Nama PT", "JK", "SDR/SDRI", "Nama Peserta", "PHQ", "Keterangan PHQ"]
        self.personal_inputs = []
        for i, placeholder in enumerate(personal_fields):
            label = QLabel(placeholder + ":")
            label.setFont(QFont("Arial", 10))
            
            # Special handling for JK, TGL Lahir, dan Tgl Test
            if placeholder == "TGL Lahir" or placeholder == "Tgl Test":
                field = QPushButton("Pilih Tanggal")
                field.clicked.connect(lambda checked=False, placeholder=placeholder: self.show_calendar(placeholder))
            elif placeholder == "JK":
                field = QPushButton("Pilih Jenis Kelamin")
                field.clicked.connect(self.show_gender_dialog)
            elif placeholder == "SDR/SDRI" or placeholder == "Keterangan PHQ":
                # Read-only fields that will be auto-populated
                field = QLineEdit()
                field.setReadOnly(True)
                field.setStyleSheet("background-color: #f0f0f0;")  # Light gray background to indicate read-only
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

        # Connect event handlers for auto-populating fields
        # PHQ field index is 8, Keterangan PHQ is 9
        self.personal_inputs[8].textChanged.connect(self.update_keterangan_phq)
        # JK field index is 5, SDR/SDRI is 6
        self.personal_inputs[5].clicked.connect(self.update_sdr_sdri)  # Will update after gender dialog closes

        # Group for IST
        ist_group = QGroupBox("IST")
        ist_group.setFont(QFont("Arial", 11, QFont.Bold))
        ist_layout = QGridLayout()  # Changed to grid layout
        
        # Sesuaikan field IST berdasarkan log
        ist_fields = ["IQ ", "SE / Konkrit Praktis", "WA/ Verbal", " AN / Flexibilitas Pikir", 
                      "GE / Daya Abstraksi Verbal", "RA / Berpikir Praktis", "Unnamed: 16"]
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
        
        # Sesuaikan field PAPIKOSTICK berdasarkan log
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
                value = item.text() if item else ""
                # Ganti 'nan' dengan string kosong
                if value.lower() == 'nan':
                    value = ""
                row_data[column_name] = value

            # Debug prints
            print(f"Selected Row: {selected_row}")
            
            # Hanya tampilkan kolom yang digunakan (hapus log untuk kolom yang tidak dipakai)
            used_columns = [col for col in self.columns if col not in self.columns_to_hide]
            
            for key, value in row_data.items():
                if key in used_columns:
                    print(f"{key}: {value}")

            # Populate personal inputs
            for i, field in enumerate(self.personal_inputs):
                col_name = self.columns[i] if i < len(self.columns) else ""
                if col_name:
                    value = row_data.get(col_name, "")
                    # Skip kolom 6 (SDR/SDRI) dan 9 (Keterangan PHQ) karena akan diupdate otomatis
                    if i != 6 and i != 9:  
                        field.setText(value)
            
            # Manually update SDR/SDRI based on JK
            jk_value = self.personal_inputs[5].text()
            if jk_value == "P":
                self.personal_inputs[6].setText("Sdri.")
            elif jk_value == "L":
                self.personal_inputs[6].setText("Sdr.")
            else:
                self.personal_inputs[6].setText("")
                
            # Manually update Keterangan PHQ based on PHQ
            phq_text = self.personal_inputs[8].text().strip()
            if phq_text:
                try:
                    phq_value = float(phq_text)
                    keterangan_field = self.personal_inputs[9]
                    
                    # Apply formula: =@IFS(I4<5,"Tidak ada",I4<10,"Ringan",I4<15,"Sedang",I4<20,"Cukup Berat",I4<28,"Parah")
                    if phq_value < 5:
                        keterangan_field.setText("Tidak ada")
                    elif phq_value < 10:
                        keterangan_field.setText("Ringan")
                    elif phq_value < 15:
                        keterangan_field.setText("Sedang")
                    elif phq_value < 20:
                        keterangan_field.setText("Cukup Berat")
                    else:
                        keterangan_field.setText("Parah")
                except ValueError:
                    self.personal_inputs[9].setText("")
            else:
                self.personal_inputs[9].setText("")

            # Populate IST inputs
            ist_start_idx = len(self.personal_inputs)
            for i, field in enumerate(self.ist_inputs):
                col_idx = ist_start_idx + i
                if col_idx < len(self.columns):
                    value = row_data.get(self.columns[col_idx], "")
                    # Ganti 'nan' dengan string kosong
                    if value.lower() == 'nan':
                        value = ""
                    field.setText(value)

            # Populate PAPIKOSTICK inputs menggunakan mapping khusus
            papiko_columns = ["N", "G", "A", "L", "P", "I", "T", "V", "S", "B", "O", "X", "C", "D", "R", "Z", "E", "K", "F", "W"]
            
            # Debug untuk mapping PAPIKOSTICK
            print("DEBUG - Populate PAPIKOSTICK fields from selection:")
            
            # Telusuri setiap kolom PAPIKOSTICK untuk menemukan indeks yang tepat
            papiko_indices = {}
            for col_name in papiko_columns:
                col_idx = self.get_column_index(col_name)
                if col_idx >= 0:
                    papiko_indices[col_name] = col_idx
                    
            # Telusuri setiap input field PAPIKOSTICK
            for field_idx, field_label in enumerate(papiko_columns):
                if field_label in papiko_indices:
                    col_idx = papiko_indices[field_label]
                    value = self.get_cell_text(selected_row, col_idx)
                    
                    # Ganti 'nan' dengan string kosong
                    if value.lower() == 'nan':
                        value = ""
                    
                    # Log debugging
                    print(f"DEBUG - Populate field PAPIKOSTICK {field_idx} ({field_label}) dengan nilai: '{value}'")
                    
                    # Skip kolom C (Coding) karena akan dihitung otomatis
                    if field_label == "C (Coding)":
                        print(f"DEBUG - Melewati field C (Coding) karena dihitung otomatis")
                        continue
                    
                    # Set nilai ke field input yang sesuai
                    if field_idx < len(self.papikostick_inputs):
                        self.papikostick_inputs[field_idx].setText(value)
                else:
                    print(f"DEBUG - Kolom {field_label} tidak ditemukan di tabel")
                    # Kosongkan kolom jika tidak ditemukan
                    if field_idx < len(self.papikostick_inputs):
                        self.papikostick_inputs[field_idx].setText("")

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
            
            # Membaca seluruh sheet dalam file Excel dengan opsi keep_default_na=False
            # untuk mencegah nilai kosong menjadi NaN
            self.excel_data = pd.read_excel(file_path, sheet_name=None, keep_default_na=False)  # Baca semua sheet
            
            # Pastikan Sheet1 dan Sheet2 terbaca
            if "Sheet1" in self.excel_data:
                self.sheet1_data = self.excel_data["Sheet1"]
                # Ganti nilai NaN dengan string kosong
                self.sheet1_data = self.sheet1_data.fillna("").astype(str).replace("nan", "")
            else:
                print("Sheet1 tidak ditemukan!")
                self.btn_select.setEnabled(True)
                return

            if "Sheet2" in self.excel_data:
                self.sheet2_data = self.excel_data["Sheet2"]
                # Ganti nilai NaN dengan string kosong
                self.sheet2_data = self.sheet2_data.fillna("").astype(str).replace("nan", "")
            else:
                print("Sheet2 tidak ditemukan!")
                self.btn_select.setEnabled(True)
                return

            # Proses data setelah membaca
            self.process_excel(file_path)
            
        except Exception as e:
            print(f"Error loading file: {e}")
            self.btn_select.setEnabled(True)
            QMessageBox.critical(self, "Error", "Failed to load Excel file. Please try again.")
    def process_excel(self, file_path):
        try:
            # Baca semua sheet dalam file Excel
            sheets = pd.read_excel(file_path, sheet_name=None, engine='openpyxl', keep_default_na=False)

            # Pastikan Sheet1 dan Sheet2 ada
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

            # Proses Sheet1 (jika ada)
            if df_sheet1 is not None:
                # Tampilkan kolom asli untuk debugging
                original_cols = df_sheet1.columns.tolist()
                print("Original columns (Sheet1):", original_cols)

                # Mencari baris awal data berdasarkan keyword "No"
                start_row = None
                for idx, row in df_sheet1.iterrows():
                    if any(str(cell).strip().lower() == 'no' for cell in row):
                        start_row = idx
                        break

                if start_row is not None:
                    # Tambahkan parameter keep_default_na=False untuk mencegah nilai kosong menjadi NaN
                    df_sheet1 = pd.read_excel(file_path, sheet_name="Sheet1", engine='openpyxl', 
                                             skiprows=start_row+1, keep_default_na=False)
                    new_df = df_sheet1.copy()

                    # Ganti nilai 'nan' dengan string kosong untuk semua kolom
                    for col in new_df.columns:
                        new_df[col] = new_df[col].astype(str).replace('nan', '')
                        # Ganti nilai 'NaN' dengan string kosong
                        new_df[col] = new_df[col].replace('NaN', '')
                    
                    # Pastikan semua kolom PAPIKOSTICK ada dalam dataframe
                    papiko_columns = ["N", "G", "A", "L", "P", "I", "T", "V", "S", "B", "O", "X", "C", "D", "R", "Z", "E", "K", "F", "W"]
                    for col in papiko_columns:
                        if col not in new_df.columns:
                            new_df[col] = ""  # Tambahkan kolom jika tidak ada
                            print(f"Menambahkan kolom PAPIKOSTICK yang tidak ada: {col}")
                        elif new_df[col].isna().any() or (new_df[col] == "").any() or (new_df[col].str.lower() == "nan").any():
                            print(f"Mengisi nilai kosong/NaN pada kolom: {col}")
                            new_df[col] = new_df[col].replace(["", "nan", "NaN"], "0")

                    # Konversi kolom tertentu menjadi string
                    str_columns = ['No', 'No Tes', 'Tgl Test', 'TGL Lahir', 'Nama PT', 'JK', 'SDR/SDRI', 'Nama Peserta', 'PHQ', 'Keterangan PHQ']
                    for col in str_columns:
                        if col in new_df.columns:
                            new_df[col] = new_df[col].astype(str)
                            # Pastikan nilai kolom tidak 'nan'
                            new_df[col] = new_df[col].replace('nan', '')

                    # Konversi kolom angka ke numeric - sesuaikan dengan kolom yang dipakai
                    numeric_columns = ['IQ ', 'SE / Konkrit Praktis', 'WA/ Verbal', ' AN / Flexibilitas Pikir', 
                                       'GE / Daya Abstraksi Verbal', 'RA / Berpikir Praktis', 'Unnamed: 16']
                    for col in numeric_columns:
                        if col in new_df.columns:
                            # Simpan nilai asli tanpa konversi untuk kolom 'IQ ' dan 'Unnamed: 16'
                            if col == 'IQ ' or col == 'Unnamed: 16':
                                # Hanya ganti string kosong dan 'nan' dengan "" tetapi biarkan nilai lainnya
                                new_df[col] = new_df[col].replace(['nan', 'NaN'], '')
                                # Pastikan angka 0 string tidak diganti
                                new_df[col] = new_df[col].apply(lambda x: '' if x.strip() == '0' and col in ['IQ ', 'Unnamed: 16'] else x)
                            else:
                                # Untuk kolom lain, lakukan seperti biasa
                                new_df[col] = new_df[col].replace('', '0')
                                new_df[col] = pd.to_numeric(new_df[col], errors='coerce').fillna(0)
                        else:
                            # Tambahkan kolom yang tidak ada, dengan nilai kosong untuk IQ dan Unnamed:16
                            if col == 'IQ ' or col == 'Unnamed: 16':
                                new_df[col] = ''
                            else:
                                new_df[col] = 0

                    # Update SDR/SDRI berdasarkan JK
                    if 'JK' in new_df.columns and 'SDR/SDRI' in new_df.columns:
                        # Formula: =IF(F4="P","Sdri.","Sdr.")
                        def get_sdr_sdri(jk):
                            if jk == 'P':
                                return "Sdri."
                            elif jk == 'L':
                                return "Sdr."
                            else:
                                return ""
                        
                        new_df['SDR/SDRI'] = new_df['JK'].apply(get_sdr_sdri)
                        
                    # Update Keterangan PHQ berdasarkan PHQ
                    if 'PHQ' in new_df.columns and 'Keterangan PHQ' in new_df.columns:
                        def get_keterangan_phq(phq_str):
                            try:
                                if phq_str.strip() == '':
                                    return ''
                                phq = float(phq_str)
                                # Formula: =@IFS(I4<5,"Tidak ada",I4<10,"Ringan",I4<15,"Sedang",I4<20,"Cukup Berat",I4<28,"Parah")
                                if phq < 5:
                                    return "Tidak ada"
                                elif phq < 10:
                                    return "Ringan"
                                elif phq < 15:
                                    return "Sedang"
                                elif phq < 20:
                                    return "Cukup Berat"
                                else:
                                    return "Parah"
                            except (ValueError, AttributeError):
                                return ''
                                
                        new_df['Keterangan PHQ'] = new_df['PHQ'].apply(get_keterangan_phq)
                        
                    # Hitung IQ berdasarkan nilai-nilai IST
                    if all(col in new_df.columns for col in ["SE / Konkrit Praktis", "WA/ Verbal", " AN / Flexibilitas Pikir", "GE / Daya Abstraksi Verbal", "RA / Berpikir Praktis"]):
                        # Convert columns to numeric for calculation
                        numeric_cols = ["SE / Konkrit Praktis", "WA/ Verbal", " AN / Flexibilitas Pikir", "GE / Daya Abstraksi Verbal", "RA / Berpikir Praktis"]
                        for col in numeric_cols:
                            new_df[col] = pd.to_numeric(new_df[col], errors='coerce').fillna(0)
                        
                        # Formula: =SUM(L4:P4)/5
                        new_df['IQ '] = (new_df["SE / Konkrit Praktis"] + new_df["WA/ Verbal"] + new_df[" AN / Flexibilitas Pikir"] + 
                                         new_df["GE / Daya Abstraksi Verbal"] + new_df["RA / Berpikir Praktis"]) / 5
                    
                    # Hitung Unnamed: 16 = (AN + GE)/2
                    if all(col in new_df.columns for col in [" AN / Flexibilitas Pikir", "GE / Daya Abstraksi Verbal"]):
                        # Formula: =(M4+O4)/2
                        new_df['Unnamed: 16'] = (new_df[" AN / Flexibilitas Pikir"] + new_df["GE / Daya Abstraksi Verbal"]) / 2

                    # Konversi angka ke string untuk tampilan
                    for col in numeric_columns:
                        if col in new_df.columns and col not in ['IQ ', 'Unnamed: 16']:
                            new_df[col] = new_df[col].astype(str)

                    # Pastikan kolom W memiliki nilai
                    if "W" in new_df.columns:
                        # Jika ada nilai W di file Excel, gunakan nilai tersebut
                        # Pastikan nilai W tidak 'nan'
                        new_df["W"] = new_df["W"].replace('nan', '')
                    else:
                        # Jika tidak ada, buat kolom dengan nilai kosong
                        new_df["W"] = ""
                        print("Menambahkan kolom W dengan nilai kosong")

                    # Hitung KLASIFIKASI berdasarkan IQ
                    if "IQ " in new_df.columns:
                        # Gunakan konversi yang lebih aman untuk IQ
                        def safe_convert_iq(val):
                            try:
                                if val.strip() == '':
                                    return 0
                                return float(val)
                            except (ValueError, AttributeError):
                                return 0
                                
                        iq_vals = new_df["IQ "].apply(safe_convert_iq)
                        
                        def get_klasifikasi(iq):
                            # Formula: =@IFS(K4<79,"Rendah",K4<90,"Dibawah Rata-Rata",K4<110,"Rata-Rata",K4<120,"Diatas Rata-Rata",K4>119,"Superior")
                            if iq < 79:
                                return "Rendah"
                            elif 79 <= iq < 90:
                                return "Dibawah Rata-Rata"
                            elif 90 <= iq < 110:
                                return "Rata-Rata"
                            elif 110 <= iq < 120:
                                return "Diatas Rata-Rata"
                            else:
                                return "Superior"
                        
                        new_df["KLASIFIKASI"] = iq_vals.apply(get_klasifikasi)
                        print("Menghitung ulang nilai KLASIFIKASI berdasarkan IQ")

                    # Tambahkan kolom penting lainnya jika belum ada
                    important_columns = ["KLASIFIKASI"]
                    for col in important_columns:
                        if col not in new_df.columns:
                            print(f"Menambahkan kolom yang hilang: {col}")
                            new_df[col] = ""

                    # Pastikan tidak ada nilai NaN yang tersisa
                    self.df_sheet1 = new_df.fillna("")
                    self.columns = list(new_df.columns)
                    self.show_table(self.df_sheet1)
                    
                    # Recalculate values for all rows
                    for row in range(self.table.rowCount()):
                        self.recalculate_values(row)
                        
                else:
                    print("Could not find the start of data in Sheet1")
                    self.btn_select.setEnabled(True)
                    return

            # Proses Sheet2 (jika ada)
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
            # Pastikan kolom-kolom penting ada
            important_columns = ["KLASIFIKASI"]
            df_with_columns = df.copy()
            
            # Tambahkan kolom yang hilang
            for col in important_columns:
                if col not in df_with_columns.columns:
                    print(f"Menambahkan kolom yang hilang: {col}")
                    df_with_columns[col] = ""
            
            # Setup table untuk menampilkan data
            self.table.setRowCount(df_with_columns.shape[0])
            self.table.setColumnCount(len(df_with_columns.columns))
            self.table.setHorizontalHeaderLabels(df_with_columns.columns)
            
            # Populate table with data
            for i, row in df_with_columns.iterrows():
                for j, (col_name, val) in enumerate(row.items()):
                    item = QTableWidgetItem(str(val))
                    self.table.setItem(i, j, item)
                    # Debug untuk kolom W
                    if col_name == "W":
                        print(f"DEBUG - Show_table: Mengatur nilai W='{val}' ke tabel di baris {i}, kolom {j}")
            
            # Sembunyikan kolom-kolom yang tidak dipakai
            for col_name in self.columns_to_hide:
                if col_name in df_with_columns.columns:
                    col_idx = df_with_columns.columns.get_loc(col_name)
                    self.table.hideColumn(col_idx)
                    print(f"Menyembunyikan kolom yang tidak dipakai: {col_name} (indeks {col_idx})")
            
            # Resize columns to fit content
            self.table.resizeColumnsToContents()
            
            # Perbarui daftar kolom
            self.columns = list(df_with_columns.columns)
            
            print("Berhasil menampilkan tabel")
        except Exception as e:
            print(f"Error saat menampilkan tabel: {e}")
            import traceback
            traceback.print_exc()

    def show_calendar(self, field_type):
        calendar = QCalendarWidget(self)
        calendar.setWindowFlags(Qt.Popup)
        
        if field_type == "TGL Lahir":
            # For TGL Lahir, use the 3rd input field
            target_field = self.personal_inputs[3]
            calendar.clicked.connect(lambda date: self.set_date(date, "TGL Lahir"))
        elif field_type == "Tgl Test":
            # For Tgl Test, use the 2nd input field
            target_field = self.personal_inputs[2]
            calendar.clicked.connect(lambda date: self.set_date(date, "Tgl Test"))
        
        # Position calendar below the button
        pos = target_field.mapToGlobal(target_field.rect().bottomLeft())
        calendar.move(pos)
        calendar.show()

    def set_date(self, date, field_type):
        # Format date as dd/MM/yyyy
        formatted_date = date.toString("dd/MM/yyyy")
        
        if field_type == "TGL Lahir":
            # Set TGL Lahir field (index 3)
            self.personal_inputs[3].setText(formatted_date)
        elif field_type == "Tgl Test":
            # Set Tgl Test field (index 2)
            self.personal_inputs[2].setText(formatted_date)
            
    def update_sdr_sdri(self):
        # This will be called after gender dialog closes
        jk_value = self.personal_inputs[5].text()
        sdr_sdri_field = self.personal_inputs[6]
        
        # Formula: =IF(F4="P","Sdri.","Sdr.")
        if jk_value == "P":
            sdr_sdri_field.setText("Sdri.")
        elif jk_value == "L":
            sdr_sdri_field.setText("Sdr.")
        else:
            sdr_sdri_field.setText("")
            
    def update_keterangan_phq(self):
        try:
            # Get PHQ value
            phq_text = self.personal_inputs[8].text().strip()
            keterangan_field = self.personal_inputs[9]
            
            if not phq_text:
                keterangan_field.setText("")
                return
                
            phq_value = float(phq_text)
            
            # Apply formula: =@IFS(I4<5,"Tidak ada",I4<10,"Ringan",I4<15,"Sedang",I4<20,"Cukup Berat",I4<28,"Parah")
            if phq_value < 5:
                keterangan_field.setText("Tidak ada")
            elif phq_value < 10:
                keterangan_field.setText("Ringan")
            elif phq_value < 15:
                keterangan_field.setText("Sedang")
            elif phq_value < 20:
                keterangan_field.setText("Cukup Berat")
            else:
                keterangan_field.setText("Parah")
        except ValueError:
            # Handle case where PHQ is not a valid number
            self.personal_inputs[9].setText("")

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
                self.personal_inputs[5].setText("L")
            elif radio_p.isChecked():
                self.personal_inputs[5].setText("P")
            
            # Update SDR/SDRI field based on gender
            self.update_sdr_sdri()

    def add_or_update_row(self, mode="add"):
        # Collect values from all three input groups
        values = []
        
        # For personal fields
        for i, field in enumerate(self.personal_inputs):
            # Ensure SDR/SDRI and Keterangan PHQ are set correctly
            if i == 6:  # SDR/SDRI field
                jk_value = self.personal_inputs[5].text()
                if jk_value == "P":
                    values.append("Sdri.")
                elif jk_value == "L":
                    values.append("Sdr.")
                else:
                    values.append("")
            elif i == 9:  # Keterangan PHQ field
                phq_text = self.personal_inputs[8].text().strip()
                if phq_text:
                    try:
                        phq_value = float(phq_text)
                        if phq_value < 5:
                            values.append("Tidak ada")
                        elif phq_value < 10:
                            values.append("Ringan")
                        elif phq_value < 15:
                            values.append("Sedang")
                        elif phq_value < 20:
                            values.append("Cukup Berat")
                        else:
                            values.append("Parah")
                    except ValueError:
                        values.append("")
                else:
                    values.append("")
            else:
                values.append(field.text())
        
        # Add IST values
        values.extend([field.text() for field in self.ist_inputs])
        
        # Add PAPIKOSTICK values
        papikostick_values = []
        for field in self.papikostick_inputs:
            text = field.text().strip()
            # Jika teks kosong, gunakan "0" untuk X dan PAPIKOSTICK lainnya
            if not text:
                papikostick_values.append("0")
            else:
                papikostick_values.append(text)
        values.extend(papikostick_values)
        
        # Debug information
        print(f"Number of values: {len(values)}")
        print(f"Number of columns: {len(self.columns)}")
        
        # Debug untuk melihat struktur kolom secara detail
        for i, col in enumerate(self.columns):
            print(f"DEBUG - Column {i}: {col}")
            
        # Debug untuk papikostick fields
        print("DEBUG - PAPIKOSTICK fields:")
        for i, field in enumerate(self.papikostick_inputs):
            print(f"DEBUG - Field {i}: {field.text()}")
        
        print(f"Columns: {self.columns}")  # Print column names for debugging
        
        # Check if any field is empty
        if any(not val.strip() for val in values[:5]) or not values[7].strip():  # Check important fields only
            print("Please fill all required fields (No, No Tes, Tgl Test, TGL Lahir, Nama PT, JK, Nama Peserta)")
            QMessageBox.warning(self, "Warning", "Harap isi semua field penting (No, No Tes, Tgl Test, TGL Lahir, Nama PT, JK, Nama Peserta)")
            return
        
        # Ensure the values list has the correct number of elements
        expected_fields = len(self.personal_inputs) + len(self.ist_inputs) + len(self.papikostick_inputs)
        if len(values) < expected_fields:
            print(f"Expected {expected_fields} fields, got {len(values)}")
            return
        
        try:
            # Convert numeric values
            for i in range(10, len(values)):  # Start after personal fields
                if i < len(values) and values[i].strip():  # Check if index is within bounds and value is not empty
                    try:
                        values[i] = float(values[i])
                    except ValueError:
                        print(f"Invalid numeric value: {values[i]} at position {i}")
                        QMessageBox.warning(self, "Warning", f"Nilai tidak valid: {values[i]} pada posisi {i+1}")
                        return
        except ValueError as e:
            print(f"Please enter valid numeric values for numeric fields: {e}")
            QMessageBox.warning(self, "Warning", f"Harap masukkan nilai numerik yang valid: {e}")
            return
        
        # Create a complete row with all columns
        row_data = {}
        
        # Add personal info
        for i, field in enumerate(self.personal_inputs):
            if i < len(self.columns):
                if i == 6:  # SDR/SDRI
                    jk_value = self.personal_inputs[5].text()
                    if jk_value == "P":
                        row_data[self.columns[i]] = "Sdri."
                    elif jk_value == "L":
                        row_data[self.columns[i]] = "Sdr."
                    else:
                        row_data[self.columns[i]] = ""
                elif i == 9:  # Keterangan PHQ
                    phq_text = self.personal_inputs[8].text().strip()
                    if phq_text:
                        try:
                            phq_value = float(phq_text)
                            if phq_value < 5:
                                row_data[self.columns[i]] = "Tidak ada"
                            elif phq_value < 10:
                                row_data[self.columns[i]] = "Ringan"
                            elif phq_value < 15:
                                row_data[self.columns[i]] = "Sedang"
                            elif phq_value < 20:
                                row_data[self.columns[i]] = "Cukup Berat"
                            else:
                                row_data[self.columns[i]] = "Parah"
                        except ValueError:
                            row_data[self.columns[i]] = ""
                    else:
                        row_data[self.columns[i]] = ""
                else:
                    row_data[self.columns[i]] = field.text()
        
        # Add IST values
        ist_start_idx = len(self.personal_inputs)
        ist_values = {}
        for i, field in enumerate(self.ist_inputs):
            col_idx = ist_start_idx + i
            if col_idx < len(self.columns):
                try:
                    input_text = field.text()
                    # Jika input kosong, biarkan kosong (string kosong, bukan 0)
                    if not input_text.strip():
                        row_data[self.columns[col_idx]] = ""
                    else:
                        # Simpan untuk perhitungan IQ
                        ist_values[self.columns[col_idx]] = float(input_text)
                        row_data[self.columns[col_idx]] = float(input_text)
                except ValueError:
                    print(f"Invalid numeric value: {field.text()}")
                    QMessageBox.warning(self, "Warning", f"Nilai tidak valid: {field.text()}")
                    return
        
        # Hitung nilai IQ = SUM(L4:P4)/5
        if all(col in ist_values for col in ["SE / Konkrit Praktis", "WA/ Verbal", " AN / Flexibilitas Pikir", "GE / Daya Abstraksi Verbal", "RA / Berpikir Praktis"]):
            iq_val = (ist_values["SE / Konkrit Praktis"] + ist_values["WA/ Verbal"] + ist_values[" AN / Flexibilitas Pikir"] + 
                     ist_values["GE / Daya Abstraksi Verbal"] + ist_values["RA / Berpikir Praktis"]) / 5
            row_data["IQ "] = iq_val
            print(f"Menghitung IQ dalam add_or_update_row: {iq_val}")
            
            # Hitung Unnamed: 16 = (AN + GE)/2
            unnamed_16_val = (ist_values[" AN / Flexibilitas Pikir"] + ist_values["GE / Daya Abstraksi Verbal"]) / 2
            row_data["Unnamed: 16"] = unnamed_16_val
            print(f"Menghitung Unnamed: 16 dalam add_or_update_row: {unnamed_16_val}")
            
            # Hitung KLASIFIKASI berdasarkan IQ
            # Formula: =@IFS(K4<79,"Rendah",K4<90,"Dibawah Rata-Rata",K4<110,"Rata-Rata",K4<120,"Diatas Rata-Rata",K4>119,"Superior")
            if iq_val < 79:
                row_data["KLASIFIKASI"] = "Rendah"
            elif 79 <= iq_val < 90:
                row_data["KLASIFIKASI"] = "Dibawah Rata-Rata"
            elif 90 <= iq_val < 110:
                row_data["KLASIFIKASI"] = "Rata-Rata"
            elif 110 <= iq_val < 120:
                row_data["KLASIFIKASI"] = "Diatas Rata-Rata"
            else:
                row_data["KLASIFIKASI"] = "Superior"
            print(f"Menghitung KLASIFIKASI dalam add_or_update_row: {row_data['KLASIFIKASI']}")
        
        # Add PAPIKOSTICK values dengan pendekatan yang lebih langsung
        papiko_columns = ["N", "G", "A", "L", "P", "I", "T", "V", "S", "B", "O", "X", "C", "D", "R", "Z", "E", "K", "F", "W"]
        
        # Debug informasi kolom
        print("DEBUG - Mapping PAPIKOSTICK ke kolom di tabel:")
        
        # Telusuri setiap kolom PAPIKOSTICK untuk menemukan indeks yang tepat
        papiko_indices = {}
        for col_name in papiko_columns:
            for col_idx, table_col in enumerate(self.columns):
                if table_col == col_name:
                    papiko_indices[col_name] = col_idx
                    print(f"DEBUG - Kolom {col_name} ada di indeks tabel {col_idx}")
                    break
        
        # Telusuri setiap input field PAPIKOSTICK
        for field_idx, field_label in enumerate(papiko_columns):
            if field_idx < len(self.papikostick_inputs):
                input_text = self.papikostick_inputs[field_idx].text().strip()
                
                # Pastikan kolom ada di tabel
                if field_label in papiko_indices:
                    col_idx = papiko_indices[field_label]
                    
                    # Skip kolom C (Coding) karena akan dihitung otomatis
                    if field_label == "C (Coding)":
                        print(f"DEBUG - Melewati kolom C (Coding)")
                        continue
                    
                    # Pastikan input kosong menjadi "0" untuk X dan papiko lainnya
                    if not input_text:
                        # Untuk X, konversi kosong ke "0"
                        input_text = "0"
                    
                    # Konversi ke tipe data yang sesuai (float/int)
                    try:
                        value = float(input_text)
                        # Jika nilai bulat, simpan sebagai int
                        if value.is_integer():
                            row_data[self.columns[col_idx]] = int(value)
                        else:
                            row_data[self.columns[col_idx]] = value
                    except ValueError:
                        # Jika tidak bisa dikonversi, gunakan string
                        row_data[self.columns[col_idx]] = input_text
                    
                    # Log khusus untuk kolom
                    print(f"DEBUG - Menetapkan nilai kolom {field_label} ke {row_data[self.columns[col_idx]]}")
        
        # Determine action based on mode
        if mode == "add":
            # Add new row to the table
            row = self.table.rowCount()
            self.table.insertRow(row)
            for col, column_name in enumerate(self.columns):
                value = row_data.get(column_name, "")
                
                # Penanganan khusus untuk nilai X
                if column_name == "X" and (value == "" or value == 0):
                    value = 0
                
                # Penanganan khusus untuk kolom W
                if column_name == "W":
                    print(f"DEBUG - Setting W value in table to: '{value}'")
                    # Pastikan nilai W disimpan dengan benar
                    w_item = QTableWidgetItem()
                    if isinstance(value, (int, float)):
                        if isinstance(value, float) and value.is_integer():
                            w_item.setData(Qt.DisplayRole, int(value))
                        else:
                            w_item.setData(Qt.DisplayRole, value)
                    else:
                        w_item.setText(str(value))
                    
                    self.table.setItem(row, col, w_item)
                    continue  # Lanjutkan ke kolom berikutnya

                # Penanganan khusus untuk kolom deskripsi psikogram
                elif col >= 54 and col <= 65 and value:
                    # Gunakan setText langsung untuk memastikan nilai teks dipertahankan
                    psiko_item = QTableWidgetItem()
                    psiko_item.setText(str(value))
                    self.table.setItem(row, col, psiko_item)
                    continue  # Lanjutkan ke kolom berikutnya
                
                # Penanganan untuk kolom-kolom lainnya
                else:
                    # Buat QTableWidgetItem dengan nilai untuk kolom lainnya
                    table_item = QTableWidgetItem()
                    
                    # Jika nilai numerik (integer atau float), simpan sebagai numerik
                    if isinstance(value, (int, float)):
                        if isinstance(value, float) and value.is_integer():
                            # Jika float tanpa desimal, konversi ke integer
                            table_item.setData(Qt.DisplayRole, int(value))
                        else:
                            table_item.setData(Qt.DisplayRole, value)
                    else:
                        # Jika string atau lainnya, simpan sebagai string
                        table_item.setText(str(value))
                    
                    self.table.setItem(row, col, table_item)
            
            # Recalculate values for the new row
            self.recalculate_values(row)
            
        elif mode == "edit":
            # Update existing row
            selected_row = self.table.currentRow()
            if selected_row >= 0:
                for col, column_name in enumerate(self.columns):
                    value = row_data.get(column_name, "")
                    
                    # Penanganan khusus untuk nilai X
                    if column_name == "X" and (value == "" or value == 0):
                        value = 0
                    
                    # Penanganan khusus untuk kolom W
                    if column_name == "W":
                        print(f"DEBUG - Updating W value in table to: '{value}'")
                        # Pastikan nilai W disimpan dengan benar
                        w_item = QTableWidgetItem()
                        if isinstance(value, (int, float)):
                            if isinstance(value, float) and value.is_integer():
                                w_item.setData(Qt.DisplayRole, int(value))
                            else:
                                w_item.setData(Qt.DisplayRole, value)
                        else:
                            w_item.setText(str(value))
                        
                        self.table.setItem(selected_row, col, w_item)
                        continue  # Lanjutkan ke kolom berikutnya
    
                    # Penanganan khusus untuk kolom deskripsi psikogram
                    elif col >= 54 and col <= 65 and value:
                        # Gunakan setText langsung untuk memastikan nilai teks dipertahankan
                        psiko_item = QTableWidgetItem()
                        psiko_item.setText(str(value))
                        self.table.setItem(selected_row, col, psiko_item)
                        continue  # Lanjutkan ke kolom berikutnya
                    
                    # Penanganan untuk kolom-kolom lainnya
                    else:
                        # Buat QTableWidgetItem dengan nilai untuk kolom lainnya
                        table_item = QTableWidgetItem()
                        
                        # Jika nilai numerik (integer atau float), simpan sebagai numerik
                        if isinstance(value, (int, float)):
                            if isinstance(value, float) and value.is_integer():
                                # Jika float tanpa desimal, konversi ke integer
                                table_item.setData(Qt.DisplayRole, int(value))
                            else:
                                table_item.setData(Qt.DisplayRole, value)
                        else:
                            # Jika string atau lainnya, simpan sebagai string
                            table_item.setText(str(value))
                        
                        self.table.setItem(selected_row, col, table_item)
                
                # Recalculate values for the edited row
                self.recalculate_values(selected_row)
            else:
                print("No row selected for editing")
                QMessageBox.warning(self, "Warning", "Tidak ada baris yang dipilih untuk diedit")

        # Clear input fields after adding/updating
        for field in self.personal_inputs + self.ist_inputs + self.papikostick_inputs:
            if isinstance(field, QLineEdit):
                field.clear()
            elif isinstance(field, QPushButton):
                if field.text() not in ["Pilih Jenis Kelamin", "Pilih Tanggal"]:
                    field.setText("Pilih Tanggal" if "Tanggal" in field.text() else field.text())

    def recalculate_values(self, row):
        try:
            # Simpan nilai W sebelum perhitungan untuk memastikan tidak berubah
            original_w_value = self.get_cell_text(row, 36)
            print(f"DEBUG - Nilai W SEBELUM perhitungan: '{original_w_value}'")
            
            # Hitung nilai IQ = SUM(L4:P4)/5
            l_idx = self.get_column_index("SE / Konkrit Praktis")
            wa_idx = self.get_column_index("WA/ Verbal")
            an_idx = self.get_column_index(" AN / Flexibilitas Pikir")
            ge_idx = self.get_column_index("GE / Daya Abstraksi Verbal")
            ra_idx = self.get_column_index("RA / Berpikir Praktis")
            iq_idx = self.get_column_index("IQ ")
            
            if l_idx >= 0 and wa_idx >= 0 and an_idx >= 0 and ge_idx >= 0 and ra_idx >= 0 and iq_idx >= 0:
                l_val = self.convert_to_float(self.get_cell_text(row, l_idx))
                wa_val = self.convert_to_float(self.get_cell_text(row, wa_idx))
                an_val = self.convert_to_float(self.get_cell_text(row, an_idx))
                ge_val = self.convert_to_float(self.get_cell_text(row, ge_idx))
                ra_val = self.convert_to_float(self.get_cell_text(row, ra_idx))
                
                if l_val is not None and wa_val is not None and an_val is not None and ge_val is not None and ra_val is not None:
                    # Formula: =SUM(L4:P4)/5
                    iq_val = (l_val + wa_val + an_val + ge_val + ra_val) / 5
                    iq_item = QTableWidgetItem()
                    if iq_val.is_integer():
                        iq_item.setData(Qt.DisplayRole, int(iq_val))
                    else:
                        iq_item.setData(Qt.DisplayRole, iq_val)
                    self.table.setItem(row, iq_idx, iq_item)
                    print(f"Menghitung IQ: {iq_val}")
                    
                    # Hitung Unnamed: 16 = (M4+O4)/2
                    unnamed_16_idx = self.get_column_index("Unnamed: 16")
                    if unnamed_16_idx >= 0:
                        # Formula: =(M4+O4)/2 --> (AN + GE)/2
                        unnamed_16_val = (an_val + ge_val) / 2
                        unnamed_16_item = QTableWidgetItem()
                        if unnamed_16_val.is_integer():
                            unnamed_16_item.setData(Qt.DisplayRole, int(unnamed_16_val))
                        else:
                            unnamed_16_item.setData(Qt.DisplayRole, unnamed_16_val)
                        self.table.setItem(row, unnamed_16_idx, unnamed_16_item)
                        print(f"Menghitung Unnamed: 16 = (AN + GE)/2: {unnamed_16_val}")
                    
                    # Hitung KLASIFIKASI berdasarkan IQ
                    # Formula: =@IFS(K4<79,"Rendah",K4<90,"Dibawah Rata-Rata",K4<110,"Rata-Rata",K4<120,"Diatas Rata-Rata",K4>119,"Superior")
                    if iq_val < 79:
                        iq_klasifikasi = "Rendah"
                    elif 79 <= iq_val < 90:
                        iq_klasifikasi = "Dibawah Rata-Rata"
                    elif 90 <= iq_val < 110:
                        iq_klasifikasi = "Rata-Rata"
                    elif 110 <= iq_val < 120:
                        iq_klasifikasi = "Diatas Rata-Rata"
                    else:
                        iq_klasifikasi = "Superior"
                        
                    # Cari indeks kolom KLASIFIKASI
                    klasifikasi_idx = self.get_column_index("KLASIFIKASI")
                    if klasifikasi_idx >= 0:
                        self.table.setItem(row, klasifikasi_idx, QTableWidgetItem(iq_klasifikasi))
                    print(f"Menghitung KLASIFIKASI: {iq_klasifikasi}")
            
            # Dapatkan nilai C untuk perhitungan C (Coding)
            c_idx = self.get_column_index("C")
            coding_idx = self.get_column_index("C (Coding)")
            
            if c_idx >= 0 and coding_idx >= 0:
                c_val = self.convert_to_float(self.get_cell_text(row, c_idx))
                
                # Formula: =@IFS(AE4=1,9,AE4=2,8,AE4=3,7,AE4=4,6,AE4=5,5,AE4=6,4,AE4=7,3,AE4=8,2,AE4=9,1)
                if c_val is not None and 1 <= c_val <= 9:
                    c_coding = 10 - c_val
                    # Jika c_coding adalah integer, gunakan QTableWidgetItem dengan setData
                    c_coding_item = QTableWidgetItem()
                    if isinstance(c_coding, int) or (isinstance(c_coding, float) and c_coding.is_integer()):
                        c_coding_item.setData(Qt.DisplayRole, int(c_coding))
                    else:
                        c_coding_item.setData(Qt.DisplayRole, c_coding)
                    self.table.setItem(row, coding_idx, c_coding_item)
                    print(f"Formula IFS untuk kolom {coding_idx}: C (Coding) = 10 - C")
                    
            # Dapatkan nilai N dan G untuk perhitungan NG
            n_idx = self.get_column_index("N")
            g_idx = self.get_column_index("G")
            ng_idx = self.get_column_index("NG")
            
            if n_idx >= 0 and g_idx >= 0 and ng_idx >= 0:
                n_val = self.convert_to_float(self.get_cell_text(row, n_idx))
                g_val = self.convert_to_float(self.get_cell_text(row, g_idx))
                
                # Formula: =(S4+T4)/2
                if n_val is not None and g_val is not None:
                    ng_val = (n_val + g_val) / 2
                    ng_item = QTableWidgetItem()
                    if ng_val.is_integer():
                        ng_item.setData(Qt.DisplayRole, int(ng_val))
                    else:
                        ng_item.setData(Qt.DisplayRole, ng_val)
                    self.table.setItem(row, ng_idx, ng_item)
                    print(f"Formula normal untuk kolom {ng_idx}: NG = (N + G) / 2")
                    
            # Dapatkan nilai C, D, R untuk perhitungan CDR
            c_idx = self.get_column_index("C")
            d_idx = self.get_column_index("D")
            r_idx = self.get_column_index("R")
            cdr_idx = self.get_column_index("CDR")
            
            if c_idx >= 0 and d_idx >= 0 and r_idx >= 0 and cdr_idx >= 0:
                c_val = self.convert_to_float(self.get_cell_text(row, c_idx))
                d_val = self.convert_to_float(self.get_cell_text(row, d_idx))
                r_val = self.convert_to_float(self.get_cell_text(row, r_idx))
                
                # Formula: =(AE4+AG4+AH4)/3
                if c_val is not None and d_val is not None and r_val is not None:
                    cdr_val = (c_val + d_val + r_val) / 3
                    cdr_item = QTableWidgetItem()
                    if cdr_val.is_integer():
                        cdr_item.setData(Qt.DisplayRole, int(cdr_val))
                    else:
                        cdr_item.setData(Qt.DisplayRole, cdr_val)
                    self.table.setItem(row, cdr_idx, cdr_item)
                    print(f"Formula normal untuk kolom {cdr_idx}: CDR = (C + D + R) / 3")
                    
            # Dapatkan nilai T dan V untuk perhitungan TV
            t_idx = self.get_column_index("T")
            v_idx = self.get_column_index("V")
            tv_idx = self.get_column_index("TV")
            
            if t_idx >= 0 and v_idx >= 0 and tv_idx >= 0:
                t_val = self.convert_to_float(self.get_cell_text(row, t_idx))
                v_val = self.convert_to_float(self.get_cell_text(row, v_idx))
                
                # Formula: =(Y4+Z4)/2
                if t_val is not None and v_val is not None:
                    tv_val = (t_val + v_val) / 2
                    tv_item = QTableWidgetItem()
                    if tv_val.is_integer():
                        tv_item.setData(Qt.DisplayRole, int(tv_val))
                    else:
                        tv_item.setData(Qt.DisplayRole, tv_val)
                    self.table.setItem(row, tv_idx, tv_item)
                    print(f"Formula normal untuk kolom {tv_idx}: TV = (T + V) / 2")
                    
            # Dapatkan nilai P dan I untuk perhitungan PI
            p_idx = self.get_column_index("P")
            i_idx = self.get_column_index("I")
            pi_idx = self.get_column_index("PI")
            
            if p_idx >= 0 and i_idx >= 0 and pi_idx >= 0:
                p_val = self.convert_to_float(self.get_cell_text(row, p_idx))
                i_val = self.convert_to_float(self.get_cell_text(row, i_idx))
                
                # Formula: =(W4+X4)/2
                if p_val is not None and i_val is not None:
                    pi_val = (p_val + i_val) / 2
                    pi_item = QTableWidgetItem()
                    if pi_val.is_integer():
                        pi_item.setData(Qt.DisplayRole, int(pi_val))
                    else:
                        pi_item.setData(Qt.DisplayRole, pi_val)
                    self.table.setItem(row, pi_idx, pi_item)
                    print(f"Formula normal untuk kolom {pi_idx}: PI = (P + I) / 2")
                    
            # Dapatkan nilai B dan S untuk perhitungan BS
            b_idx = self.get_column_index("B")
            s_idx = self.get_column_index("S")
            bs_idx = self.get_column_index("BS")
            
            if b_idx >= 0 and s_idx >= 0 and bs_idx >= 0:
                b_val = self.convert_to_float(self.get_cell_text(row, b_idx))
                s_val = self.convert_to_float(self.get_cell_text(row, s_idx))
                
                # Formula: =(AA4+AB4)/2
                if b_val is not None and s_val is not None:
                    bs_val = (b_val + s_val) / 2
                    bs_item = QTableWidgetItem()
                    if bs_val.is_integer():
                        bs_item.setData(Qt.DisplayRole, int(bs_val))
                    else:
                        bs_item.setData(Qt.DisplayRole, bs_val)
                    self.table.setItem(row, bs_idx, bs_item)
                    print(f"Formula normal untuk kolom {bs_idx}: BS = (B + S) / 2")
                    
            # Dapatkan nilai Z dan K untuk perhitungan ZK
            z_idx = self.get_column_index("Z")
            k_idx = self.get_column_index("K")
            zk_idx = self.get_column_index("ZK")
            
            if z_idx >= 0 and k_idx >= 0 and zk_idx >= 0:
                z_val = self.convert_to_float(self.get_cell_text(row, z_idx))
                k_val = self.convert_to_float(self.get_cell_text(row, k_idx))
                
                # Formula: =(AI4+AK4)/2
                if z_val is not None and k_val is not None:
                    zk_val = (z_val + k_val) / 2
                    zk_item = QTableWidgetItem()
                    if zk_val.is_integer():
                        zk_item.setData(Qt.DisplayRole, int(zk_val))
                    else:
                        zk_item.setData(Qt.DisplayRole, zk_val)
                    self.table.setItem(row, zk_idx, zk_item)
                    print(f"Formula normal untuk kolom {zk_idx}: ZK = (Z + K) / 2")
            
            # Update nilai W (posisi terakhir)
            w_idx = self.get_column_index("W")
            if w_idx >= 0:
                original_w_value_adjusted = original_w_value.strip() if original_w_value else "0"
                if not original_w_value_adjusted:
                    original_w_value_adjusted = "0"
                    
                self.table.setItem(row, w_idx, QTableWidgetItem(original_w_value_adjusted))
                print(f"DEBUG - Nilai W SETELAH perhitungan: '{original_w_value_adjusted}'")
                
        except Exception as e:
            print(f"Error dalam recalculate_values: {e}")
            import traceback
            traceback.print_exc()

    def get_cell_value(self, row, col):
        item = self.table.item(row, col)
        if item and item.text().strip():
            text = item.text()
            # Jika nilai adalah 'nan', kembalikan None
            if text.lower() == 'nan':
                return None
            try:
                return float(text)
            except ValueError:
                return None
        return None
        
    def get_cell_text(self, row, col):
        item = self.table.item(row, col)
        if item:
            # Jika nilai adalah 'nan', kembalikan string kosong
            if item.text().lower() == 'nan':
                return ""
            return item.text()
        return ""
        
    def convert_to_float(self, text):
        if text and text.strip():
            try:
                value = float(text)
                # Jika nilai bulat, konversi ke integer
                if value.is_integer():
                    return int(value)
                return value
            except ValueError:
                return None
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
                    43: "=IF(L{row}<80,\"R\",IF(L{row}<100,\"K\",IF(L{row}<120,\"C\",IF(L{row}<140,\"B\",\"T\"))))",  # Logika Berpikir 1
                    44: "=IF(M{row}<80,\"R\",IF(M{row}<100,\"K\",IF(M{row}<120,\"C\",IF(M{row}<140,\"B\",\"T\"))))",  # Daya Analisa 3
                    45: "=IF(O{row}<80,\"R\",IF(O{row}<100,\"K\",IF(O{row}<120,\"C\",IF(O{row}<140,\"B\",\"T\"))))",  # Kemampuan Verbal 2 dam 4
                    46: "=IF(N{row}<80,\"R\",IF(N{row}<100,\"K\",IF(N{row}<120,\"C\",IF(N{row}<140,\"B\",\"T\"))))",  # Kemampuan Numerik 5
                    47: "=IF(M{row}<80,\"R\",IF(M{row}<100,\"K\",IF(M{row}<120,\"C\",IF(M{row}<140,\"B\",\"T\"))))",  # Inisiatif
                    48: "=@IFS(AO{row}<2,\"R\",AO{row}<4,\"K\",AO{row}<6,\"C\",AO{row}<9,\"B\",AO{row}=9,\"T\")",  # Sistematika Kerja/ C D R
                    49: "=@IFS(AN{row}<2,\"R\",AN{row}<4,\"K\",AN{row}<6,\"C\",AN{row}<9,\"B\",AN{row}=9,\"T\")",  # Orientasi Hasil/ N G
                    50: "=@IFS(AP{row}<2,\"R\",AP{row}<4,\"K\",AP{row}<6,\"C\",AP{row}<9,\"B\",AP{row}=9,\"T\")",  # Fleksibilitas/ T V
                    51: "=@IFS(U{row}<2,\"R\",U{row}<4,\"K\",U{row}<6,\"C\",U{row}<9,\"B\",U{row}=9,\"T\")",  # Motivasi Berprestasi/ A
                    52: "=@IFS(AQ{row}<2,\"R\",AQ{row}<4,\"K\",AQ{row}<6,\"C\",AQ{row}<9,\"B\",AQ{row}=9,\"T\")",  # Kerjasama/ P I
                    53: "=@IFS(AR{row}<2,\"R\",AR{row}<4,\"K\",AR{row}<6,\"C\",AR{row}<9,\"B\",AR{row}=9,\"T\")",  # Keterampilan Interpersonal/ B S
                    54: "=@IFS(AJ{row}<2,\"R\",AJ{row}<4,\"K\",AJ{row}<6,\"C\",AJ{row}<9,\"B\",AJ{row}=9,\"T\")",  # Stabilitas Emosi/ E PHQ
                    55: "=@IFS(AM{row}<2,\"R\",AM{row}<4,\"K\",AM{row}<6,\"C\",AM{row}<9,\"B\",AM{row}=9,\"T\")",  # Pegembangan Diri/ W
                    56: "=@IFS(AS{row}<2,\"R\",AS{row}<4,\"K\",AS{row}<6,\"C\",AS{row}<9,\"B\",AS{row}=9,\"T\")",  # Mengelola Perubahan/ Z K
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
                        for col_letter in ['F', 'I', 'N', 'O', 'M', 'AK', 'AL', 'AF', 'AH', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB']:
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
            
            # Cari kolom IQ dan Flexibilitas Pikir tanpa spasi di seluruh kolom
            for i, col_name in enumerate(self.columns):
                # Cari kolom duplikat IQ tanpa spasi
                if col_name.strip() == "IQ" and col_name != "IQ " and "IQ " in self.columns:
                    duplicate_columns.append(i)
                    print(f"Kolom duplikat IQ ditemukan di indeks {i}: '{col_name}'")
                
                # Cari kolom duplikat Flexibilitas Pikir tanpa spasi
                if col_name.strip() == "Flexibilitas Pikir" and col_name != "Flexibilitas Pikir " and "Flexibilitas Pikir " in self.columns:
                    duplicate_columns.append(i)
                    print(f"Kolom duplikat Flexibilitas Pikir ditemukan di indeks {i}: '{col_name}'")
            
            # Jika menggunakan versi tanpa spasi (versi lama), ubah informasi di special_columns
            if "IQ " in self.columns and iq_column is None:
                iq_idx = self.columns.index("IQ ")
                iq_column = iq_idx + 1  # +1 karena kolom Excel dimulai dari 1
                print(f"Menggunakan kolom IQ dengan spasi di indeks {iq_idx} (Excel: {iq_column})")
            
            if "Flexibilitas Pikir " in self.columns and flex_column is None:
                flex_idx = self.columns.index("Flexibilitas Pikir ")
                flex_column = flex_idx + 1  # +1 karena kolom Excel dimulai dari 1
                print(f"Menggunakan kolom Flexibilitas Pikir dengan spasi di indeks {flex_idx} (Excel: {flex_column})")
            
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
                
                # Memastikan kolom Unnamed: 13, Unnamed: 14, dan KLASIFIKASI memiliki nilai yang benar
                # WA GE (Unnamed: 13)
                try:
                    if "Verbal" in row_data and "Daya Abstraksi Verbal" in row_data:
                        verbal_val = float(row_data.get("Verbal", 0))
                        dav_val = float(row_data.get("Daya Abstraksi Verbal", 0))
                        wa_ge = (verbal_val + dav_val) / 2
                        row_data["Unnamed: 13"] = str(wa_ge)
                        print(f"Menghitung WA GE: {wa_ge}")
                except (ValueError, TypeError) as e:
                    print(f"Error menghitung WA GE: {e}")

                # RA ZR (Unnamed: 14)
                try:
                    if "Berpikir Praktis" in row_data and "Berpikir Teoritis" in row_data:
                        bp_val = float(row_data.get("Berpikir Praktis", 0))
                        bt_val = float(row_data.get("Berpikir Teoritis", 0))
                        ra_zr = (bp_val + bt_val) / 2
                        row_data["Unnamed: 14"] = str(ra_zr)
                        print(f"Menghitung RA ZR: {ra_zr}")
                except (ValueError, TypeError) as e:
                    print(f"Error menghitung RA ZR: {e}")
                        
                # Pastikan nilai W ada dan benar berdasarkan gambar 1
                try:
                    # Tambahkan debug untuk melihat semua kunci di row_data
                    print("DEBUG - Kunci di row_data:", list(row_data.keys()))
                    
                    # Gunakan nilai W yang sudah dimasukkan pengguna, tidak ada nilai default
                    if "W" in row_data:
                        print(f"DEBUG - Nilai W di row_data: '{row_data['W']}'")
                        
                        if row_data["W"] and row_data["W"].strip():
                            # Gunakan nilai asli tanpa konversi ke float
                            print(f"Menggunakan nilai W dari input: {row_data['W']}")
                        else:
                            # Jika tidak ada input, biarkan kosong sebagai string
                            print(f"Nilai W kosong, membiarkannya tetap kosong")
                    else:
                        # Jika tidak ada kunci W di row_data, tandai sebagai kosong
                        row_data["W"] = ""
                        print(f"Kunci W tidak ditemukan di row_data, mengatur ke kosong")
                except Exception as e:
                    print(f"Error mengambil nilai W: {e}")
                    # Tidak perlu mengatur ke kosong jika sudah ada nilai

                # KLASIFIKASI
                try:
                    if "IQ " in row_data:
                        iq = float(row_data.get("IQ ", 0))
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
                        
                        row_data["KLASIFIKASI"] = iq_klasifikasi
                        print(f"Menghitung KLASIFIKASI: {iq_klasifikasi}")
                except (ValueError, TypeError) as e:
                    print(f"Error menghitung KLASIFIKASI: {e}")
                
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
                    if col_name == "W":
                        # Cari secara langsung kolom W berdasarkan nama kolom
                        for possible_idx in range(1, sheet.max_column + 1):
                            cell_val = sheet.cell(row=header_row, column=possible_idx).value
                            if cell_val == "W":
                                col_idx = possible_idx
                                print(f"DEBUG - Menemukan kolom W di Excel pada indeks {col_idx}")
                                break
                                
                        # Jika ditemukan, simpan nilai langsung sebagai string
                        if col_idx:
                            print(f"DEBUG - Menyimpan nilai W='{value}' ke Excel di kolom {col_idx}")
                            target_cell = sheet.cell(row=target_row, column=col_idx)
                            # Gunakan nilai string langsung, jangan konversi
                            target_cell.value = value
                            # Lanjutkan ke kolom berikutnya
                            continue
                    elif col_name in special_columns:
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
                                    
                                    # Penanganan khusus untuk kolom deskripsi psikogram (kolom 55-66)
                                    # Kolom ini harus menggunakan nilai langsung jika sudah terisi daripada formula
                                    if col_idx >= 55 and col_idx <= 66 and value and str(value).strip():
                                        target_cell.value = value  # Gunakan nilai langsung dari UI
                                        print(f"Formula IFS untuk kolom {col_idx}: {formula}")
                                    else:
                                        # Gunakan formula untuk kolom lain
                                        target_cell.value = formula
                                        print(f"Formula IFS untuk kolom {col_idx}: {formula}")
                                        
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
                                    print(f"Formula normal untuk kolom {col_idx}: {formula}")
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
                                    # Cek jika ini kolom PAPIKOSTICK (kolom 16-36)
                                    if col_idx >= 16 and col_idx <= 36:
                                        # Untuk kolom PAPIKOSTICK, simpan nilai seperti apa adanya sebagai string
                                        print(f"DEBUG - Menyimpan nilai PAPIKOSTICK '{value}' di kolom {col_idx} sebagai integer jika memungkinkan")
                                        try:
                                            # Coba konversi ke integer jika nilai bulat
                                            num_value = float(value)
                                            if num_value.is_integer():
                                                target_cell.value = int(num_value)
                                            else:
                                                target_cell.value = num_value
                                        except (ValueError, TypeError):
                                            # Jika tidak bisa dikonversi, simpan sebagai string
                                            target_cell.value = value
                                    # Untuk kolom W, simpan sebagai string
                                    elif col_name == "W":
                                        print(f"DEBUG - Menyimpan nilai W '{value}' di kolom {col_idx} sebagai integer jika memungkinkan")
                                        try:
                                            # Coba konversi ke integer jika nilai bulat
                                            num_value = float(value)
                                            if num_value.is_integer():
                                                target_cell.value = int(num_value)
                                            else:
                                                target_cell.value = num_value
                                        except (ValueError, TypeError):
                                            # Jika tidak bisa dikonversi, simpan sebagai string
                                            target_cell.value = value
                                    # Untuk kolom C (Coding), simpan sebagai string
                                    elif col_name == "C (Coding)":
                                        print(f"DEBUG - Menyimpan nilai C (Coding) '{value}' di kolom {col_idx} sebagai integer jika memungkinkan")
                                        try:
                                            # Coba konversi ke integer jika nilai bulat
                                            num_value = float(value)
                                            if num_value.is_integer():
                                                target_cell.value = int(num_value)
                                            else:
                                                target_cell.value = num_value
                                        except (ValueError, TypeError):
                                            # Jika tidak bisa dikonversi, simpan sebagai string
                                            target_cell.value = value
                                    # Untuk kolom lain, konversi ke angka
                                    else:
                                        # Coba konversi ke integer jika nilai merupakan angka bulat
                                        try:
                                            if float(value).is_integer():
                                                target_cell.value = int(float(value))
                                            else:
                                                # Jika ada desimal, tetap gunakan float
                                                target_cell.value = float(value)
                                        except ValueError:
                                            target_cell.value = value
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
                # Terapkan formula-formula langsung ke sheet
                self.set_formulas_direct(sheet, header_row)
                
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
        try:
            return self.columns.index(column_name)
        except ValueError:
            # If exact match not found, try case insensitive match
            for i, col in enumerate(self.columns):
                if col.lower() == column_name.lower():
                    return i
                # For special case like "IQ" and "IQ " (with space)
                if column_name == "IQ" and col == "IQ ":
                    return i
                if column_name == "IQ " and col == "IQ":
                    return i
            # If still not found, try to find by partial match
            for i, col in enumerate(self.columns):
                if column_name.lower() in col.lower():
                    return i
            # If nothing found
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

            # Get absolute path for logo
            import os
            import base64
            from datetime import datetime
            current_dir = os.path.dirname(os.path.abspath(__file__))
            logo_path = os.path.join(current_dir, "logo.png")
            
            # Baca file logo dan encode ke base64
            try:
                # Coba cara pertama dengan path relatif
                if os.path.exists(logo_path):
                    with open(logo_path, "rb") as image_file:
                        logo_base64 = base64.b64encode(image_file.read()).decode('utf-8')
                        # Konversi ke data URL untuk digunakan dalam HTML dengan format PNG
                        logo_data_url = f"data:image/png;base64,{logo_base64}"
                        print(f"Logo encoded successfully. Path: {logo_path}")
                else:
                    # Coba dengan direktori saat ini
                    logo_path_alt = "logo.png"
                    if os.path.exists(logo_path_alt):
                        with open(logo_path_alt, "rb") as image_file:
                            logo_base64 = base64.b64encode(image_file.read()).decode('utf-8')
                            logo_data_url = f"data:image/png;base64,{logo_base64}"
                            print(f"Logo encoded successfully using alternative path: {logo_path_alt}")
                    else:
                        # Jika masih gagal, periksa direktori di atas
                        logo_path_parent = os.path.join(os.path.dirname(current_dir), "logo.png")
                        if os.path.exists(logo_path_parent):
                            with open(logo_path_parent, "rb") as image_file:
                                logo_base64 = base64.b64encode(image_file.read()).decode('utf-8')
                                logo_data_url = f"data:image/png;base64,{logo_base64}"
                                print(f"Logo encoded successfully using parent directory: {logo_path_parent}")
                        else:
                            # Fallback jika semua upaya gagal
                            print(f"Logo not found in any of these locations: {logo_path}, {logo_path_alt}, {logo_path_parent}")
                            logo_data_url = ""
            except Exception as e:
                print(f"Error encoding logo: {e}")
                # Fallback jika logo tidak bisa dibaca
                logo_data_url = ""
            
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
                    <div style="width: 150px;">
                        <img src="{logo_data_url}" alt="Logo" style="width: 100%; height: auto;">
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
                        <div style="width: 150px;">
                            <img src="{logo_data_url}" alt="Logo" style="width: 100%; height: auto;">
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
                                {row_data.get('Intelegensi Umum.1', '')}
                            </td>
                        </tr>
                        <tr>
                            <td style="padding: 8px; vertical-align: top; border: 1px solid black; font-weight: bold;">SIKAP DAN CARA KERJA</td>
                            <td style="padding: 8px; text-align: justify; border: 1px solid black;">
                                {row_data.get('Sistematika Kerja/ cd.1', '')}
                            </td>
                        </tr>
                        <tr>
                            <td style="padding: 8px; vertical-align: top; border: 1px solid black; font-weight: bold;">KEPRIBADIAN</td>
                            <td style="padding: 8px; text-align: justify; border: 1px solid black;">
                                {row_data.get('Stabilitas Emosi / E.1', '')}
                            </td>
                        </tr>
                        <tr>
                            <td style="padding: 8px; vertical-align: top; border: 1px solid black; font-weight: bold;">KEMAMPUAN BELAJAR</td>
                            <td style="padding: 8px; text-align: justify; border: 1px solid black;">
                                {row_data.get('Fleksibilitas', '')}
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" style="padding: 8px;">
                                <div style="font-weight: bold; margin: 10px 0; font-size: 12px;">PENGEMBANGAN</div>
                                <div style="text-align: justify; border: 1px solid black; padding: 8px;">
                                    {row_data.get('Keterampilan Sosial / X S', '')}
                                </div>
                            </td>
                        </tr>
                    </table>

                    <table class="psikogram" style="width: 100%; border-collapse: collapse; margin-top: 20px; font-family: Arial, sans-serif;">
                        <tr>
                            <th colspan="2" style="text-align: center; padding: 8px; background-color: #fbe4d5; border: 1px solid black;">Kategori Hasil Screening</th>
                        </tr>
                        <tr>
                            <td style="width: 5%; text-align: center; border: 1px solid black; padding: 8px;">{"X" if float(row_data.get('IQ ', 0)) >= 90 else ""}</td>
                            <td style="padding: 8px; border: 1px solid black;">Tahapan Normal<br><span style="font-size: 10px; color: #666;">Individu menunjukkan adaptasi gejala gangguan mental yang mengganggu fungsi sehari-hari</span></td>
                        </tr>
                        <tr>
                            <td style="text-align: center; border: 1px solid black; padding: 8px;">{"X" if 80 <= float(row_data.get('IQ ', 0)) < 90 else ""}</td>
                            <td style="padding: 8px; border: 1px solid black;">Kecenderungan Stress dalam Tekanan<br><span style="font-size: 10px; color: #666;">Dalam situasi yg menimbulkan tekanan dapat berdampak pada kondisi individu & respon emosional yg ditampilkan</span></td>
                        </tr>
                        <tr>
                            <td style="text-align: center; border: 1px solid black; padding: 8px;">{"X" if float(row_data.get('IQ ', 0)) < 80 else ""}</td>
                            <td style="padding: 8px; border: 1px solid black;">Gangguan<br><span style="font-size: 10px; color: #666;">Individu menunjukkan gejala-gejala gangguan yang dapat mengganggu fungsi sehari-hari</span></td>
                        </tr>
                    </table>
                    
                    <table class="psikogram" style="width: 100%; border-collapse: collapse; margin-top: 20px; font-family: Arial, sans-serif;">
                        <tr>
                            <th colspan="2" style="text-align: center; padding: 8px; background-color: #fbe4d5; border: 1px solid black;">Kesimpulan Keseluruhan</th>
                        </tr>
                        <tr>
                            <td style="width: 8%; text-align: center; border: 1px solid black; padding: 8px;">{"X" if float(row_data.get('IQ ', 0)) >= 110 else ""}</td>
                            <td style="padding: 8px; border: 1px solid black;">LAYAK DIREKOMENDASIKAN</td>
                        </tr>
                        <tr>
                            <td style="width: 8%; text-align: center; border: 1px solid black; padding: 8px;">{"X" if 90 <= float(row_data.get('IQ ', 0)) < 110 else ""}</td>
                            <td style="padding: 8px; border: 1px solid black;">LAYAK DIPERTIMBANGKAN</td>
                        </tr>
                        <tr>
                            <td style="text-align: center; border: 1px solid black; padding: 8px;">{"X" if float(row_data.get('IQ ', 0)) < 90 else ""}</td>
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
                        <img src="{logo_data_url}" alt="Logo" style="width: 150px; height: auto; margin-right: 20px;">
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
                                <span>: {datetime.now().strftime("%d %B %Y")}</span>
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
            
            if not file_name:
                return  # User canceled the save dialog
                
            # Tambahkan ekstensi .pdf jika tidak ada
            if not file_name.lower().endswith('.pdf'):
                file_name += '.pdf'
                
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
            
            # Variable to track if we've already shown a message
            self.message_shown = False
            
            def handle_pdf(data):
                if data and not self.message_shown:
                    self.message_shown = True
                    try:
                        with open(file_name, 'wb') as f:
                            f.write(data)
                        QMessageBox.information(self, "Success", "PDF saved successfully!")
                    except Exception as e:
                        QMessageBox.warning(self, "Warning", f"Failed to write PDF: {str(e)}")
                elif not data and not self.message_shown:
                    self.message_shown = True
                    QMessageBox.warning(self, "Warning", "Failed to generate PDF data")
            
            # Penting untuk tetap referensi ke web_view agar tidak di-garbage collect
            self.temp_web_view = web_view
            web_view.page().printToPdf(handle_pdf)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error saving PDF: {e}")
            print(f"Error saving PDF: {e}")
    
    def save_pdf_file(self, pdf_data, file_name):
        try:
            with open(file_name, 'wb') as f:
                f.write(pdf_data)
        except Exception as e:
            print(f"Error writing PDF file: {e}")

    def generate_common_formulas(self, row_number):
        """
        Menghasilkan formula-formula umum dengan nomor baris yang sesuai.
        """
        formulas = {
            'SDR/SDRI': f"=IF(F{row_number}=\"P\",\"Sdri.\",\"Sdr.\")",
            'Keterangan PHQ': f"=@IFS(I{row_number}<5,\"Tidak ada\",I{row_number}<10,\"Ringan\",I{row_number}<15,\"Sedang\",I{row_number}<20,\"Cukup Berat\",I{row_number}<28,\"Parah\")",
            'IQ ': f"=SUM(L{row_number}:P{row_number})/5",
            'Unnamed: 16': f"=(M{row_number}+O{row_number})/2",
            'KLASIFIKASI': f"=@IFS(K{row_number}<79,\"Rendah\",K{row_number}<90,\"Dibawah Rata-Rata\",K{row_number}<110,\"Rata-Rata\",K{row_number}<120,\"Diatas Rata-Rata\",K{row_number}>119,\"Superior\")",
            'C (Coding)': f"=@IFS(AE{row_number}=1,9,AE{row_number}=2,8,AE{row_number}=3,7,AE{row_number}=4,6,AE{row_number}=5,5,AE{row_number}=6,4,AE{row_number}=7,3,AE{row_number}=8,2,AE{row_number}=9,1)",
            'NG': f"=(S{row_number}+T{row_number})/2",
            'CDR': f"=(AE{row_number}+AG{row_number}+AH{row_number})/3",
            'TV': f"=(Y{row_number}+Z{row_number})/2",
            'PI': f"=(W{row_number}+X{row_number})/2",
            'BS': f"=(AA{row_number}+AB{row_number})/2",
            'ZK': f"=(AI{row_number}+AK{row_number})/2",
        }
        return formulas

    def set_formulas_direct(self, sheet, header_row):
        """
        Fungsi untuk mengatur formula Excel secara langsung pada sheet
        """
        try:
            # Looping untuk setiap baris data (di bawah header)
            for row_idx in range(header_row + 1, sheet.max_row + 1):
                # Formula untuk SDR/SDRI
                sdr_sdri_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("SDR/SDRI"))
                sdr_sdri_cell.value = f'=IF(F{row_idx}="P","Sdri.","Sdr.")'
                
                # Formula untuk Keterangan PHQ
                ket_phq_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("Keterangan PHQ"))
                ket_phq_cell.value = f'=IFS(I{row_idx}<5,"Tidak ada",I{row_idx}<10,"Ringan",I{row_idx}<15,"Sedang",I{row_idx}<20,"Cukup Berat",I{row_idx}<28,"Parah")'
                
                # Formula untuk IQ
                iq_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("IQ "))
                iq_cell.value = f'=SUM(L{row_idx}:P{row_idx})/5'
                
                # Formula untuk Unnamed: 16
                unnamed_16_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("Unnamed: 16"))
                unnamed_16_cell.value = f'=(M{row_idx}+O{row_idx})/2'
                
                # Formula untuk KLASIFIKASI
                klasifikasi_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("KLASIFIKASI"))
                klasifikasi_cell.value = f'=IFS(K{row_idx}<79,"Rendah",K{row_idx}<90,"Dibawah Rata-Rata",K{row_idx}<110,"Rata-Rata",K{row_idx}<120,"Diatas Rata-Rata",K{row_idx}>119,"Superior")'
                
                # Formula untuk C (Coding)
                c_coding_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("C (Coding)"))
                c_coding_cell.value = f'=IFS(AE{row_idx}=1,9,AE{row_idx}=2,8,AE{row_idx}=3,7,AE{row_idx}=4,6,AE{row_idx}=5,5,AE{row_idx}=6,4,AE{row_idx}=7,3,AE{row_idx}=8,2,AE{row_idx}=9,1)'
                
                # Formula untuk NG
                ng_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("NG"))
                ng_cell.value = f'=(S{row_idx}+T{row_idx})/2'
                
                # Formula untuk CDR
                cdr_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("CDR"))
                cdr_cell.value = f'=(AE{row_idx}+AG{row_idx}+AH{row_idx})/3'
                
                # Formula untuk TV
                tv_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("TV"))
                tv_cell.value = f'=(Y{row_idx}+Z{row_idx})/2'
                
                # Formula untuk PI
                pi_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("PI"))
                pi_cell.value = f'=(W{row_idx}+X{row_idx})/2'
                
                # Formula untuk BS
                bs_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("BS"))
                bs_cell.value = f'=(AA{row_idx}+AB{row_idx})/2'
                
                # Formula untuk ZK
                zk_cell = sheet.cell(row=row_idx, column=self.get_excel_column_index("ZK"))
                zk_cell.value = f'=(AI{row_idx}+AK{row_idx})/2'
            
            print("Formula berhasil diaplikasikan langsung ke sheet Excel")
            return True
        except Exception as e:
            print(f"Error saat menerapkan formula: {e}")
            return False
    
    def get_excel_column_index(self, column_name):
        """
        Mendapatkan indeks kolom Excel (1-based) dari nama kolom
        """
        try:
            # Cari indeks kolom (0-based)
            idx = self.get_column_index(column_name)
            # Konversi ke indeks Excel (1-based)
            return idx + 1 if idx >= 0 else None
        except Exception as e:
            print(f"Error mendapatkan indeks kolom Excel: {e}")
            return None

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelViewerApp()
    window.show()
    sys.exit(app.exec_())
