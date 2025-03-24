import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout, 
    QTableWidget, QTableWidgetItem, QLabel, QHBoxLayout, QLineEdit,
    QGridLayout, QGroupBox, QFormLayout, QHeaderView, QDialog, QRadioButton, QDialogButtonBox
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QCalendarWidget

class ExcelViewerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Aplikasi Laporan Psikologi")
        self.setGeometry(100, 100, 1400, 800)
        # Updated columns based on the images provided
        self.columns = [
            "No", "No Tes", "TGL Lahir", "JK", "Nama Peserta", 
            "IQ", "Konkrit Praktis", "Verbal", "Flexibilitas Pikir", 
            "Daya Abstraksi Verbal", "Berpikir Praktis", "Berpikir Teoritis", 
            "Memori", "WA GE", "RA ZR", "IQ KLASIFIKASI",
            "N", "G", "A", "L", "P", "I", "T", "V", "S", "B", "O", "X", 
            "C (Coding)", "D", "R", "Z", "E", "K", "F", "W", "CD", "TV", "BO", "SO", "BX"
        ]
        self.input_columns = self.columns.copy()
        self.initUI()
        self.excel_file_path = ""
        self.df = pd.DataFrame(columns=self.columns)

    def initUI(self):
        main_layout = QVBoxLayout()
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)

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

        # Add button for adding/updating data
        self.btn_add_update = QPushButton("Tambah / Edit Data")
        self.btn_add_update.setFont(QFont("Arial", 10))
        self.btn_add_update.setFixedHeight(35)
        self.btn_add_update.clicked.connect(self.add_or_update_row)
        main_layout.addWidget(self.btn_add_update)

        # Add toggle buttons for each group
        toggle_layout = QHBoxLayout()
        toggle_personal = QPushButton("Toggle Personal Information")
        toggle_personal.setCheckable(True)
        toggle_personal.setChecked(True)
        toggle_personal.setFont(QFont("Arial", 10))
        toggle_personal.setFixedHeight(35)
        toggle_personal.toggled.connect(lambda checked: personal_group.setVisible(checked))
        toggle_layout.addWidget(toggle_personal)
        
        toggle_ist = QPushButton("Toggle IST")
        toggle_ist.setCheckable(True)
        toggle_ist.setChecked(True)
        toggle_ist.setFont(QFont("Arial", 10))
        toggle_ist.setFixedHeight(35)
        toggle_ist.toggled.connect(lambda checked: ist_group.setVisible(checked))
        toggle_layout.addWidget(toggle_ist)
        
        toggle_papikostick = QPushButton("Toggle PAPIKOSTICK")
        toggle_papikostick.setCheckable(True)
        toggle_papikostick.setChecked(True)
        toggle_papikostick.setFont(QFont("Arial", 10))
        toggle_papikostick.setFixedHeight(35)
        toggle_papikostick.toggled.connect(lambda checked: papikostick_group.setVisible(checked))
        toggle_layout.addWidget(toggle_papikostick)
        
        main_layout.addLayout(toggle_layout)

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
        
        self.btn_delete = QPushButton("Hapus Baris Terpilih")
        self.btn_delete.setFont(QFont("Arial", 10))
        self.btn_delete.setFixedHeight(35)
        self.btn_delete.clicked.connect(self.delete_selected_row)
        button_layout.addWidget(self.btn_delete)

        self.btn_print = QPushButton("Print Data Terpilih")
        self.btn_print.setFont(QFont("Arial", 10))
        self.btn_print.setFixedHeight(35)
        self.btn_print.clicked.connect(self.print_selected_row)
        button_layout.addWidget(self.btn_print)

        self.btn_save_excel = QPushButton("Simpan Perubahan ke Excel")
        self.btn_save_excel.setFont(QFont("Arial", 10))
        self.btn_save_excel.setFixedHeight(35)
        self.btn_save_excel.clicked.connect(self.save_to_excel)
        button_layout.addWidget(self.btn_save_excel)
        
        
        table_layout.addLayout(button_layout)
        table_group.setLayout(table_layout)
        main_layout.addWidget(table_group)

        self.setLayout(main_layout)

    def add_toggle_button(self, layout, group, text):
        toggle_button = QPushButton(text)
        toggle_button.setCheckable(True)
        toggle_button.setChecked(True)
        toggle_button.setFont(QFont("Arial", 10))
        toggle_button.setFixedHeight(35)
        toggle_button.toggled.connect(lambda checked: group.setVisible(checked))
        layout.addWidget(toggle_button)

    # Input Form Section
        input_group = QGroupBox("Input Data")
        input_group.setFont(QFont("Arial", 11, QFont.Bold))
        input_layout = QGridLayout()  # Use QGridLayout for side-by-side arrangement
        
        self.input_fields = []
        self.placeholders = ["No", "No Tes", "TGL Lahir", "JK", "Nama Peserta", 
                             "IQ", "Konkrit Praktis", "Verbal", "Flexibilitas Pikir", 
                             "Daya Abstraksi Verbal", "Berpikir Praktis", "Berpikir Teoritis", 
                             "Memori", "N", "G", "A", "L", "P", "I", "T", "V", "S", "B", "O", "X", 
                             "C (Coding)", "D", "R", "Z", "E", "K", "F", "W", "CD", "TV", "BO", "SO", "BX"]
        
        for i, placeholder in enumerate(self.placeholders):
            label = QLabel(placeholder + ":")
            label.setFont(QFont("Arial", 10))
            field = QLineEdit() if placeholder != "TGL Lahir" else None
            if placeholder == "TGL Lahir":
                field = QPushButton("Pilih Tanggal")
                field.clicked.connect(self.show_calendar)
            field.setFixedHeight(30)
            self.input_fields.append(field)
            row = i // 3  # Arrange fields in rows of 3
            col = (i % 3) * 2
            input_layout.addWidget(label, row, col)
            input_layout.addWidget(field, row, col + 1)

        self.btn_add_update = QPushButton("Tambah / Edit Data")
        self.btn_add_update.setFont(QFont("Arial", 10))
        self.btn_add_update.setFixedHeight(35)
        self.btn_add_update.clicked.connect(self.add_or_update_row)
        input_layout.addWidget(self.btn_add_update, (len(self.placeholders) // 3) + 1, 0, 1, 6)
        
        input_group.setLayout(input_layout)
        main_layout.addWidget(input_group)

        # Table Section
        table_group = QGroupBox("Data Hasil")
        table_group.setFont(QFont("Arial", 11, QFont.Bold))
        table_layout = QVBoxLayout()
        
        self.table = QTableWidget()
        self.table.setFont(QFont("Arial", 9))
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # Stretch columns to fit
        table_layout.addWidget(self.table)

        # Buttons below table
        button_layout = QHBoxLayout()
        
        self.btn_delete = QPushButton("Hapus Baris Terpilih")
        self.btn_delete.setFont(QFont("Arial", 10))
        self.btn_delete.setFixedHeight(35)
        self.btn_delete.clicked.connect(self.delete_selected_row)
        button_layout.addWidget(self.btn_delete)

        self.btn_print = QPushButton("Print Data Terpilih")
        self.btn_print.setFont(QFont("Arial", 10))
        self.btn_print.setFixedHeight(35)
        self.btn_print.clicked.connect(self.print_selected_row)
        button_layout.addWidget(self.btn_print)

        self.btn_save_excel = QPushButton("Simpan Perubahan ke Excel")
        self.btn_save_excel.setFont(QFont("Arial", 10))
        self.btn_save_excel.setFixedHeight(35)
        self.btn_save_excel.clicked.connect(self.save_to_excel)
        button_layout.addWidget(self.btn_save_excel)
        
        table_layout.addLayout(button_layout)
        table_group.setLayout(table_layout)
        main_layout.addWidget(table_group)

        self.setLayout(main_layout)
    def populate_fields_from_selection(self):
        selected_row = self.table.currentRow()
        if selected_row >= 0:
            row_data = {}
            for col, column_name in enumerate(self.columns):
                item = self.table.item(selected_row, col)
                row_data[column_name] = item.text() if item else ""

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

        if file_path:
            self.label.setText(f"File: {file_path}")
            self.excel_file_path = file_path
            
            # **🔹 Membaca seluruh sheet dalam file Excel**
            self.excel_data = pd.read_excel(file_path, sheet_name=None)  # Baca semua sheet
            
            # **🔹 Pastikan Sheet1 dan Sheet2 terbaca**
            if "Sheet1" in self.excel_data:
                self.sheet1_data = self.excel_data["Sheet1"]
            else:
                print("Sheet1 tidak ditemukan!")

            if "Sheet2" in self.excel_data:
                self.sheet2_data = self.excel_data["Sheet2"]
            else:
                print("Sheet2 tidak ditemukan!")

            # **🔹 Proses data setelah membaca**
            self.process_excel(file_path)

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

            if "Sheet2" in sheets:
                df_sheet2 = sheets["Sheet2"]
            else:
                print("Sheet2 tidak ditemukan!")
                df_sheet2 = None

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

            # 🔹 Proses Sheet2 (jika ada)
            if df_sheet2 is not None:
                print("Original columns (Sheet2):", df_sheet2.columns.tolist())

                # Simpan data Sheet2 tanpa pemrosesan tambahan
                self.df_sheet2 = df_sheet2.fillna("")
        
        except Exception as e:
            print(f"Error loading Excel file: {e}")
            import traceback
            traceback.print_exc()


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

    def add_or_update_row(self):
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



        # Create a new row for the table
        row = self.table.rowCount()
        self.table.insertRow(row)
        
        # Add values to the table
        for col, column_name in enumerate(self.columns):
            value = row_data.get(column_name, "")
            self.table.setItem(row, col, QTableWidgetItem(str(value)))
        
        # Update the DataFrame
        self.df.loc[len(self.df)] = row_data
        
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
            data = []
            for row in range(self.table.rowCount()):
                row_data = {}
                for col, column_name in enumerate(self.columns):
                    item = self.table.item(row, col)
                    row_data[column_name] = item.text() if item else ""
                data.append(row_data)

            df_new = pd.DataFrame(data, columns=self.columns)
            new_path = self.excel_file_path.replace(".xlsx", "_new.xlsx")
            df_new.to_excel(new_path, index=False, engine="openpyxl")
            print(f"Data berhasil disimpan ke {new_path}")
        except Exception as e:
            print(f"Error saving to Excel: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelViewerApp()
    window.show()
    sys.exit(app.exec_())
