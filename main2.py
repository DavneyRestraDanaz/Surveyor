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

        # Add button for adding/updating data
        self.btn_add_update = QPushButton("Tambah / Edit Data")
        self.btn_add_update.setFont(QFont("Arial", 10))
        self.btn_add_update.setFixedHeight(35)
        self.btn_add_update.setFixedWidth(200) # Set fixed width for the button
        self.btn_add_update.clicked.connect(self.add_or_update_row)
        self.btn_add_update.setEnabled(False) # Initially disabled
        
        # Right-align the button in the layout
        button_container = QWidget()
        button_layout = QHBoxLayout(button_container)
        button_layout.addStretch() # Push button to right
        button_layout.addWidget(self.btn_add_update)
        button_layout.setContentsMargins(0, 0, 20, 0) # Add right margin
        main_layout.addWidget(button_container)
        
        # Enable button when file is loaded
        self.btn_select.clicked.connect(lambda: self.btn_add_update.setEnabled(True))

        # Add toggle buttons for each group
        toggle_layout = QHBoxLayout()
        toggle_personal = QPushButton("Toggle Personal Information")
        toggle_personal.setCheckable(True)
        toggle_personal.setChecked(False)  # Initially unchecked
        toggle_personal.setFont(QFont("Arial", 10))
        toggle_personal.setFixedHeight(35)
        toggle_personal.toggled.connect(lambda checked: personal_group.setVisible(checked))
        toggle_layout.addWidget(toggle_personal)
        
        toggle_ist = QPushButton("Toggle IST")
        toggle_ist.setCheckable(True)
        toggle_ist.setChecked(False)  # Initially unchecked
        toggle_ist.setFont(QFont("Arial", 10))
        toggle_ist.setFixedHeight(35)
        toggle_ist.toggled.connect(lambda checked: ist_group.setVisible(checked))
        toggle_layout.addWidget(toggle_ist)
        
        toggle_papikostick = QPushButton("Toggle PAPIKOSTICK")
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

        # Add Preview PDF button
        self.btn_preview_pdf = QPushButton("Preview PDF")
        self.btn_preview_pdf.setFont(QFont("Arial", 10))
        self.btn_preview_pdf.setFixedHeight(35)
        self.btn_preview_pdf.clicked.connect(self.preview_pdf)
        button_layout.addWidget(self.btn_preview_pdf)

        self.btn_save_excel = QPushButton("Simpan Perubahan ke Excel")
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
            
        search_text = self.search_input.text().lower()
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
            iq_col = self.get_column_index("IQ")
            nama_col = self.get_column_index("Nama Peserta")
            tgl_lahir_col = self.get_column_index("TGL Lahir")

            # Get data from selected row
            iq_val = self.table.item(selected_row, iq_col)
            iq_value = iq_val.text() if iq_val else "0"  # Simplified to directly use the text value
            
            nama_val = self.table.item(selected_row, nama_col)
            nama = nama_val.text() if nama_val else ""
            tgl_lahir_val = self.table.item(selected_row, tgl_lahir_col)
            tgl_lahir = tgl_lahir_val.text() if tgl_lahir_val else ""

            # For debugging
            print(f"IQ Value from table: {iq_value}")            
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
            <body>
                <div class="header">
                    <div class="title">HASIL PEMERIKSAAN PSIKOLOGIS</div>
                    <div>(Asesmen Intelegensi, Kepribadian dan Minat)</div>
                </div>
            """

            # Add personal info
            html_content += f"""
            <div style="width: 100%; margin: 0 auto;">
                <table class="info-table">
                <tr>
                    <td width="20%">NAMA</td>
                    <td width="30%">: {nama}</td>
                    <td width="20%">PERUSAHAAN</td>
                    <td width="30%">: PT. BAM</td>
                </tr>
                <tr>
                    <td>TANGGAL LAHIR</td>
                    <td>: {tgl_lahir}</td>
                    <td>TANGGAL TES</td>
                    <td>: 26 Februari 2025</td>
                </tr>
                <tr>
                    <td>PEMERIKSA</td>
                    <td>: Chitra Ananda Mulia, M.Psi., Psikolog</td>
                    <td>LEMBAGA</td>
                    <td>: BEHAVYOURS</td>
                </tr>
                <tr>
                    <td>ALAMAT LEMBAGA</td>
                    <td colspan="3">: Jl. Patal Senayan No.01</td>
                </tr>
            </table>
            """
                
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
                                                        <td style="padding: 8px; text-align: center;">{iq_value}</td>
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
                                            <td style="border-left: 1px solid black; border-right: 1px solid black; padding: 4px; text-align: center;">
                                                Rendah
                                            </td>
                                            <td style="border-left: 1px solid black; border-right: 1px solid black; padding: 4px; text-align: center;">
                                                Dibawah<br>Rata-Rata
                                            </td>
                                            <td style="border-left: 1px solid black; border-right: 1px solid black; padding: 4px; text-align: center;">
                                                Rata-Rata
                                            </td>
                                            <td style="border-left: 1px solid black; border-right: 1px solid black; padding: 4px; text-align: center;">
                                                Diatas Rata-<br>Rata
                                            </td>
                                            <td style="border-left: 1px solid black; border-right: 1px solid black; padding: 4px; text-align: center;">
                                                Superior
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="border-left: 1px solid black; border-right: 1px solid black; border-bottom: 1px solid black; padding: 4px; text-align: center;">&lt; 79</td>
                                            <td style="border-left: 1px solid black; border-right: 1px solid black; border-bottom: 1px solid black; padding: 4px; text-align: center;">80 - 89</td>
                                            <td style="border-left: 1px solid black; border-right: 1px solid black; border-bottom: 1px solid black; padding: 4px; text-align: center;">90 - 109</td>
                                            <td style="border-left: 1px solid black; border-right: 1px solid black; border-bottom: 1px solid black; padding: 4px; text-align: center;">110 - 119</td>
                                            <td style="border-left: 1px solid black; border-right: 1px solid black; border-bottom: 1px solid black; padding: 4px; text-align: center;">&gt; 120</td>
                                        </tr>
                                """

           # Add IQ data
            iq_val = self.table.item(selected_row, self.get_column_index("IQ"))
            iq_class = self.table.item(selected_row, self.get_column_index("KLASIFIKASI"))
            if iq_val and iq_class:
                    html_content += f"""
                        <tr>
                            <td style="border-right: 1px solid black; padding: 8px; text-align: center;">{'X' if iq_class.text() == 'Rendah' else ''}</td>
                            <td style="border-right: 1px solid black; padding: 8px; text-align: center;">{'X' if iq_class.text() == 'Di Bawah Rata-rata' else ''}</td>
                            <td style="border-right: 1px solid black; padding: 8px; text-align: center;">{'X' if iq_class.text() == 'Rata-rata' else ''}</td>
                            <td style="border-right: 1px solid black; padding: 8px; text-align: center;">{'X' if iq_class.text() == 'Di Atas Rata-rata' else ''}</td>
                            <td style="padding: 8px; text-align: center;">{'X' if iq_class.text() == 'Superior' else ''}</td>
                        </tr>
                    """

                    html_content += """
                            </table>
                        </td>
                    </tr>
                </table>
                
                <table class="psikogram" style="width: 100%; margin-top: 20px; border-collapse: collapse; border: 1px solid black;">
                    <tr>
                        <th colspan="8" style="text-align: center; padding: 8px; background-color: #deeaf6; border: 1px solid black;">PSIKOGRAM</th>
                    </tr>
                    <tr>
                        <th style="width: 5%; border: 1px solid black; padding: 8px; background-color: #deeaf6;">NO</th>
                        <th style="width: 15%; border: 1px solid black; padding: 8px; background-color: #deeaf6;">ASPEK<br>PSIKOLOGIS</th>
                        <th style="width: 40%; border: 1px solid black; padding: 8px; background-color: #deeaf6;">DEFINISI</th>
                        <th style="width: 8%; border: 1px solid black; text-align: center; padding: 8px; background-color: #deeaf6;">R</th>
                        <th style="width: 8%; border: 1px solid black; text-align: center; padding: 8px; background-color: #deeaf6;">K</th>
                        <th style="width: 8%; border: 1px solid black; text-align: center; padding: 8px; background-color: #deeaf6;">C</th>
                        <th style="width: 8%; border: 1px solid black; text-align: center; padding: 8px; background-color: #deeaf6;">B</th>
                        <th style="width: 8%; border: 1px solid black; text-align: center; padding: 8px; background-color: #deeaf6;">T</th>
                    </tr>

                    <tr>
                        <td colspan="8" style="background-color: #fbe4d5; text-align: center; border: 1px solid black;">KEMAMPUAN INTELEKTUAL</td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">1.</td>
                        <td>Logika Berpikir</td>
                        <td>Kemampuan untuk berpikir secara logis dan sistematis.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">2.</td>
                        <td>Daya Analisa</td>
                        <td>Kemampuan untuk melihat permasalahan dan memahami hubungan sebab akibat permasalahan.</td>
                        <td style="text-align: center;">X</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">3.</td>
                        <td>Kemampuan Numerikal</td>
                        <td>Kemampuan untuk berpikir praktis dalam memahami konsep angka dan hitungan.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">4.</td>
                        <td>Kemampuan Verbal</td>
                        <td>Kemampuan untuk memahami konsep dan pola dalam bentuk kata dan mengekspresikan gagasan secara verbal.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                        <td style="text-align: center;"></td>
                    </tr>
                    <tr>
                        <td colspan="8" style="background-color: #fbe4d5; text-align: center;">SIKAP DAN CARA KERJA</td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">5.</td>
                        <td>Orientasi Hasil</td>
                        <td>Kemampuan untuk mempertahankan komitmen untuk menyelesaikan tugas secara bertanggung jawab dan memperhatikan keterhubungan antara perencanaan dan hasil kerja.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                        <td style="text-align: center;"></td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">6.</td>
                        <td>Fleksibilitas</td>
                        <td>Kemampuan untuk menyesuaikan diri dalam menghadapi permasalahan.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">7.</td>
                        <td>Sistematika Kerja</td>
                        <td>Kemampuan untuk merencanakan hingga mengorganisasikan cara kerja dalam proses penyelesaian pekerjaannya.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                        <td style="text-align: center;"></td>
                    </tr>
                    <tr>
                        <td colspan="8" style="background-color: #fbe4d5; text-align: center;">KEPRIBADIAN</td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">8.</td>
                        <td>Motivasi Berprestasi</td>
                        <td>Kemampuan untuk menunjukkan prestasi dan mencapai target.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">9.</td>
                        <td>Kerjasama</td>
                        <td>Kemampuan untuk menjalin, membina dan mengoptimalkan hubungan kerja yang efektif demi tercapainya tujuan bersama.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">10.</td>
                        <td>Keterampilan Interpersonal</td>
                        <td>Kemampuan untuk menjalin hubungan sosial dan mampu memahami kebutuhan orang lain.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                        <td style="text-align: center;"></td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">11.</td>
                        <td>Stabilitas Emosi</td>
                        <td>Kemampuan untuk memahami dan mengontrol emosi.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                    </tr>
                    <tr>
                        <td colspan="8" style="background-color: #fbe4d5; text-align: center;">KEMAMPUAN BELAJAR</td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">12.</td>
                        <td>Pengembangan Diri</td>
                        <td>Kemampuan untuk meningkatkan pengetahuan dan menyempurnakan keterampilan diri.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                    </tr>
                    <tr>
                        <td style="text-align: center; background-color: #deeaf6;">13.</td>
                        <td>Mengelola Perubahan</td>
                        <td>Kemampuan dalam menyesuaikan diri dengan situasi yang baru.</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;">X</td>
                        <td style="text-align: center;"></td>
                        <td style="text-align: center;"></td>
                    </tr>
                    <tr style="border-top: 1px solid black;">
                        <td colspan="8" style="text-align: center; padding: 2px; font-family: Arial; font-size: 11px; background-color: #deeaf6;">
                            <div style="display: inline-block; width: 100%;">
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

            # Close tables and add footer
            html_content += """
                </table>
                <div class="footer">
                    Laporan ini bersifat confidential dan diketahui oleh Psikolog
                </div>
            </body>
            </html>
            """

            # Add page break and second page content
            html_content += f"""
                <div class="page-break"></div>
                <div class="page">
                    <table class="psikogram" style="width: 100%; border-collapse: collapse; margin-top: 20px;">
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

                    <table class="psikogram" style="width: 100%; border-collapse: collapse; margin-top: 20px;">
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
                    
                    <table class="psikogram" style="width: 100%; border-collapse: collapse; margin-top: 20px;">
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
                            <div style="position: absolute; bottom: 2cm; width: calc(100% - 4cm); text-align: center; font-style: italic;">
                        Laporan ini bersifat confidential dan diketahui oleh Psikolog
                    </div>
                </div>
            """

            # Add page break and third page content
            html_content += f"""
                <div class="page-break"></div>
                <div class="page" style="padding: 2cm; font-family: Arial;">
                    <div style="display: flex; align-items: center; margin-bottom: 20px;">
                        <img src="behanyours.png" alt="Logo" style="width: 80px; height: auto; margin-right: 20px;">
                        <div style="flex-grow: 1; text-align: center;">
                            <div style="font-size: 14px; font-weight: bold;">HASIL PEMERIKSAAN PSIKOLOGIS</div>
                            <div style="font-size: 12px;">(Asesmen Intelegensi, Kepribadian dan Minat)</div>
                        </div>
                        <div style="text-align: right; font-size: 12px;">
                            <div style="font-weight: bold;">RAHASIA</div>
                            <div>No. 158/02/JMI/2025</div>
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
                                <span>: <img src="signature.png" alt="Signature" style="height: 40px; vertical-align: middle;"></span>
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

                    <div style="position: absolute; bottom: 2cm; width: calc(100% - 4cm); text-align: center; font-style: italic;">
                        Laporan ini bersifat confidential dan diketahui oleh Psikolog
                    </div>
                </div>
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
            web_view1.setZoomFactor(0.8)
            web_view1.setFixedWidth(int(dialog_width * 0.3))  # Adjust width to 30% for 3 pages
            # Split HTML content at page break
            pages = html_content.split('<div class="page-break"></div>')
            web_view1.setHtml(pages[0])
            preview_layout.addWidget(web_view1)
            
            # Create web view for page 2
            web_view2 = QWebEngineView(preview_dialog)
            web_view2.setZoomFactor(0.8)
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
                            margin-top: 15px;
                            text-align: center;
                            font-style: italic;
                            font-size: 10px;
                        }
                        .legend-row td {
                            text-align: center;
                            padding: 2px;
                            font-size: 11px;
                            border: none;
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
