import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QFileDialog, QVBoxLayout, 
    QTableWidget, QTableWidgetItem, QLabel, QHBoxLayout, QLineEdit,
    QGridLayout, QGroupBox, QFormLayout, QHeaderView
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

class ExcelViewerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Aplikasi Laporan Psikologi")
        self.setGeometry(100, 100, 1400, 800)
        self.initUI()
        self.excel_file_path = ""
        self.df = pd.DataFrame()

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

        # Input Form Section
        input_group = QGroupBox("Input Data")
        input_group.setFont(QFont("Arial", 11, QFont.Bold))
        input_layout = QGridLayout()
        
        self.input_fields = []
        self.placeholders = ["No Tes", "Nama Peserta", "Tanggal Lahir", "Jenis Kelamin", 
                           "IQ", "WA", "GE", "RA", "ZR"]
        
        for i, placeholder in enumerate(self.placeholders):
            label = QLabel(placeholder + ":")
            label.setFont(QFont("Arial", 10))
            field = QLineEdit()
            field.setFixedHeight(30)
            self.input_fields.append(field)
            row = i // 3
            col = i % 3 * 2
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
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        table_layout.addWidget(self.table)

        # Buttons below table
        button_layout = QHBoxLayout()
        
        self.btn_delete = QPushButton("Hapus Baris Terpilih")
        self.btn_delete.setFont(QFont("Arial", 10))
        self.btn_delete.setFixedHeight(35)
        self.btn_delete.clicked.connect(self.delete_selected_row)
        button_layout.addWidget(self.btn_delete)

        self.btn_save_excel = QPushButton("Simpan Perubahan ke Excel")
        self.btn_save_excel.setFont(QFont("Arial", 10))
        self.btn_save_excel.setFixedHeight(35)
        self.btn_save_excel.clicked.connect(self.save_to_excel)
        button_layout.addWidget(self.btn_save_excel)
        
        table_layout.addLayout(button_layout)
        table_group.setLayout(table_layout)
        main_layout.addWidget(table_group)

        self.setLayout(main_layout)

    def load_excel(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Pilih File Excel", "", "Excel Files (*.xlsx);;All Files (*)", options=options)

        if file_path:
            self.label.setText(f"File: {file_path}")
            self.excel_file_path = file_path
            self.process_excel(file_path)

    def process_excel(self, file_path):
        df = pd.read_excel(file_path, engine='openpyxl', header=0, keep_default_na=False).fillna("")
        self.df = df
        self.show_table(df)

    def show_table(self, df):
        self.table.setRowCount(df.shape[0])
        self.table.setColumnCount(df.shape[1])
        self.table.setHorizontalHeaderLabels(df.columns)

        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                self.table.setItem(row, col, QTableWidgetItem(str(df.iat[row, col])))

        self.table.resizeColumnsToContents()

    def add_or_update_row(self):
        try:
            selected_row = self.table.currentRow()
            values = [field.text().strip() for field in self.input_fields]

            # Jika "No Tes" kosong, lewati input tanpa error
            if values[0] == "":
                values[0] = ""

            if selected_row >= 0:  # Jika ada baris yang dipilih, update data
                for col, value in enumerate(values):
                    if value:  # Jika input tidak kosong, update kolom terkait
                        self.table.setItem(selected_row, col, QTableWidgetItem(value))
                self.recalculate_values(selected_row)
            else:  # Jika tidak ada baris yang dipilih, tambahkan data baru
                row_position = self.table.rowCount()
                self.table.insertRow(row_position)

                for col, value in enumerate(values):
                    if value:
                        self.table.setItem(row_position, col, QTableWidgetItem(value))

                self.recalculate_values(row_position)

            for field in self.input_fields:
                field.clear()

            self.table.resizeColumnsToContents()
        except Exception as e:
            print(f"Terjadi kesalahan: {e}")

    def recalculate_values(self, row):
        try:
            iq = self.get_cell_value(row, 4)
            wa = self.get_cell_value(row, 5)
            ge = self.get_cell_value(row, 6)
            ra = self.get_cell_value(row, 7)
            zr = self.get_cell_value(row, 8)

            if wa is not None and ge is not None:
                wa_ge = (wa + ge) / 2
                self.table.setItem(row, 9, QTableWidgetItem(str(wa_ge)))

            if ra is not None and zr is not None:
                ra_zr = (ra + zr) / 2
                self.table.setItem(row, 10, QTableWidgetItem(str(ra_zr)))

            if iq is not None:
                if iq < 79:
                    iq_klasifikasi = "Rendah"
                elif 79 <= iq < 90:
                    iq_klasifikasi = "Di Bawah Rata-rata"
                elif 90 <= iq < 110:
                    iq_klasifikasi = "Rata-rata"
                elif 110 <= iq < 120:
                    iq_klasifikasi = "Di Atas Rata-rata"
                else:
                    iq_klasifikasi = "Superior"
                self.table.setItem(row, 11, QTableWidgetItem(iq_klasifikasi))
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

    def save_to_excel(self):
        if not self.excel_file_path:
            return

        data = []
        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)

        header_labels = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]

        if not data or len(data[0]) != len(header_labels):
            print("Error: Jumlah kolom tidak sesuai!")
            return

        df_new = pd.DataFrame(data, columns=header_labels)
        new_path = self.excel_file_path.replace(".xlsx", "_new.xlsx")
        df_new.to_excel(new_path, index=False, engine="openpyxl")
        print(f"Data berhasil disimpan ke {new_path}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelViewerApp()
    window.show()
    sys.exit(app.exec_())
