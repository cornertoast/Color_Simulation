from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QHBoxLayout, QGridLayout, \
    QFormLayout, QLineEdit, QTabWidget, QTableWidgetItem, QTableWidget, QSizePolicy, QFrame, \
    QPushButton, QAbstractItemView,QComboBox,QPushButton,QCheckBox,QDialog,QFileDialog
from PySide6.QtGui import QKeyEvent,QColor,QPalette
from PySide6.QtCore import Qt
from PySide6.QtCharts import QChart
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import sqlite3
from Setting import *
import chardet
import openpyxl

class Light_Source_BLU_button(QWidget):
    def __init__(self):
        super().__init__()

        # 實例化
        self.LightSourceBLUTable = Light_Source_BLU_Table()
        self.import_data_button = QPushButton("Import_data")
        self.export_data_button = QPushButton("Export_data")

        self.Light_Source_BLU_button_layout = QHBoxLayout()
        self.Light_Source_BLU_button_layout.addWidget(self.import_data_button)
        self.Light_Source_BLU_button_layout.addWidget(self.export_data_button)

        self.setLayout(self.Light_Source_BLU_button_layout)

        # 連接功能
        # 連接匯入按鈕的槽函數
        self.import_data_button.clicked.connect(self.LightSourceBLUTable.import_data)
class Light_Source_BLU_Table(QTableWidget):
    def __init__(self):
        super().__init__()

        self.setColumnCount(16)
        self.setHorizontalHeaderLabels(["波長", "項目"])
        # 添加初始的行
        self.setRowCount(430)


        # 設定默認值
        for row, i in enumerate(range(380,801)):
            item = QTableWidgetItem(str(i))
            item.setBackground(QColor(173, 216, 230))  # 設置背景顏色為淺藍色
            item.setTextAlignment(Qt.AlignCenter)  # 設置文本居中對齊
            self.setItem(row, 0, item)


        # 設置表格可編輯
        self.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.SelectedClicked)

        # Set Background
        self.setStyleSheet("background-color: lightblue;")

    def import_data(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly

        file_name, _ = QFileDialog.getOpenFileName(
            self, "選擇檔案", "", "Excel Files (*.xlsx *.xls);;Text Files (*.txt);;All Files (*)",
            options=options
        )

        if file_name:
            # 讀取 Excel 檔案
            wb = openpyxl.load_workbook(file_name)
            sheet = wb.active

            # 設定表格列數
            current_row = self.rowCount()

            # 從第一列開始讀取數據
            for col_num, column in enumerate(sheet.iter_cols(min_row=1, max_row=402, values_only=True)):
                # 取得名稱和數值
                item_name = column[0]
                data_values = column[1:]

                # 將名稱放入表格的 header
                header_item = QTableWidgetItem(str(item_name))
                self.setVerticalHeaderItem(col_num, header_item)

                # 將數值放入表格
                for row, value in enumerate(data_values):
                    # 檢查是否需要新增行
                    if row >= current_row:
                        self.setRowCount(current_row + 1)

                    value_item = QTableWidgetItem(str(value))
                    self.setItem(row, col_num, value_item)

            # 創建 SQLite 數據庫，保存項目名稱和頻譜值
            #self.create_database(lines)

    def create_database(self, data):
        conn = sqlite3.connect('your_database.db')
        cursor = conn.cursor()

        # 創建表
        cursor.execute('''CREATE TABLE IF NOT EXISTS spectrum_data
                          (wavelength INTEGER, item TEXT)''')

        # 插入數據
        for line in data:
            wavelength, item = line.strip().split('\t')
            cursor.execute("INSERT INTO spectrum_data VALUES (?, ?)", (int(wavelength), item))

        conn.commit()
        conn.close()