from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QHBoxLayout, QGridLayout, \
    QFormLayout, QLineEdit, QTabWidget, QTableWidgetItem, QTableWidget, QSizePolicy, QFrame, \
    QPushButton, QAbstractItemView, QComboBox, QPushButton, QCheckBox
from PySide6.QtGui import QKeyEvent, QColor, QPalette
from PySide6.QtCore import Qt
from PySide6.QtCharts import QChart
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import sqlite3
from Setting import *
from Light_Source_BLU import *
import pandas as pd


class Color_Simulation(QWidget):
    def __init__(self):
        super().__init__()

        # 創建實例
        # color_result_table = Color_Result_Table()
        color_enter = Color_Enter()
        color_select = Color_Select()

        # layout
        Color_Simulation_layout = QVBoxLayout()
        # Color_Simulation_layout.addWidget(color_result_table)
        Color_Simulation_layout.addWidget(color_enter)
        Color_Simulation_layout.addWidget(color_select)

        # 在 Color_Simulation_layout 中，設置行列的伸展因子(空白的間距)
        Color_Simulation_layout.setStretch(0, 2)  # 第一行伸展因子為1
        Color_Simulation_layout.setStretch(1, 1)  # 第二行伸展因子为2
        Color_Simulation_layout.setStretch(2, 1)  # 第三行伸展因子为1

        # 放置latout
        self.setLayout(Color_Simulation_layout)


# class Color_Result_Table(QTableWidget):
#     def __init__(self):
#         super().__init__()
#
#
#         self.setColumnCount(16)
#         self.setHorizontalHeaderLabels(["項目", "Light", "背光名稱", "背光名稱", "CF/T", "R-CF 名稱",
#                                         "厚度", "CF/T", "G-CF名稱", "厚度", "CF/T", "B-CF名稱",
#                                         "厚度", "NTSC%", "BLU", "BLU"])
#         # 添加初始的行
#         self.setRowCount(3)
#
#         # 設定默認值
#         column1_default_values = ["色度", "x", "y", "Y", "x", "y", "Y", "x", "y", "Y",
#                                   "x", "y", "Y", "NTSC%", "x", "y"]
#         for row, values in enumerate([column1_default_values, ["整體模擬"], ["C-light"]]):
#             for column, value in enumerate(values):
#                 item = QTableWidgetItem(value)
#                 item.setBackground(QColor(173, 216, 230))  # 設置背景顏色為淺藍色
#                 item.setTextAlignment(Qt.AlignCenter)  # 設置文本居中對齊
#                 self.setItem(row, column, item)
#
#         # 設定特定單元格顏色
#         color_ranges = [(1, 3), (4, 6), (7, 9), (10, 12)]
#         colors = [QColor(RESULTWHITE), QColor(RESULTRED), QColor(RESULTGREEN), QColor(RESULTBLUE)]
#
#         for row in range(self.rowCount()):
#             for (start_col, end_col), color in zip(color_ranges, colors):
#                 for col in range(start_col, end_col + 1):
#                     item = self.item(row, col)
#                     if item:
#                         item.setBackground(color)
#
#         # 設置表格可編輯
#         self.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.SelectedClicked)
#
#         # Set Background
#         self.setStyleSheet("background-color: lightblue;")


class Color_Enter(QWidget):
    def __init__(self):
        super().__init__()

        # color_table區
        self.color_table = QTableWidget()
        self.color_table.setColumnCount(16)
        self.color_table.setHorizontalHeaderLabels(["項目", "Light", "背光名稱", "背光名稱", "CF/T", "R-CF 名稱",
                                        "厚度", "CF/T", "G-CF名稱", "厚度", "CF/T", "B-CF名稱",
                                        "厚度", "NTSC%", "BLU", "BLU"])
        # 添加初始的行
        self.color_table.setRowCount(3)

        # 設定默認值
        column1_default_values = ["色度", "x", "y", "Y", "x", "y", "Y", "x", "y", "Y",
                                  "x", "y", "Y", "NTSC%", "x", "y"]
        for row, values in enumerate([column1_default_values, ["整體模擬"], ["C-light"]]):
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setBackground(QColor(173, 216, 230))  # 設置背景顏色為淺藍色
                item.setTextAlignment(Qt.AlignCenter)  # 設置文本居中對齊
                self.color_table.setItem(row, column, item)

        # 設定特定單元格顏色
        color_ranges = [(1, 3), (4, 6), (7, 9), (10, 12)]
        colors = [QColor(RESULTWHITE), QColor(RESULTRED), QColor(RESULTGREEN), QColor(RESULTBLUE)]

        for row in range(self.color_table.rowCount()):
            for (start_col, end_col), color in zip(color_ranges, colors):
                for col in range(start_col, end_col + 1):
                    item = self.color_table.item(row, col)
                    if item:
                        item.setBackground(color)

        # 設置表格可編輯
        self.color_table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.SelectedClicked)

        # Set Background
        self.setStyleSheet("background-color: lightblue;")



        # 總Layout
        self.Color_Enter_layout = QGridLayout()

        # Light source區域----------------------------------------------------------------------------------------------
        self.light_source_mode = QComboBox()
        light_source_mode_box_items = ["未選", "自訂", "模擬", "替換"]
        for item in light_source_mode_box_items:
            self.light_source_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.light_source_mode.setStyleSheet(QCOMBOXMODESETTING)
        # self.light_source_box.setBackgroundRole(QPalette.Window)

        self.light_source = QComboBox()
        self.light_source_datatable = QComboBox()
        self.light_source_datatable.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 更新light_source_table
        self.update_light_source_datatable()
        # 觸發table更新,lightsource表單
        self.updateLightSourceComboBox()# 初始化
        self.light_source_datatable.currentIndexChanged.connect(self.updateLightSourceComboBox)
        self.light_source.setStyleSheet(QCOMBOXSETTING)
        # 初始时更新 Light_Source ComboBox 的选项
        # self.updateLightSourceComboBox()

        # 连接到信号，当 Light_Source_BLU 类的表格选择变更时调用 updateLightSourceComboBox 方法
        # self.classlightsource.select_db_table.currentIndexChanged.connect(self.updateLightSourceComboBox)
        # if self.classlightsource.select_db_table.changeEvent():
        #     self.updateLightSourceComboBox()
        # self.classlightsource.tablename_signal.connect(self.updateLightSourceComboBox)

        # 設定當前選中項目的文字顏色

        self.light_source_led = QComboBox()
        self.light_source_led_datatable = QComboBox()
        self.light_source_led_datatable.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 設定當前選中項目的文字顏色
        self.light_source_led.setStyleSheet(QCOMBOXSETTING)
        # 更新light_source_led_table
        self.update_light_source_led_datatable()
        # 觸發table更新,lightsource_led表單
        self.updateLightSourceledComboBox()  # 初始化
        self.light_source_led_datatable.currentIndexChanged.connect(self.updateLightSourceledComboBox)


        self.source_led = QComboBox()
        self.source_led_datatable = QComboBox()
        self.source_led_datatable.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 設定當前選中項目的文字顏色
        self.source_led.setStyleSheet(QCOMBOXSETTING)
        # 更新source_led_table
        self.update_source_led_datatable()
        # 觸發table更新,lightsource_led表單
        self.updateSourceledComboBox()  # 初始化
        self.source_led_datatable.currentIndexChanged.connect(self.updateSourceledComboBox)

        self.label_source = QLabel("Light_Source_BLU")
        self.label_source_led = QLabel("Light_source_LED")
        self.label_led_data = QLabel("led_data_base")

        # Layer區---------------------------------------------------------------
        self.layer1_mode = QComboBox()
        layer1_mode_items = ["未選", "自訂", "模擬"]
        for item in layer1_mode_items:
            self.layer1_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer1_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer1 = QLabel("Layer_1")
        self.layer1_box = QComboBox()
        layer1_box_items = ["未選"]
        for item in layer1_box_items:
            self.layer1_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer1_box.setStyleSheet(QCOMBOXSETTING)

        self.layer2_mode = QComboBox()
        layer2_mode_items = ["未選", "自訂", "模擬"]
        for item in layer2_mode_items:
            self.layer2_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer2_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer2 = QLabel("Layer_2")
        self.layer2_box = QComboBox()
        layer2_box_items = ["未選"]
        for item in layer2_box_items:
            self.layer2_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer2_box.setStyleSheet(QCOMBOXSETTING)

        self.layer3_mode = QComboBox()
        layer3_mode_items = ["未選", "自訂", "模擬"]
        for item in layer3_mode_items:
            self.layer3_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer3_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer3 = QLabel("Layer_3")
        self.layer3_box = QComboBox()
        layer3_box_items = ["未選"]
        for item in layer3_box_items:
            self.layer3_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer3_box.setStyleSheet(QCOMBOXSETTING)

        self.layer4_mode = QComboBox()
        layer4_mode_items = ["未選", "自訂", "模擬"]
        for item in layer4_mode_items:
            self.layer4_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer4_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer4 = QLabel("Layer_4")
        self.layer4_box = QComboBox()
        layer4_box_items = ["未選"]
        for item in layer4_box_items:
            self.layer4_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer4_box.setStyleSheet(QCOMBOXSETTING)

        self.layer5_mode = QComboBox()
        layer5_mode_items = ["未選", "自訂", "模擬"]
        for item in layer5_mode_items:
            self.layer5_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer5_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer5 = QLabel("Layer_5")
        self.layer5_box = QComboBox()
        layer5_box_items = ["未選"]
        for item in layer5_box_items:
            self.layer5_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer5_box.setStyleSheet(QCOMBOXSETTING)

        self.layer6_mode = QComboBox()
        layer6_mode_items = ["未選", "自訂", "模擬"]
        for item in layer6_mode_items:
            self.layer6_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer6_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer6 = QLabel("Layer_6")
        self.layer6_box = QComboBox()
        layer6_box_items = ["未選"]
        for item in layer6_box_items:
            self.layer6_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer6_box.setStyleSheet(QCOMBOXSETTING)
        # RGB-Fix區域---------------------------------------------------
        self.RGB_fix_label = QLabel("RGB-Fix")
        self.RGB_fix_label.setStyleSheet("color: #5151A2; font-weight: bold; border: 2px solid black;")
        self.R_fix_mode = QComboBox()
        R_fix_mode_items = ["未選", "自訂", "模擬"]
        for item in R_fix_mode_items:
            self.R_fix_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.R_fix_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.R_fix_label = QLabel("R-CF-Fix")
        self.R_fix_box = QComboBox()
        R_fix_box_items = ["未選"]
        for item in R_fix_box_items:
            self.R_fix_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.R_fix_box.setStyleSheet(QCOMBOXSETTING)
        self.R_TK_edit_label = QLabel("R-Fix-TK")
        self.R_TK_edit = QLineEdit()
        self.R_TK_edit.setFixedSize(100, 25)
        # G-Fix
        self.G_fix_mode = QComboBox()
        G_fix_mode_items = ["未選", "自訂", "模擬"]
        for item in G_fix_mode_items:
            self.G_fix_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.G_fix_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.G_fix_label = QLabel("G-CF-Fix")
        self.G_fix_box = QComboBox()
        G_fix_box_items = ["未選"]
        for item in G_fix_box_items:
            self.G_fix_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.G_fix_box.setStyleSheet(QCOMBOXSETTING)
        self.G_TK_edit_label = QLabel("G-Fix-TK")
        self.G_TK_edit = QLineEdit()
        self.G_TK_edit.setFixedSize(100, 25)
        # B-Fix
        self.B_fix_mode = QComboBox()
        B_fix_mode_items = ["未選", "自訂", "模擬"]
        for item in B_fix_mode_items:
            self.B_fix_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.B_fix_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.B_fix_label = QLabel("B-CF-Fix")
        self.B_fix_box = QComboBox()
        B_fix_box_items = ["未選"]
        for item in B_fix_box_items:
            self.B_fix_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.B_fix_box.setStyleSheet(QCOMBOXSETTING)
        self.B_TK_edit_label = QLabel("B-Fix-TK")
        self.B_TK_edit = QLineEdit()
        self.B_TK_edit.setFixedSize(100, 25)
        # RGB-α,K 區域---------------------------------------------------
        self.RGB_aK_label = QLabel("RGB-α,K")
        self.RGB_aK_label.setStyleSheet("color: #5151A2; font-weight: bold; border: 2px solid black;")
        self.R_aK_mode = QComboBox()
        R_aK_mode_items = ["未選", "自訂", "模擬"]
        for item in R_aK_mode_items:
            self.R_aK_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.R_aK_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.R_aK_label = QLabel("R-CF-α,K")
        self.R_aK_box = QComboBox()
        R_aK_box_items = ["未選"]
        for item in R_aK_box_items:
            self.R_aK_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.R_aK_box.setStyleSheet(QCOMBOXSETTING)
        self.R_aK_TK_edit_label = QLabel("R-α,K-TK")
        self.R_aK_TK_edit = QLineEdit()
        self.R_aK_TK_edit.setFixedSize(100, 25)
        # G-α,K
        self.G_aK_mode = QComboBox()
        G_aK_mode_items = ["未選", "自訂", "模擬"]
        for item in G_aK_mode_items:
            self.G_aK_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.G_aK_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.G_aK_label = QLabel("G-F-α,K")
        self.G_aK_box = QComboBox()
        G_aK_box_items = ["未選"]
        for item in G_aK_box_items:
            self.G_aK_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.G_aK_box.setStyleSheet(QCOMBOXSETTING)
        self.G_aK_TK_edit_label = QLabel("G-α,K-TK")
        self.G_aK_TK_edit = QLineEdit()
        self.G_aK_TK_edit.setFixedSize(100, 25)
        # B-α,K
        self.B_aK_mode = QComboBox()
        B_aK_mode_items = ["未選", "自訂", "模擬"]
        for item in B_aK_mode_items:
            self.B_aK_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.B_aK_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.B_aK_label = QLabel("B-CF-α,K")
        self.B_aK_box = QComboBox()
        B_aK_box_items = ["未選"]
        for item in B_aK_box_items:
            self.B_aK_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.B_aK_box.setStyleSheet(QCOMBOXSETTING)
        self.B_aK_TK_edit_label = QLabel("B-α,K-TK")
        self.B_aK_TK_edit = QLineEdit()
        self.B_aK_TK_edit.setFixedSize(100, 25)

        # RGB-套餐 區域---------------------------------------------------
        self.RGB_set_label = QLabel("RGB-set")
        self.RGB_set_label.setStyleSheet("color: #5151A2; font-weight: bold; border: 2px solid black;")
        self.R_set_mode = QComboBox()
        R_set_mode_items = ["未選", "自訂", "模擬"]
        for item in R_set_mode_items:
            self.R_set_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.R_set_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.R_set_label = QLabel("R-CF-set")
        self.R_set_box = QComboBox()
        R_set_box_items = ["未選"]
        for item in R_set_box_items:
            self.R_set_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.R_set_box.setStyleSheet(QCOMBOXSETTING)
        self.R_set_TK_edit_label = QLabel("R-set-TK")
        self.R_set_TK_edit = QLineEdit()
        self.R_set_TK_edit.setFixedSize(100, 25)
        # G-set
        self.G_set_mode = QComboBox()
        G_set_mode_items = ["未選", "自訂", "模擬"]
        for item in G_set_mode_items:
            self.G_set_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.G_set_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.G_set_label = QLabel("G-F-set")
        self.G_set_box = QComboBox()
        G_set_box_items = ["未選"]
        for item in G_set_box_items:
            self.G_set_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.G_set_box.setStyleSheet(QCOMBOXSETTING)
        self.G_set_TK_edit_label = QLabel("G-set-TK")
        self.G_set_TK_edit = QLineEdit()
        self.G_set_TK_edit.setFixedSize(100, 25)
        # B-set
        self.B_set_mode = QComboBox()
        B_set_mode_items = ["未選", "自訂", "模擬"]
        for item in B_set_mode_items:
            self.B_set_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.B_set_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.B_set_label = QLabel("B-CF-set")
        self.B_set_box = QComboBox()
        B_set_box_items = ["未選"]
        for item in B_set_box_items:
            self.B_set_box.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.B_set_box.setStyleSheet(QCOMBOXSETTING)
        self.B_set_TK_edit_label = QLabel("B-set-TK")
        self.B_set_TK_edit = QLineEdit()
        self.B_set_TK_edit.setFixedSize(100, 25)

        # 計算Button
        self.calculate = QPushButton("Calculate")
        # 設置樣式表
        self.calculate.setStyleSheet("QPushButton {"
                                     "    background-color: #8080C0;"  # 背景顏色
                                     "    color: black;"  # 文字顏色
                                     "    border: 2px solid #4CAF50;"  # 邊框
                                     "    border-radius: 5px;"  # 圓角
                                     "    font-weight: bold;"
                                     "border: 2px solid black;"
                                     "    padding: 5px 10px;"  # 內邊距
                                     "}"

                                     "QPushButton:hover {"
                                     "    background-color: #C7C7E2;"  # 滑鼠懸停時的背景顏色
                                     "}")

        # 設置游標樣式
        self.calculate.setCursor(Qt.PointingHandCursor)  # 手指形狀

        # refreshButton
        self.refresh = QPushButton("refresh")
        # 設置樣式表
        self.refresh.setStyleSheet("QPushButton {"
                                     "    background-color: #8080C0;"  # 背景顏色
                                     "    color: black;"  # 文字顏色
                                     "    border: 2px solid #4CAF50;"  # 邊框
                                     "    border-radius: 5px;"  # 圓角
                                     "    font-weight: bold;"
                                     "border: 2px solid black;"
                                     "    padding: 5px 10px;"  # 內邊距
                                     "}"

                                     "QPushButton:hover {"
                                     "    background-color: #C7C7E2;"  # 滑鼠懸停時的背景顏色
                                     "}")

        # 設置游標樣式
        self.refresh.setCursor(Qt.PointingHandCursor)  # 手指形狀

        self.refresh.clicked.connect(self.get_name)
        self.calculate.clicked.connect(self.calculate_color)

        # widget放置
        # color table
        self.Color_Enter_layout.addWidget(self.color_table, 0, 0,4,10)
        # Light source區域---------------------------------------
        self.Color_Enter_layout.addWidget(self.label_source, 4, 1, 1, 1)
        self.Color_Enter_layout.addWidget(self.light_source_datatable, 4, 2, 1, 1)
        self.Color_Enter_layout.addWidget(self.label_source_led, 4, 3, 1, 1)
        self.Color_Enter_layout.addWidget(self.light_source_led_datatable, 4, 4, 1, 1)
        self.Color_Enter_layout.addWidget(self.label_led_data, 4, 5, 1, 1)
        self.Color_Enter_layout.addWidget(self.source_led_datatable, 4, 6, 1, 1)
        self.Color_Enter_layout.addWidget(self.light_source_mode, 5, 0)
        self.Color_Enter_layout.addWidget(self.light_source, 5, 1, 1, 2)
        self.Color_Enter_layout.addWidget(self.light_source_led, 5, 3, 1, 2)
        self.Color_Enter_layout.addWidget(self.source_led, 5, 5, 1, 2)
        self.Color_Enter_layout.addWidget(self.calculate, 4, 7, 1, 1)
        self.Color_Enter_layout.addWidget(self.refresh, 5, 7, 1, 1)
        # Layer區域---------------------------------------------------
        self.Color_Enter_layout.addWidget(self.layer1_mode, 7, 0)
        self.Color_Enter_layout.addWidget(self.layer1_box, 7, 1, 1, 2)
        self.Color_Enter_layout.addWidget(self.label_layer1, 6, 0, 1, 3)  # 佔兩欄

        self.Color_Enter_layout.addWidget(self.layer2_mode, 7, 3)
        self.Color_Enter_layout.addWidget(self.layer2_box, 7, 4, 1, 2)
        self.Color_Enter_layout.addWidget(self.label_layer2, 6, 3, 1, 3)

        self.Color_Enter_layout.addWidget(self.layer3_mode, 7, 6)
        self.Color_Enter_layout.addWidget(self.layer3_box, 7, 7, 1, 2)
        self.Color_Enter_layout.addWidget(self.label_layer3, 6, 6, 1, 3)

        self.Color_Enter_layout.addWidget(self.layer4_mode, 9, 0)
        self.Color_Enter_layout.addWidget(self.layer4_box, 9, 1, 1, 2)
        self.Color_Enter_layout.addWidget(self.label_layer4, 8, 0, 1, 3)

        self.Color_Enter_layout.addWidget(self.layer5_mode, 9, 3)
        self.Color_Enter_layout.addWidget(self.layer5_box, 9, 4, 1, 2)
        self.Color_Enter_layout.addWidget(self.label_layer5, 8, 3, 1, 3)

        self.Color_Enter_layout.addWidget(self.layer6_mode, 9, 6)
        self.Color_Enter_layout.addWidget(self.layer6_box, 9, 7, 1, 2)
        self.Color_Enter_layout.addWidget(self.label_layer6, 8, 6, 1, 3)

        # RGB-Fix區域---------------------------------------------------
        self.Color_Enter_layout.addWidget(self.RGB_fix_label, 10, 0)
        self.Color_Enter_layout.addWidget(self.R_fix_mode, 11, 0)
        self.Color_Enter_layout.addWidget(self.R_fix_box, 11, 1)
        self.Color_Enter_layout.addWidget(self.R_fix_label, 10, 1)
        self.Color_Enter_layout.addWidget(self.R_TK_edit_label, 10, 2)
        self.Color_Enter_layout.addWidget(self.R_TK_edit, 11, 2)

        self.Color_Enter_layout.addWidget(self.G_fix_mode, 11, 3)
        self.Color_Enter_layout.addWidget(self.G_fix_box, 11, 4)
        self.Color_Enter_layout.addWidget(self.G_fix_label, 10, 4)
        self.Color_Enter_layout.addWidget(self.G_TK_edit_label, 10, 5)
        self.Color_Enter_layout.addWidget(self.G_TK_edit, 11, 5)

        self.Color_Enter_layout.addWidget(self.B_fix_mode, 11, 6)
        self.Color_Enter_layout.addWidget(self.B_fix_box, 11, 7)
        self.Color_Enter_layout.addWidget(self.B_fix_label, 10, 7)
        self.Color_Enter_layout.addWidget(self.B_TK_edit_label, 10, 8)
        self.Color_Enter_layout.addWidget(self.B_TK_edit, 11, 8)

        # RGB-α,K區域---------------------------------------------------
        self.Color_Enter_layout.addWidget(self.RGB_aK_label, 12, 0)
        self.Color_Enter_layout.addWidget(self.R_aK_mode, 13, 0)
        self.Color_Enter_layout.addWidget(self.R_aK_box, 13, 1)
        self.Color_Enter_layout.addWidget(self.R_aK_label, 12, 1)
        self.Color_Enter_layout.addWidget(self.R_aK_TK_edit_label, 12, 2)
        self.Color_Enter_layout.addWidget(self.R_aK_TK_edit, 13, 2)

        self.Color_Enter_layout.addWidget(self.G_aK_mode, 13, 3)
        self.Color_Enter_layout.addWidget(self.G_aK_box, 13, 4)
        self.Color_Enter_layout.addWidget(self.G_aK_label, 12, 4)
        self.Color_Enter_layout.addWidget(self.G_aK_TK_edit_label, 12, 5)
        self.Color_Enter_layout.addWidget(self.G_aK_TK_edit, 13, 5)

        self.Color_Enter_layout.addWidget(self.B_aK_mode, 13, 6)
        self.Color_Enter_layout.addWidget(self.B_aK_box, 13, 7)
        self.Color_Enter_layout.addWidget(self.B_aK_label, 12, 7)
        self.Color_Enter_layout.addWidget(self.B_aK_TK_edit_label, 12, 8)
        self.Color_Enter_layout.addWidget(self.B_aK_TK_edit, 13, 8)

        # RGB-set區域---------------------------------------------------
        self.Color_Enter_layout.addWidget(self.RGB_set_label, 14, 0)
        self.Color_Enter_layout.addWidget(self.R_set_mode, 15, 0)
        self.Color_Enter_layout.addWidget(self.R_set_box, 15, 1)
        self.Color_Enter_layout.addWidget(self.R_set_label, 14, 1)
        self.Color_Enter_layout.addWidget(self.R_set_TK_edit_label, 14, 2)
        self.Color_Enter_layout.addWidget(self.R_set_TK_edit, 15, 2)

        self.Color_Enter_layout.addWidget(self.G_set_mode, 15, 3)
        self.Color_Enter_layout.addWidget(self.G_set_box, 15, 4)
        self.Color_Enter_layout.addWidget(self.G_set_label, 14, 4)
        self.Color_Enter_layout.addWidget(self.G_set_TK_edit_label, 14, 5)
        self.Color_Enter_layout.addWidget(self.G_set_TK_edit, 15, 5)

        self.Color_Enter_layout.addWidget(self.B_set_mode, 15, 6)
        self.Color_Enter_layout.addWidget(self.B_set_box, 15, 7)
        self.Color_Enter_layout.addWidget(self.B_set_label, 14, 7)
        self.Color_Enter_layout.addWidget(self.B_set_TK_edit_label, 14, 8)
        self.Color_Enter_layout.addWidget(self.B_set_TK_edit, 15, 8)


        # layout放置
        self.setLayout(self.Color_Enter_layout)

        # Set Background
        self.setStyleSheet("background-color: lightyellow;")

    # def updateLightSourceComboBox(self,table_name):
    #     print("OK")
    #
    #     # 當 Light_Source_BLU 类的表格选择变更时调用,select_db_table是Light_Source_BLU的combobox選項
    #     selected_table = table_name
    #     print("selected_table",table_name)
    #     # self.classlightsource.loadTableData(table_name =selected_table)
    #     if selected_table:
    #         # 獲取 BLUspectrum 的表头选项
    #         header_items = self.classlightsource.getTableHeader(table_name=selected_table)
    #         print("select_table",selected_table)
    #         print("header_items",header_items)
    #
    #         # 清空 light_source 的选项
    #         self.light_source.clear()
    #
    #
    #         # 将获取的表头选项添加到 light_source 中
    #         for item in header_items:
    #             self.light_source.addItem(str(item))
    #             print("item", str(item))

    def get_name(self):
        print("self.classlightsource.tablename_signal.connect(self.updateLightSourceComboBox)",self.classlightsource.tablename_signal.connect(self.updateLightSourceComboBox))

        self.classlightsource.tablename_signal.connect(self.updateLightSourceComboBox)

        print("PP")

    def test(self):
        self.color_table.setItem(2, 5, QTableWidgetItem(f"{self.light_source_mode.currentText()}"))
        if self.light_source_datatable.currentText() == 'blu_data':
            print("SSS")

    def calculate_color(self):
        print("calculate start")
        # BLU
        if self.light_source_mode.currentText() == "自訂":
            print("OKin")
            connection = sqlite3.connect("blu_database.db")
            cursor = connection.cursor()

            # 確保表格存在
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name=?;",
                           (self.light_source_datatable.currentText(),))
            if cursor.fetchone() is not None:
                # 確保欄位存在
                cursor.execute(f"PRAGMA table_info('{self.light_source_datatable.currentText()}');")
                columns = [column[1] for column in cursor.fetchall()]
                if self.light_source.currentText() in columns:

                    # 取得資料
                    column_name = self.light_source.currentText()
                    table_name = self.light_source_datatable.currentText()
                    # 使用正確的引號包裹表名和列名
                    query = f"SELECT '{column_name}' FROM '{table_name}';"
                    cursor.execute(query)
                    result = cursor.fetchall()
                    print("result",result)

                    # 取得指定欄位的數據
                    column_data = [row[0] for row in result]
                    BLU_spectrum_Series = pd.Series(column_data)
                    print("BLU_spectrum_Series", BLU_spectrum_Series)

                else:
                    print(
                        f"欄位 {self.light_source.currentText()} 不存在於表格 {self.light_source_datatable.currentText()} 中")
            else:
                print(f"表格 {self.light_source_datatable.currentText()} 不存在")

            # 關閉連線
            connection.close()
        else:
            print("NG")
# table更新function區--------------------------------------------------------------------------------------
    def update_light_source_datatable(self):
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("blu_database.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.light_source_datatable.clear()
        for table in tables:
            self.light_source_datatable.addItem(table[0])

        # 關閉連線
        conn.close()
    def updateLightSourceComboBox(self):
        connection = sqlite3.connect("blu_database.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.light_source_datatable.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        print("headerlabels-from source", header_labels)
        self.light_source.clear()
        for item in header_labels:
            self.light_source.addItem(str(item))
            print("item", str(item))
        # 關閉連線
        connection.close()
    def update_light_source_led_datatable(self):
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("led_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.light_source_led_datatable.clear()
        for table in tables:
            self.light_source_led_datatable.addItem(table[0])

        # 關閉連線
        conn.close()
    def updateLightSourceledComboBox(self):
        connection = sqlite3.connect("led_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        # cursor.execute(f"PRAGMA table_info({self.light_source_led_datatable.currentText()});")
        cursor.execute(f"PRAGMA table_info('{self.light_source_led_datatable.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        print("headerlabels-from source", header_labels)
        self.light_source_led.clear()
        for item in header_labels:
            self.light_source_led.addItem(str(item))
            print("item", str(item))
        # 關閉連線
        connection.close()
    def update_source_led_datatable(self):
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("led_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.source_led_datatable.clear()
        for table in tables:
            self.source_led_datatable.addItem(table[0])

        # 關閉連線
        conn.close()
    def updateSourceledComboBox(self):
        connection = sqlite3.connect("led_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.source_led_datatable.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        print("headerlabels-from source", header_labels)
        self.source_led.clear()
        for item in header_labels:
            self.source_led.addItem(str(item))
            print("item", str(item))
        # 關閉連線
        connection.close()

class Color_Select(QWidget):
    def __init__(self):
        super().__init__()
        self.Color_Select_layout = QGridLayout()
        Color_Select_label = QLabel("Color_Select")

        # 篩選Create
        # W
        self.W_checkbox = QCheckBox()
        self.W_checkbox.setChecked(False)  # 預設非勾選狀態
        self.W_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.Wx_label = QLabel("Wx")
        self.Wx_edit = QLineEdit()
        self.Wy_label = QLabel("Wy")
        self.Wy_edit = QLineEdit()
        self.W_tolerance = QLabel("Tolerance")
        self.W_tolerance_edit = QLineEdit()
        # R
        self.R_checkbox = QCheckBox()
        self.R_checkbox.setChecked(False)  # 預設非勾選狀態
        self.R_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.Rx_label = QLabel("Rx")
        self.Rx_edit = QLineEdit()
        self.Ry_label = QLabel("Ry")
        self.Ry_edit = QLineEdit()
        self.R_tolerance = QLabel("Tolerance")
        self.R_tolerance_edit = QLineEdit()
        # G
        self.G_checkbox = QCheckBox()
        self.G_checkbox.setChecked(False)  # 預設非勾選狀態
        self.G_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.Gx_label = QLabel("Gx")
        self.Gx_edit = QLineEdit()
        self.Gy_label = QLabel("Gy")
        self.Gy_edit = QLineEdit()
        self.G_tolerance = QLabel("Tolerance")
        self.G_tolerance_edit = QLineEdit()
        # B
        self.B_checkbox = QCheckBox()
        self.B_checkbox.setChecked(False)  # 預設非勾選狀態
        self.B_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.Bx_label = QLabel("Bx")
        self.Bx_edit = QLineEdit()
        self.By_label = QLabel("By")
        self.By_edit = QLineEdit()
        self.B_tolerance = QLabel("Tolerance")
        self.B_tolerance_edit = QLineEdit()

        # widget放置區
        self.Color_Select_layout.addWidget(Color_Select_label, 0, 0, 0, 14)
        self.Color_Select_layout.addWidget(self.W_checkbox, 1, 0)
        self.Color_Select_layout.addWidget(self.Wx_label, 1, 1)
        self.Color_Select_layout.addWidget(self.Wx_edit, 1, 2)
        self.Color_Select_layout.addWidget(self.Wy_label, 1, 3)
        self.Color_Select_layout.addWidget(self.Wy_edit, 1, 4)
        self.Color_Select_layout.addWidget(self.W_tolerance, 1, 5)
        self.Color_Select_layout.addWidget(self.W_tolerance_edit, 1, 6)

        self.Color_Select_layout.addWidget(self.R_checkbox, 1, 7)
        self.Color_Select_layout.addWidget(self.Rx_label, 1, 8)
        self.Color_Select_layout.addWidget(self.Rx_edit, 1, 9)
        self.Color_Select_layout.addWidget(self.Ry_label, 1, 10)
        self.Color_Select_layout.addWidget(self.Ry_edit, 1, 11)
        self.Color_Select_layout.addWidget(self.R_tolerance, 1, 12)
        self.Color_Select_layout.addWidget(self.R_tolerance_edit, 1, 13)

        self.Color_Select_layout.addWidget(self.G_checkbox, 2, 0)
        self.Color_Select_layout.addWidget(self.Gx_label, 2, 1)
        self.Color_Select_layout.addWidget(self.Gx_edit, 2, 2)
        self.Color_Select_layout.addWidget(self.Gy_label, 2, 3)
        self.Color_Select_layout.addWidget(self.Gy_edit, 2, 4)
        self.Color_Select_layout.addWidget(self.G_tolerance, 2, 5)
        self.Color_Select_layout.addWidget(self.G_tolerance_edit, 2, 6)

        self.Color_Select_layout.addWidget(self.B_checkbox, 2, 7)
        self.Color_Select_layout.addWidget(self.Bx_label, 2, 8)
        self.Color_Select_layout.addWidget(self.Bx_edit, 2, 9)
        self.Color_Select_layout.addWidget(self.By_label, 2, 10)
        self.Color_Select_layout.addWidget(self.By_edit, 2, 11)
        self.Color_Select_layout.addWidget(self.B_tolerance, 2, 12)
        self.Color_Select_layout.addWidget(self.B_tolerance_edit, 2, 13)

        self.Color_Select_layout.setHorizontalSpacing(5)  # 設定水平間距為5像素

        self.Color_Select_layout.setVerticalSpacing(1)  # 設定水平間距為5像素

        # layout放置
        self.setLayout(self.Color_Select_layout)

        # Set Background
        self.setStyleSheet("background-color: #FFBD9D;")