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
import math


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

        #self.light_source.currentIndexChanged.connect(self.calculate_color)
        self.light_source_datatable = QComboBox()
        #self.light_source_datatable.currentIndexChanged.connect(self.calculate_color)
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
        layer1_mode_items = ["未選", "自訂"]
        for item in layer1_mode_items:
            self.layer1_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer1_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer1 = QLabel("Layer_1(Cell總和)")
        # cell_頻譜選擇
        self.layer1_box = QComboBox()
        # layer1_table
        self.layer1_table = QComboBox()
        self.layer1_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_layer1_datatable()
        # 觸發layer1 table
        self.updatelayer1ComboBox()# 初始化
        self.layer1_table.currentIndexChanged.connect(self.updatelayer1ComboBox)
        # 設定當前選中項目的文字顏色
        self.layer1_box.setStyleSheet(QCOMBOXSETTING)

        self.layer2_mode = QComboBox()
        layer2_mode_items = ["未選", "自訂"]
        for item in layer2_mode_items:
            self.layer2_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer2_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer2 = QLabel("Layer_2(BSITO)")
        # cell_頻譜選擇
        self.layer2_box = QComboBox()
        # layer2_table
        self.layer2_table = QComboBox()
        self.layer2_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_layer2_datatable()
        # 觸發layer2 table
        self.updatelayer2ComboBox()  # 初始化
        self.layer2_table.currentIndexChanged.connect(self.updatelayer2ComboBox)
        # 設定當前選中項目的文字顏色
        self.layer2_box.setStyleSheet(QCOMBOXSETTING)

        self.layer3_mode = QComboBox()
        layer3_mode_items = ["未選", "自訂"]
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
        layer4_mode_items = ["未選", "自訂"]
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
        layer5_mode_items = ["未選", "自訂"]
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
        layer6_mode_items = ["未選", "自訂"]
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
        # RGB-Fix區域--------------------------------------------------------------------------------------------------
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
        # 設定當前選中項目的文字顏色
        self.R_fix_box.setStyleSheet(QCOMBOXSETTING)
        self.R_fix_table = QComboBox()
        self.R_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 觸發table連動改變
        self.update_RCF_Fix_datatable()
        self.updateRCF_Fix_ComboBox()
        self.R_fix_table.currentIndexChanged.connect(self.update_RCF_Fix_datatable)

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
        # 設定當前選中項目的文字顏色
        self.G_fix_box.setStyleSheet(QCOMBOXSETTING)

        self.G_fix_table = QComboBox()
        self.G_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 觸發table連動改變
        self.update_GCF_Fix_datatable()
        self.updateGCF_Fix_ComboBox()
        self.G_fix_table.currentIndexChanged.connect(self.update_GCF_Fix_datatable)

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
        # 設定當前選中項目的文字顏色
        self.B_fix_box.setStyleSheet(QCOMBOXSETTING)
        self.B_fix_table = QComboBox()
        self.B_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 觸發table連動改變
        self.update_BCF_Fix_datatable()
        self.updateBCF_Fix_ComboBox()
        self.B_fix_table.currentIndexChanged.connect(self.update_BCF_Fix_datatable)

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

        self.refresh.clicked.connect(self.test)
        self.calculate.clicked.connect(self.calculate_color_customize)

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
        self.Color_Enter_layout.addWidget(self.label_layer1, 6, 0, 1, 2)  # 佔兩欄
        self.Color_Enter_layout.addWidget(self.layer1_table, 6, 2, 1, 1)  # 佔兩欄

        self.Color_Enter_layout.addWidget(self.layer2_mode, 7, 3)
        self.Color_Enter_layout.addWidget(self.layer2_box, 7, 4, 1, 2)
        self.Color_Enter_layout.addWidget(self.label_layer2, 6, 3, 1, 2)
        self.Color_Enter_layout.addWidget(self.layer2_table, 6, 5, 1, 1)

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
        self.Color_Enter_layout.addWidget(self.R_fix_table, 10, 1)
        self.Color_Enter_layout.addWidget(self.R_TK_edit_label, 10, 2)
        self.Color_Enter_layout.addWidget(self.R_TK_edit, 11, 2)

        self.Color_Enter_layout.addWidget(self.G_fix_mode, 11, 3)
        self.Color_Enter_layout.addWidget(self.G_fix_box, 11, 4)
        self.Color_Enter_layout.addWidget(self.G_fix_table, 10, 4)
        self.Color_Enter_layout.addWidget(self.G_TK_edit_label, 10, 5)
        self.Color_Enter_layout.addWidget(self.G_TK_edit, 11, 5)

        self.Color_Enter_layout.addWidget(self.B_fix_mode, 11, 6)
        self.Color_Enter_layout.addWidget(self.B_fix_box, 11, 7)
        self.Color_Enter_layout.addWidget(self.B_fix_table, 10, 7)
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


    def test(self):
        self.color_table.setItem(2, 5, QTableWidgetItem(f"{self.light_source_mode.currentText()}"))
        if self.light_source_datatable.currentText() == 'blu_data':
            print("SSS")

    def calculate_BLU(self):
        # BLU
        # 選項關鍵----------------------------------------------
        if self.light_source_mode.currentText() == "自訂":
            print("OKin_BLU_自訂")
            connection_BLU = sqlite3.connect("blu_database.db")
            cursor_BLU = connection_BLU.cursor()
            # 取得BLU資料
            column_name_BLU = self.light_source.currentText()
            table_name_BLU = self.light_source_datatable.currentText()
            # 使用正確的引號包裹表名和列名
            query_BLU = f"SELECT * FROM '{table_name_BLU}';"
            cursor_BLU.execute(query_BLU)
            result_BLU = cursor_BLU.fetchall()
            # print("result_BLU",result_BLU)

            # 找到指定標題的欄位索引
            header_BLU = [column[0] for column in cursor_BLU.description]
            print("header_BLU",header_BLU)
            column_index_BLU = header_BLU.index(f"{column_name_BLU}")

            # 取得下一個欄位的名稱
            next_column_name_BLU = header_BLU[column_index_BLU + 1]
            print("next_column_name_BLU",next_column_name_BLU)

            # 取得指定欄位的數據並轉換為 Series
            BLU_spectrum_Series = pd.Series([row[column_index_BLU] for row in result_BLU])
            print("column_index_BLU",column_index_BLU)
            BLU_spectrum_Series_test =pd.Series([row[column_index_BLU+1] for row in result_BLU])
            print("BLU_spectrum_Series_test", BLU_spectrum_Series_test)

            # 將 Series 中的字符串轉換為數值
            BLU_spectrum_Series = pd.to_numeric(BLU_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            BLU_spectrum_Series = BLU_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            BLU_spectrum_Series = BLU_spectrum_Series.dropna()

            # 關閉連線
            connection_BLU.close()
            # 自訂BLU_Spectrum回傳
            return BLU_spectrum_Series
            # 選項關鍵----------------------------------------------------------
        elif self.light_source_mode.currentText() == "替換":
            print("OKinBLU替換")
            connection_BLU = sqlite3.connect("blu_database.db")
            cursor_BLU = connection_BLU.cursor()
            # 取得BLU資料------------------------------------------------------
            column_name_BLU = self.light_source.currentText()
            table_name_BLU = self.light_source_datatable.currentText()
            # 使用正確的引號包裹表名和列名
            query_BLU = f"SELECT * FROM '{table_name_BLU}';"
            cursor_BLU.execute(query_BLU)
            result_BLU = cursor_BLU.fetchall()
            # print("result_BLU",result_BLU)

            # 找到指定標題的欄位索引
            header_BLU = [column[0] for column in cursor_BLU.description]
            column_index_BLU = header_BLU.index(f"{column_name_BLU}")

            # 取得指定欄位的數據並轉換為 Series
            BLU_spectrum_Series = pd.Series([row[column_index_BLU] for row in result_BLU])
            # 將 Series 中的字符串轉換為數值
            BLU_spectrum_Series = pd.to_numeric(BLU_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            BLU_spectrum_Series = BLU_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            BLU_spectrum_Series = BLU_spectrum_Series.dropna()

            # 轉換為 float64 數據類型
            BLU_spectrum_Series = BLU_spectrum_Series.astype(float)

            # 取得led_source資料
            connection_led = sqlite3.connect("led_spectrum.db")
            cursor_led = connection_led.cursor()
            # 取得LED_Source資料------------------------------------------------------
            column_name_led_source = self.light_source_led.currentText()
            table_name_led_source = self.light_source_led_datatable.currentText()
            # 使用正確的引號包裹表名和列名
            query_LED_light_Source = f"SELECT * FROM '{table_name_led_source}';"
            cursor_led.execute(query_LED_light_Source)
            result_led_light_source = cursor_led.fetchall()

            # 找到指定標題的欄位索引
            header_LED_light_Source = [column[0] for column in cursor_led.description]
            column_index_LED_light_Source = header_LED_light_Source.index(f"{column_name_led_source}")

            # 取得指定欄位的數據並轉換為 LED_Source_Series
            LED_light_Source_spectrum_Series = pd.Series([row[column_index_LED_light_Source] for row in result_led_light_source])
            LED_light_Source_spectrum_Series = pd.to_numeric(LED_light_Source_spectrum_Series, errors='coerce')
            LED_light_Source_spectrum_Series = LED_light_Source_spectrum_Series.dropna()
            LED_light_Source_spectrum_Series = LED_light_Source_spectrum_Series.astype(float)
            # print("LED_Source_spectrum_Series",LED_light_Source_spectrum_Series)

            # 取得LED_database資料------------------------------------------------------
            column_name_led_source_data = self.source_led.currentText()
            table_name_led_source_data = self.source_led_datatable.currentText()
            # 使用正確的引號包裹表名和列名
            query_LED_Source = f"SELECT * FROM '{table_name_led_source_data}';"
            cursor_led.execute(query_LED_Source)
            result_led_source = cursor_led.fetchall()

            # 找到指定標題的欄位索引
            header_LED_Source = [column[0] for column in cursor_led.description]
            column_index_LED_Source = header_LED_Source.index(f"{column_name_led_source_data}")

            # 取得指定欄位的數據並轉換為 LED_Source_Series
            LED_Source_spectrum_Series = pd.Series([row[column_index_LED_Source] for row in result_led_source])
            LED_Source_spectrum_Series = pd.to_numeric(LED_Source_spectrum_Series, errors='coerce')
            LED_Source_spectrum_Series = LED_Source_spectrum_Series.dropna()
            LED_Source_spectrum_Series = LED_Source_spectrum_Series.astype(float)

            # 替換BLU LED光譜
            BLU_spectrum_Series = BLU_spectrum_Series / LED_light_Source_spectrum_Series * LED_Source_spectrum_Series

            # 關閉連線
            connection_BLU.close()
            connection_led.close()
            return BLU_spectrum_Series
        else:
            print("BLU_None")
            return None
    # Cell
    def calculate_layer1(self):
        if self.layer1_mode.currentText() == "自訂":
            print("in_layer1_自訂")
            connection_layer1 = sqlite3.connect("cell_spectrum.db")
            cursor_layer1 = connection_layer1.cursor()
            # 取得BLU資料
            column_name_layer1 = self.layer1_box.currentText()
            table_name_layer1 = self.layer1_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer1 = f"SELECT * FROM '{table_name_layer1}';"
            cursor_layer1.execute(query_layer1)
            result_layer1 = cursor_layer1.fetchall()

            # 找到指定標題的欄位索引
            header_layer1 = [column[0] for column in cursor_layer1.description]
            column_index_layer1 = header_layer1.index(f"{column_name_layer1}")

            # 取得指定欄位的數據並轉換為 Series
            layer1_spectrum_Series = pd.Series([row[column_index_layer1] for row in result_layer1])
            #print("layer1_spectrum_Series", layer1_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer1_spectrum_Series = pd.to_numeric(layer1_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer1_spectrum_Series = layer1_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer1_spectrum_Series = layer1_spectrum_Series.dropna()

            # 關閉連線
            connection_layer1.close()
            # 自訂BLU_Spectrum回傳
            return layer1_spectrum_Series
        else:
            layer1_spectrum_Series = 1
            print("layer1:未選")
            return layer1_spectrum_Series

    # BSITO
    def calculate_layer2(self):
        if self.layer2_mode.currentText() == "自訂":
            print("in_layer2_自訂")
            connection_layer2 = sqlite3.connect("BSITO_spectrum.db")
            cursor_layer2 = connection_layer2.cursor()
            # 取得BLU資料
            column_name_layer2 = self.layer2_box.currentText()
            table_name_layer2 = self.layer2_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer2 = f"SELECT * FROM '{table_name_layer2}';"
            cursor_layer2.execute(query_layer2)
            result_layer2 = cursor_layer2.fetchall()

            # 找到指定標題的欄位索引
            header_layer2 = [column[0] for column in cursor_layer2.description]
            column_index_layer2 = header_layer2.index(f"{column_name_layer2}")

            # 取得指定欄位的數據並轉換為 Series
            layer2_spectrum_Series = pd.Series([row[column_index_layer2] for row in result_layer2])
            #print("layer2_spectrum_Series", layer2_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer2_spectrum_Series = pd.to_numeric(layer2_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer2_spectrum_Series = layer2_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer2_spectrum_Series = layer2_spectrum_Series.dropna()

            # 關閉連線
            connection_layer2.close()
            # 自訂BLU_Spectrum回傳
            return layer2_spectrum_Series
        else:
            layer2_spectrum_Series = 1
            print("layer2:未選")
            return layer2_spectrum_Series

    # RCF_Fix
    def calculate_RCF_Fix(self):
        if self.R_fix_mode.currentText() == "自訂":
            print("in_RCF_Fix_自訂")
            connection_RCF_Fix = sqlite3.connect("RCF_Fix_spectrum.db")
            cursor_RCF_Fix = connection_RCF_Fix.cursor()
            # 取得BLU資料
            column_name_RCF_Fix = self.R_fix_box.currentText()
            table_name_RCF_Fix = self.R_fix_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_RCF_Fix = f"SELECT * FROM '{table_name_RCF_Fix}';"
            cursor_RCF_Fix.execute(query_RCF_Fix)
            result_RCF_Fix = cursor_RCF_Fix.fetchall()

            # 找到指定標題的欄位索引
            header_RCF_Fix = [column[0] for column in cursor_RCF_Fix.description]
            column_index_RCF_Fix = header_RCF_Fix.index(f"{column_name_RCF_Fix}")

            # 取得指定欄位的數據並轉換為 Series
            RCF_Fix_spectrum_Series = pd.Series([row[column_index_RCF_Fix] for row in result_RCF_Fix])
            #print("RCF_Fix_spectrum_Series", RCF_Fix_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            RCF_Fix_spectrum_Series = pd.to_numeric(RCF_Fix_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            RCF_Fix_spectrum_Series = RCF_Fix_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            RCF_Fix_spectrum_Series = RCF_Fix_spectrum_Series.dropna()

            # 關閉連線
            connection_RCF_Fix.close()
            # 自訂RCF_Fix_Spectrum回傳
            return RCF_Fix_spectrum_Series
        else:
            print("RCF_Fix:None")
            return None

    # GCF_Fix
    def calculate_GCF_Fix(self):
        if self.G_fix_mode.currentText() == "自訂":
            print("in_GCF_Fix_自訂")
            connection_GCF_Fix = sqlite3.connect("GCF_Fix_spectrum.db")
            cursor_GCF_Fix = connection_GCF_Fix.cursor()
            # 取得BLU資料
            column_name_GCF_Fix = self.G_fix_box.currentText()
            table_name_GCF_Fix = self.G_fix_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_GCF_Fix = f"SELECT * FROM '{table_name_GCF_Fix}';"
            cursor_GCF_Fix.execute(query_GCF_Fix)
            result_GCF_Fix = cursor_GCF_Fix.fetchall()

            # 找到指定標題的欄位索引
            header_GCF_Fix = [column[0] for column in cursor_GCF_Fix.description]
            column_index_GCF_Fix = header_GCF_Fix.index(f"{column_name_GCF_Fix}")

            # 取得指定欄位的數據並轉換為 Series
            GCF_Fix_spectrum_Series = pd.Series([row[column_index_GCF_Fix] for row in result_GCF_Fix])
            #print("GCF_Fix_spectrum_Series", GCF_Fix_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            GCF_Fix_spectrum_Series = pd.to_numeric(GCF_Fix_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            GCF_Fix_spectrum_Series = GCF_Fix_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            GCF_Fix_spectrum_Series = GCF_Fix_spectrum_Series.dropna()

            # 關閉連線
            connection_GCF_Fix.close()
            # 自訂GCF_Fix_Spectrum回傳
            return GCF_Fix_spectrum_Series
        else:
            print("GCF_Fix:None")
            return None

    # BCF_Fix
    def calculate_BCF_Fix(self):
        if self.B_fix_mode.currentText() == "自訂":
            print("in_BCF_Fix_自訂")
            connection_BCF_Fix = sqlite3.connect("BCF_Fix_spectrum.db")
            cursor_BCF_Fix = connection_BCF_Fix.cursor()
            # 取得BLU資料
            column_name_BCF_Fix = self.B_fix_box.currentText()
            table_name_BCF_Fix = self.B_fix_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_BCF_Fix = f"SELECT * FROM '{table_name_BCF_Fix}';"
            cursor_BCF_Fix.execute(query_BCF_Fix)
            result_BCF_Fix = cursor_BCF_Fix.fetchall()

            # 找到指定標題的欄位索引
            header_BCF_Fix = [column[0] for column in cursor_BCF_Fix.description]
            column_index_BCF_Fix = header_BCF_Fix.index(f"{column_name_BCF_Fix}")

            # 取得指定欄位的數據並轉換為 Series
            BCF_Fix_spectrum_Series = pd.Series([row[column_index_BCF_Fix] for row in result_BCF_Fix])
            #print("BCF_Fix_spectrum_Series", BCF_Fix_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            BCF_Fix_spectrum_Series = pd.to_numeric(BCF_Fix_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            BCF_Fix_spectrum_Series = BCF_Fix_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            BCF_Fix_spectrum_Series = BCF_Fix_spectrum_Series.dropna()

            # 關閉連線
            connection_BCF_Fix.close()
            # 自訂BCF_Fix_Spectrum回傳
            return BCF_Fix_spectrum_Series
        else:
            print("BCF_Fix:None")
            return None

    def calculate_color_customize(self):
        # 取得CIE資料(共用區)---------------------------------------------
        connection_CIE = sqlite3.connect("CIE_spectrum.db")
        cursor_CIE = connection_CIE.cursor()
        # 使用正確的引號包裹表名和列名
        query_CIE = f"SELECT * FROM 'CIE1931';"
        cursor_CIE.execute(query_CIE)
        result_CIE = cursor_CIE.fetchall()

        # 找到指定標題的欄位索引
        header_CIE = [column[0] for column in cursor_CIE.description]
        column_index_CIE_X = header_CIE.index(f"x")
        column_index_CIE_Y = header_CIE.index(f"y")
        column_index_CIE_Z = header_CIE.index(f"z")

        # 取得指定欄位的數據並轉換為 Series
        CIE_spectrum_Series_X = pd.Series([row[column_index_CIE_X] for row in result_CIE])
        # print("CIE_spectrum_Series_X",CIE_spectrum_Series_X)
        CIE_spectrum_Series_Y = pd.Series([row[column_index_CIE_Y] for row in result_CIE])
        CIE_spectrum_Series_Z = pd.Series([row[column_index_CIE_Z] for row in result_CIE])

        # 將 Series 中的字符串轉換為數值
        CIE_spectrum_Series_X = pd.to_numeric(CIE_spectrum_Series_X, errors='coerce')
        CIE_spectrum_Series_Y = pd.to_numeric(CIE_spectrum_Series_Y, errors='coerce')
        CIE_spectrum_Series_Z = pd.to_numeric(CIE_spectrum_Series_Z, errors='coerce')

        # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
        CIE_spectrum_Series_X = CIE_spectrum_Series_X.fillna(0)
        CIE_spectrum_Series_Y = CIE_spectrum_Series_Y.fillna(0)
        CIE_spectrum_Series_Z = CIE_spectrum_Series_Z.fillna(0)

        # 刪除包含空值的行
        CIE_spectrum_Series_X = CIE_spectrum_Series_X.dropna()
        CIE_spectrum_Series_Y = CIE_spectrum_Series_Y.dropna()
        CIE_spectrum_Series_Z = CIE_spectrum_Series_Z.dropna()

        # 轉換為 float64 數據類型
        CIE_spectrum_Series_X = CIE_spectrum_Series_X.astype(float)
        CIE_spectrum_Series_Y = CIE_spectrum_Series_Y.astype(float)
        CIE_spectrum_Series_Z = CIE_spectrum_Series_Z.astype(float)

        # 檢查數據類型
        # print("CIE_spectrum_Series_R dtype:", CIE_spectrum_Series_R.dtype)
        # print("CIE_spectrum_Series_G dtype:", CIE_spectrum_Series_G.dtype)
        # print("CIE_spectrum_Series_B dtype:", CIE_spectrum_Series_B.dtype)

        # 取得C-light
        connection_C = sqlite3.connect("blu_database.db")
        cursor_C = connection_C.cursor()
        # 使用正確的引號包裹表名和列名
        query_C = f"SELECT * FROM 'Clight';"
        cursor_C.execute(query_C)
        result_C = cursor_C.fetchall()

        # 找到指定標題的欄位索引
        header_C = [column[0] for column in cursor_C.description]
        column_index_C = header_C.index(f"Clight")

        # 取得指定欄位的數據並轉換為 Series
        C_spectrum_Series = pd.Series([row[column_index_C] for row in result_C])
        # 將 Series 中的字符串轉換為數值
        C_spectrum_Series = pd.to_numeric(C_spectrum_Series, errors='coerce')
        # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
        C_spectrum_Series = C_spectrum_Series.fillna(0)
        # 刪除包含空值的行
        C_spectrum_Series = C_spectrum_Series.dropna()
        # 轉換為 float64 數據類型
        C_spectrum_Series = C_spectrum_Series.astype(float)


        if self.calculate_BLU() is not None:
            # 計算BLU------------------------------------------------------------------
            self.calculate_BLU()
            #print("self.calculate_BLU()",self.calculate_BLU())

            RxSxxl = CIE_spectrum_Series_X * self.calculate_BLU()
            RxSxyl = CIE_spectrum_Series_Y * self.calculate_BLU()
            RxSxzl = CIE_spectrum_Series_Z * self.calculate_BLU()

            BLU_spectrum_Series_sum = self.calculate_BLU().sum()
            k = 100 / BLU_spectrum_Series_sum
            RxSxxl_sum = RxSxxl.sum()
            RxSxyl_sum = RxSxyl.sum()
            RxSxzl_sum = RxSxzl.sum()
            RxSxxl_sum_k = RxSxxl_sum * k
            RxSxyl_sum_k = RxSxyl_sum * k
            RxSxzl_sum_k = RxSxzl_sum * k
            BLU_x = RxSxxl_sum_k / (RxSxxl_sum_k + RxSxyl_sum_k + RxSxzl_sum_k)
            BLU_y = RxSxyl_sum_k / (RxSxxl_sum_k + RxSxyl_sum_k + RxSxzl_sum_k)
            print("BLU_x", BLU_x)
            print("BLU_y", BLU_y)
            self.color_table.setItem(1, 14, QTableWidgetItem(f"{BLU_x:.3f}"))
            self.color_table.setItem(1, 15, QTableWidgetItem(f"{BLU_y:.3f}"))

        if self.calculate_BLU() is not None and self.calculate_RCF_Fix() is not None\
                and self.calculate_GCF_Fix() is not None and self.calculate_BCF_Fix() is not None:
            self.calculate_BLU()
            self.calculate_layer1()
            self.calculate_layer2()
            self.calculate_RCF_Fix()
            self.calculate_GCF_Fix()
            self.calculate_BCF_Fix()
            # BLU +Cell part
            cell_blu_total_spectrum = self.calculate_BLU() * self.calculate_layer1() * self.calculate_layer2()
            #print("self.calculate_BLU()",self.calculate_BLU())
            #print("cell_blu_total_spectrum",cell_blu_total_spectrum)
            # R_Fix
            R_Fix = cell_blu_total_spectrum * self.calculate_RCF_Fix()
            R_Fix_X = R_Fix * CIE_spectrum_Series_X
            R_Fix_Y = R_Fix * CIE_spectrum_Series_Y
            R_Fix_Z = R_Fix * CIE_spectrum_Series_Z
            R_Fix_X_sum = R_Fix_X.sum()
            R_Fix_Y_sum = R_Fix_Y.sum()
            R_Fix_Z_sum = R_Fix_Z.sum()
            R_Fix_x = R_Fix_X_sum / (R_Fix_X_sum + R_Fix_Y_sum + R_Fix_Z_sum)
            R_Fix_y = R_Fix_Y_sum / (R_Fix_X_sum + R_Fix_Y_sum + R_Fix_Z_sum)
            R_Fix_T = R_Fix_Y_sum * 3
            self.color_table.setItem(1, 4, QTableWidgetItem(f"{R_Fix_x:.3f}"))
            self.color_table.setItem(1, 5, QTableWidgetItem(f"{R_Fix_y:.3f}"))
            self.color_table.setItem(1, 6, QTableWidgetItem(f"{R_Fix_T:.3f}%"))

            # G_Fix
            G_Fix = cell_blu_total_spectrum * self.calculate_GCF_Fix()
            G_Fix_X = G_Fix * CIE_spectrum_Series_X
            G_Fix_Y = G_Fix * CIE_spectrum_Series_Y
            G_Fix_Z = G_Fix * CIE_spectrum_Series_Z
            G_Fix_X_sum = G_Fix_X.sum()
            G_Fix_Y_sum = G_Fix_Y.sum()
            G_Fix_Z_sum = G_Fix_Z.sum()
            G_Fix_x = G_Fix_X_sum / (G_Fix_X_sum + G_Fix_Y_sum + G_Fix_Z_sum)
            G_Fix_y = G_Fix_Y_sum / (G_Fix_X_sum + G_Fix_Y_sum + G_Fix_Z_sum)
            G_Fix_T = G_Fix_Y_sum * 3
            self.color_table.setItem(1, 7, QTableWidgetItem(f"{G_Fix_x:.3f}"))
            self.color_table.setItem(1, 8, QTableWidgetItem(f"{G_Fix_y:.3f}"))
            self.color_table.setItem(1, 9, QTableWidgetItem(f"{G_Fix_T:.3f}%"))

            # B_Fix
            B_Fix = cell_blu_total_spectrum * self.calculate_BCF_Fix()
            B_Fix_X = B_Fix * CIE_spectrum_Series_X
            B_Fix_Y = B_Fix * CIE_spectrum_Series_Y
            B_Fix_Z = B_Fix * CIE_spectrum_Series_Z
            B_Fix_X_sum = B_Fix_X.sum()
            B_Fix_Y_sum = B_Fix_Y.sum()
            B_Fix_Z_sum = B_Fix_Z.sum()
            B_Fix_x = B_Fix_X_sum / (B_Fix_X_sum + B_Fix_Y_sum + B_Fix_Z_sum)
            B_Fix_y = B_Fix_Y_sum / (B_Fix_X_sum + B_Fix_Y_sum + B_Fix_Z_sum)
            B_Fix_T = B_Fix_Y_sum * 3
            self.color_table.setItem(1, 10, QTableWidgetItem(f"{B_Fix_x:.3f}"))
            self.color_table.setItem(1, 11, QTableWidgetItem(f"{B_Fix_y:.3f}"))
            self.color_table.setItem(1, 12, QTableWidgetItem(f"{B_Fix_T:.3f}%"))

            # W_Fix
            W_Fix_X = R_Fix_X + G_Fix_X + B_Fix_X
            W_Fix_Y = R_Fix_Y + G_Fix_Y + B_Fix_Y
            W_Fix_Z = R_Fix_Z + G_Fix_Z + B_Fix_Z
            W_Fix_X_sum = W_Fix_X.sum()
            W_Fix_Y_sum = W_Fix_Y .sum()
            W_Fix_Z_sum = W_Fix_Z.sum()
            W_Fix_x = W_Fix_X_sum / (W_Fix_X_sum + W_Fix_Y_sum + W_Fix_Z_sum)
            W_Fix_y = W_Fix_Y_sum / (W_Fix_X_sum + W_Fix_Y_sum + W_Fix_Z_sum)
            W_Fix_T = W_Fix_Y_sum
            self.color_table.setItem(1, 1, QTableWidgetItem(f"{W_Fix_x:.3f}"))
            self.color_table.setItem(1, 2, QTableWidgetItem(f"{W_Fix_y:.3f}"))
            self.color_table.setItem(1, 3, QTableWidgetItem(f"{W_Fix_T:.3f}%"))

            #NTSC
            NTSC = 100 * 0.5 * abs((R_Fix_x * G_Fix_y + G_Fix_x * B_Fix_y + B_Fix_x * R_Fix_y-(G_Fix_x * R_Fix_y)-(B_Fix_x * G_Fix_y)-(R_Fix_x*B_Fix_y)))/ 0.1582
            self.color_table.setItem(1, 13, QTableWidgetItem(f"{NTSC:.3f}%"))

            # Clight +Cell part
            cell_C_total_spectrum = C_spectrum_Series * self.calculate_layer1() * self.calculate_layer2()
            # print("self.calculate_BLU()",self.calculate_BLU())
            # print("cell_blu_total_spectrum",cell_blu_total_spectrum)
            # RC_Fix
            RC_Fix = cell_C_total_spectrum * self.calculate_RCF_Fix()
            RC_Fix_X = RC_Fix * CIE_spectrum_Series_X
            RC_Fix_Y = RC_Fix * CIE_spectrum_Series_Y
            RC_Fix_Z = RC_Fix * CIE_spectrum_Series_Z
            RC_Fix_X_sum = RC_Fix_X.sum()
            RC_Fix_Y_sum = RC_Fix_Y.sum()
            RC_Fix_Z_sum = RC_Fix_Z.sum()
            RC_Fix_x = RC_Fix_X_sum / (RC_Fix_X_sum + RC_Fix_Y_sum + RC_Fix_Z_sum)
            RC_Fix_y = RC_Fix_Y_sum / (RC_Fix_X_sum + RC_Fix_Y_sum + RC_Fix_Z_sum)
            RC_Fix_T = RC_Fix_Y_sum * 3
            self.color_table.setItem(2, 4, QTableWidgetItem(f"{RC_Fix_x:.3f}"))
            self.color_table.setItem(2, 5, QTableWidgetItem(f"{RC_Fix_y:.3f}"))
            self.color_table.setItem(2, 6, QTableWidgetItem(f"{RC_Fix_T:.3f}%"))

            # GC_Fix
            GC_Fix = cell_blu_total_spectrum * self.calculate_GCF_Fix()
            GC_Fix_X = GC_Fix * CIE_spectrum_Series_X
            GC_Fix_Y = GC_Fix * CIE_spectrum_Series_Y
            GC_Fix_Z = GC_Fix * CIE_spectrum_Series_Z
            GC_Fix_X_sum = GC_Fix_X.sum()
            GC_Fix_Y_sum = GC_Fix_Y.sum()
            GC_Fix_Z_sum = GC_Fix_Z.sum()
            GC_Fix_x = GC_Fix_X_sum / (GC_Fix_X_sum + GC_Fix_Y_sum + GC_Fix_Z_sum)
            GC_Fix_y = G_Fix_Y_sum / (GC_Fix_X_sum + GC_Fix_Y_sum + GC_Fix_Z_sum)
            GC_Fix_T = G_Fix_Y_sum * 3
            self.color_table.setItem(2, 7, QTableWidgetItem(f"{GC_Fix_x:.3f}"))
            self.color_table.setItem(2, 8, QTableWidgetItem(f"{GC_Fix_y:.3f}"))
            self.color_table.setItem(2, 9, QTableWidgetItem(f"{GC_Fix_T:.3f}%"))

            # BC_Fix
            BC_Fix = cell_blu_total_spectrum * self.calculate_BCF_Fix()
            BC_Fix_X = BC_Fix * CIE_spectrum_Series_X
            BC_Fix_Y = BC_Fix * CIE_spectrum_Series_Y
            BC_Fix_Z = BC_Fix * CIE_spectrum_Series_Z
            BC_Fix_X_sum = BC_Fix_X.sum()
            BC_Fix_Y_sum = BC_Fix_Y.sum()
            BC_Fix_Z_sum = BC_Fix_Z.sum()
            BC_Fix_x = BC_Fix_X_sum / (BC_Fix_X_sum + BC_Fix_Y_sum + BC_Fix_Z_sum)
            BC_Fix_y = BC_Fix_Y_sum / (BC_Fix_X_sum + BC_Fix_Y_sum + BC_Fix_Z_sum)
            BC_Fix_T = BC_Fix_Y_sum * 3
            self.color_table.setItem(2, 10, QTableWidgetItem(f"{BC_Fix_x:.3f}"))
            self.color_table.setItem(2, 11, QTableWidgetItem(f"{BC_Fix_y:.3f}"))
            self.color_table.setItem(2, 12, QTableWidgetItem(f"{BC_Fix_T:.3f}%"))

            # W_Fix
            WC_Fix_X = RC_Fix_X + GC_Fix_X + BC_Fix_X
            WC_Fix_Y = RC_Fix_Y + GC_Fix_Y + BC_Fix_Y
            WC_Fix_Z = RC_Fix_Z + GC_Fix_Z + BC_Fix_Z
            WC_Fix_X_sum = WC_Fix_X.sum()
            WC_Fix_Y_sum = WC_Fix_Y.sum()
            WC_Fix_Z_sum = WC_Fix_Z.sum()
            WC_Fix_x = WC_Fix_X_sum / (WC_Fix_X_sum + WC_Fix_Y_sum + WC_Fix_Z_sum)
            WC_Fix_y = WC_Fix_Y_sum / (WC_Fix_X_sum + WC_Fix_Y_sum + WC_Fix_Z_sum)
            WC_Fix_T = WC_Fix_Y_sum
            self.color_table.setItem(2, 1, QTableWidgetItem(f"{WC_Fix_x:.3f}"))
            self.color_table.setItem(2, 2, QTableWidgetItem(f"{WC_Fix_y:.3f}"))
            self.color_table.setItem(2, 3, QTableWidgetItem(f"{WC_Fix_T:.3f}%"))

            # NTSCC
            NTSCC = 100 * 0.5 * abs((RC_Fix_x * GC_Fix_y + GC_Fix_x * BC_Fix_y + BC_Fix_x * RC_Fix_y - (GC_Fix_x * RC_Fix_y) - (
                        BC_Fix_x * GC_Fix_y) - (RC_Fix_x * BC_Fix_y))) / 0.1582
            self.color_table.setItem(2, 13, QTableWidgetItem(f"{NTSCC:.3f}%"))
            # 關閉連線
            connection_CIE.close()



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
        #print("headerlabels-from source", header_labels)
        self.light_source.clear()
        for item in header_labels:
            self.light_source.addItem(str(item))
            #print("item", str(item))
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
        #print("headerlabels-from source", header_labels)
        self.light_source_led.clear()
        for item in header_labels:
            self.light_source_led.addItem(str(item))
            #print("item", str(item))
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
        #print("headerlabels-from source", header_labels)
        self.source_led.clear()
        for item in header_labels:
            self.source_led.addItem(str(item))
            #print("item", str(item))
        # 關閉連線
        connection.close()
    # cell
    def update_layer1_datatable(self):
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("cell_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.layer1_table.clear()
        for table in tables:
            self.layer1_table.addItem(table[0])

        # 關閉連線
        conn.close()
    def updatelayer1ComboBox(self):
        connection = sqlite3.connect("cell_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.layer1_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        #print("headerlabels-from source", header_labels)
        self.layer1_box.clear()
        for item in header_labels:
            self.layer1_box.addItem(str(item))
            #print("item", str(item))
        # 關閉連線
        connection.close()
    # BSITO
    def update_layer2_datatable(self):
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("BSITO_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.layer2_table.clear()
        for table in tables:
            self.layer2_table.addItem(table[0])

        # 關閉連線
        conn.close()
    def updatelayer2ComboBox(self):
        connection = sqlite3.connect("BSITO_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.layer2_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        #print("headerlabels-from source", header_labels)
        self.layer2_box.clear()
        for item in header_labels:
            self.layer2_box.addItem(str(item))
            #print("item", str(item))
        # 關閉連線
        connection.close()


    # RCF_Fix
    def update_RCF_Fix_datatable(self):
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("RCF_Fix_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.R_fix_table.clear()
        for table in tables:
            self.R_fix_table.addItem(table[0])

        # 關閉連線
        conn.close()
    def updateRCF_Fix_ComboBox(self):
        connection = sqlite3.connect("RCF_Fix_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.R_fix_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        #print("headerlabels-from source", header_labels)
        self.R_fix_box.clear()
        for item in header_labels:
            self.R_fix_box.addItem(str(item))
            #print("item", str(item))
        # 關閉連線
        connection.close()

    # GCF_Fix
    def update_GCF_Fix_datatable(self):
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("GCF_Fix_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.G_fix_table.clear()
        for table in tables:
            self.G_fix_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def updateGCF_Fix_ComboBox(self):
        connection = sqlite3.connect("GCF_Fix_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.G_fix_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.G_fix_box.clear()
        for item in header_labels:
            self.G_fix_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    # BCF_Fix
    def update_BCF_Fix_datatable(self):
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("BCF_Fix_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.B_fix_table.clear()
        for table in tables:
            self.B_fix_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def updateBCF_Fix_ComboBox(self):
        connection = sqlite3.connect("BCF_Fix_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.B_fix_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.B_fix_box.clear()
        for item in header_labels:
            self.B_fix_box.addItem(str(item))
            # print("item", str(item))
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