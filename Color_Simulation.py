from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QHBoxLayout, QGridLayout, \
    QFormLayout, QLineEdit, QTabWidget, QTableWidgetItem, QTableWidget, QSizePolicy, QFrame, \
    QPushButton, QAbstractItemView, QComboBox, QPushButton, QCheckBox,QApplication,QDialog,QSpacerItem
from PySide6.QtGui import QKeyEvent, QColor, QPalette, QFont, QMouseEvent,QShortcut,QKeySequence,QClipboard
from PySide6.QtCore import Qt
from PySide6.QtCharts import QChart
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.patches as patches
import sqlite3
from Setting import *
# from Light_Source_BLU import *
import pandas as pd
import math
import numpy as np
from colour import xy_to_XYZ, XYZ_to_sRGB, CCS_ILLUMINANTS
from colour.plotting import plot_chromaticity_diagram_CIE1931, plot_single_colour_swatch
import colour
from colour.plotting import *
from colour import sd_single_led, plotting
from colour import sd_blackbody, SpectralShape
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backend_bases import MouseEvent
import re
from signal_manager import global_signal_manager
import itertools



class Color_Simulation(QWidget):
    def __init__(self):
        super().__init__()

        # 創建實例
        # color_result_table = Color_Result_Table()
        color_enter = Color_Enter()
        # color_select = Color_Select()

        # layout
        Color_Simulation_layout = QVBoxLayout()
        # Color_Simulation_layout.addWidget(color_result_table)
        Color_Simulation_layout.addWidget(color_enter)
        # Color_Simulation_layout.addWidget(color_select)

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
        # test signal
        self.check = "OK"
        self.BLUcheck = "OK"
        # 初始化属性
        self.W_x = None
        # 創建一個通用的 Ctrl+C 快捷鍵
        copy_shortcut = QShortcut(QKeySequence.Copy, self)
        copy_shortcut.activated.connect(self.copy_table_content)

        # 觀察者光源
        self.observer_D65 = [0.3127, 0.329]

        # color_table區
        self.color_table = QTableWidget()
        self.color_table.setColumnCount(22)
        self.color_table.setHorizontalHeaderLabels(["色度", "Wx", "Wy", "WY", "Rx", "Ry", "RY","Rλ","R_Purity", "Gx", "Gy", "GY",
                                  "Gλ","G_Purity","Bx", "By", "BY","B_λ","B_Purity", "NTSC%", "BLUx", "BLUy"])
        # 添加初始的行
        self.color_table.setRowCount(5)


        # 設定默認值
        column1_default_values = ["色度", "Wx", "Wy", "WY", "Rx", "Ry", "RY","Rλ","R_Purity", "Gx", "Gy", "GY",
                                  "Gλ","G_Purity","Bx", "By", "BY","B_λ","B_Purity", "NTSC%", "BLUx", "BLUy"]
        for row, values in enumerate([column1_default_values, ["整體模擬"], ["C-light"],["Item"],["材料資訊"]]):
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setBackground(QColor(173, 216, 230))  # 設置背景顏色為淺藍色
                item.setTextAlignment(Qt.AlignCenter)  # 設置文本居中對齊
                self.color_table.setItem(row, column, item)

        # 設定特定單元格顏色
        color_ranges = [(1, 3), (4, 8), (9, 13), (14, 18)]
        colors = [QColor(RESULTWHITE), QColor(RESULTRED), QColor(RESULTGREEN), QColor(RESULTBLUE)]

        for row in range(self.color_table.rowCount()):
            for (start_col, end_col), color in zip(color_ranges, colors):
                for col in range(start_col, end_col + 1):
                    item = self.color_table.item(row, col)
                    if item:
                        item.setBackground(color)

        Materialitems = ["R色阻選擇", "R色阻厚度", "G色阻選擇", "G色阻厚度", "B色阻選擇", "B色阻厚度",
                         "背光選擇","原LED選擇","替換LED","Layer_1","Layer_2","Layer_3","Layer_4","Layer_5","Layer_6"]
        for column, items in enumerate(Materialitems):
            item = QTableWidgetItem(items)
            item.setBackground(QColor("#5E86C1"))  # 青色
            self.color_table.setItem(3, column+1, item)


        # 設置表格可編輯
        self.color_table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.SelectedClicked)

        # Apply styles to the dialog
        self.color_table.setStyleSheet("""
                            QTableWidget::item:selected {
                        color: blcak; /* 設定文字顏色為黑色 */
                        background-color: #008080; /* 設定背景顏色為藍色，你可以根據需要調整 */
                            }
                        """)

        # Set Background
        self.setStyleSheet("background-color: lightblue;")

        # # 先創建dialog(這邊放圖)
        # self.picturedialog()

        # 總Layout
        self.Color_Enter_layout = QGridLayout()

        # Light source區域----------------------------------------------------------------------------------------------
        self.copy_button = QPushButton("點擊複製上方表格")
        self.light_source_mode = QComboBox()
        light_source_mode_box_items = ["未選", "自訂", "替換"]
        for item in light_source_mode_box_items:
            self.light_source_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.light_source_mode.setStyleSheet(QCOMBOXMODESETTING)
        # self.light_source_box.setBackgroundRole(QPalette.Window)

        self.light_source = QComboBox()
        self.light_source_datatable = QComboBox()
        self.light_source_datatable.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 防呆反灰設定
        self.light_source.setEnabled(False)
        self.light_source_datatable.setEnabled(False)
        self.light_source.setStyleSheet(QCOMBOBOXDISABLE)
        self.light_source_datatable.setStyleSheet(QCOMBOBOXDISABLE)

        # 更新light_source_table
        self.update_light_source_datatable()
        # 觸發table更新,lightsource表單
        self.updateLightSourceComboBox()  # 初始化
        self.light_source_datatable.currentIndexChanged.connect(self.updateLightSourceComboBox)

        # 觸發計算連動
        self.light_source_mode.currentTextChanged.connect(self.calculate_color_customize)
        self.light_source.currentTextChanged.connect(self.calculate_color_customize)
        self.light_source_datatable.currentTextChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖更新
        self.light_source_mode.currentTextChanged.connect(self.drawciechart)
        self.light_source.currentTextChanged.connect(self.drawciechart)
        self.light_source_datatable.currentTextChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.light_source_mode.currentTextChanged.connect(self.drawcie_WRGB_samplechart)
        self.light_source.currentTextChanged.connect(self.drawcie_WRGB_samplechart)
        self.light_source_datatable.currentTextChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog table
        self.light_source_mode.currentTextChanged.connect(self.set_wave_p)
        self.light_source.currentTextChanged.connect(self.set_wave_p)
        self.light_source_datatable.currentTextChanged.connect(self.set_wave_p)
        # 觸發WPC 更新
        self.light_source_mode.currentTextChanged.connect(self.calculate_WPC)
        self.light_source.currentTextChanged.connect(self.calculate_WPC)
        self.light_source_datatable.currentTextChanged.connect(self.calculate_WPC)

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
        # 防呆反灰設定
        self.light_source_led.setEnabled(False)
        self.light_source_led_datatable.setEnabled(False)
        self.light_source_led.setStyleSheet(QCOMBOBOXDISABLE)
        self.light_source_led_datatable.setStyleSheet(QCOMBOBOXDISABLE)
        # 更新light_source_led_table
        self.update_light_source_led_datatable()
        # 觸發table更新,lightsource_led表單
        self.updateLightSourceledComboBox()  # 初始化
        self.light_source_led_datatable.currentIndexChanged.connect(self.updateLightSourceledComboBox)
        # 觸發計算連動
        self.light_source_led.currentIndexChanged.connect(self.calculate_color_customize)
        self.light_source_led_datatable.currentIndexChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.light_source_led.currentIndexChanged.connect(self.drawciechart)
        self.light_source_led_datatable.currentIndexChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.light_source_led.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.light_source_led_datatable.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.light_source_led.currentIndexChanged.connect(self.set_wave_p)
        self.light_source_led_datatable.currentIndexChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.light_source_led.currentIndexChanged.connect(self.calculate_WPC)
        self.light_source_led_datatable.currentIndexChanged.connect(self.calculate_WPC)


        self.source_led = QComboBox()
        self.source_led_datatable = QComboBox()
        self.source_led_datatable.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 防呆反灰設定
        self.source_led .setEnabled(False)
        self.source_led_datatable.setEnabled(False)
        self.source_led .setStyleSheet(QCOMBOBOXDISABLE)
        self.source_led_datatable.setStyleSheet(QCOMBOBOXDISABLE)
        # 更新source_led_table
        self.update_source_led_datatable()
        # 觸發table更新,lightsource_led表單
        self.updateSourceledComboBox()  # 初始化
        self.source_led_datatable.currentIndexChanged.connect(self.updateSourceledComboBox)
        # 觸發計算連動
        self.source_led.currentIndexChanged.connect(self.calculate_color_customize)
        self.source_led_datatable.currentIndexChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.source_led.currentIndexChanged.connect(self.drawciechart)
        self.source_led_datatable.currentIndexChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.source_led.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.source_led_datatable.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.source_led.currentIndexChanged.connect(self.set_wave_p)
        self.source_led_datatable.currentIndexChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.source_led.currentIndexChanged.connect(self.calculate_WPC)
        self.source_led_datatable.currentIndexChanged.connect(self.calculate_WPC)

        self.label_source = QLabel("Light_Source_BLU")
        self.label_source_led = QLabel("Light_source_LED")
        self.label_led_data = QLabel("led_data_base")

        # Layer區---------------------------------------------------------------
        font = QFont()
        font.setPointSize(12)  # 設置字型大小
        font.setBold(True)  # 設置為粗體（可選）
        self.layer1_mode = QComboBox()
        layer1_mode_items = ["未選", "自訂"]
        for item in layer1_mode_items:
            self.layer1_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer1_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer1 = QLabel("Layer_1")
        self.label_layer1.setFont(font)
        # cell_頻譜選擇
        self.layer1_box = QComboBox()
        # layer1_table
        self.layer1_table = QComboBox()

        self.update_layer1_datatable()
        # 觸發layer1 table
        self.updatelayer1ComboBox()  # 初始化
        self.layer1_table.currentIndexChanged.connect(self.updatelayer1ComboBox)
        # 防呆反灰設定
        self.layer1_box.setEnabled(False)
        self.layer1_table.setEnabled(False)
        self.layer1_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.layer1_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 觸發計算連動
        self.layer1_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer1_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer1_table.currentIndexChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.layer1_mode.currentIndexChanged.connect(self.drawciechart)
        self.layer1_box.currentIndexChanged.connect(self.drawciechart)
        self.layer1_table.currentIndexChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.layer1_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer1_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer1_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.layer1_mode.currentIndexChanged.connect(self.set_wave_p)
        self.layer1_box.currentIndexChanged.connect(self.set_wave_p)
        self.layer1_table.currentIndexChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.layer1_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.layer1_box.currentIndexChanged.connect(self.calculate_WPC)
        self.layer1_table.currentIndexChanged.connect(self.calculate_WPC)

        self.layer2_mode = QComboBox()
        layer2_mode_items = ["未選", "自訂"]
        for item in layer2_mode_items:
            self.layer2_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer2_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer2 = QLabel("Layer_2")
        self.label_layer2.setFont(font)
        # cell_頻譜選擇
        self.layer2_box = QComboBox()
        # layer2_table
        self.layer2_table = QComboBox()
        self.update_layer2_datatable()
        # 觸發layer2 table
        self.updatelayer2ComboBox()  # 初始化
        self.layer2_table.currentIndexChanged.connect(self.updatelayer2ComboBox)
        # 防呆反灰設定
        self.layer2_box.setEnabled(False)
        self.layer2_table.setEnabled(False)
        self.layer2_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.layer2_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 計算觸發連動
        self.layer2_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer2_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer2_table.currentIndexChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.layer2_mode.currentIndexChanged.connect(self.drawciechart)
        self.layer2_box.currentIndexChanged.connect(self.drawciechart)
        self.layer2_table.currentIndexChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.layer2_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer2_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer2_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.layer2_mode.currentIndexChanged.connect(self.set_wave_p)
        self.layer2_box.currentIndexChanged.connect(self.set_wave_p)
        self.layer2_table.currentIndexChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.layer2_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.layer2_box.currentIndexChanged.connect(self.calculate_WPC)
        self.layer2_table.currentIndexChanged.connect(self.calculate_WPC)

        self.layer3_mode = QComboBox()
        layer3_mode_items = ["未選", "自訂"]
        for item in layer3_mode_items:
            self.layer3_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer3_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer3 = QLabel("Layer_3")
        self.label_layer3.setFont(font)
        self.layer3_box = QComboBox()
        # layer3 table
        self.layer3_table = QComboBox()
        self.update_layer3_datatable()
        # 防呆反灰設定
        self.layer3_box.setEnabled(False)
        self.layer3_table.setEnabled(False)
        self.layer3_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.layer3_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 觸發layer3 table
        self.updatelayer3ComboBox()  # 初始化
        self.layer3_table.currentIndexChanged.connect(self.updatelayer3ComboBox)
        # 計算觸發連動
        self.layer3_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer3_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer3_table.currentIndexChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.layer3_mode.currentIndexChanged.connect(self.drawciechart)
        self.layer3_box.currentIndexChanged.connect(self.drawciechart)
        self.layer3_table.currentIndexChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.layer3_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer3_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer3_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.layer3_mode.currentIndexChanged.connect(self.set_wave_p)
        self.layer3_box.currentIndexChanged.connect(self.set_wave_p)
        self.layer3_table.currentIndexChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.layer3_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.layer3_box.currentIndexChanged.connect(self.calculate_WPC)
        self.layer3_table.currentIndexChanged.connect(self.calculate_WPC)

        self.layer4_mode = QComboBox()
        layer4_mode_items = ["未選", "自訂"]
        for item in layer4_mode_items:
            self.layer4_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer4_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer4 = QLabel("Layer_4")
        self.label_layer4.setFont(font)
        self.layer4_box = QComboBox()
        # layer4 table
        self.layer4_table = QComboBox()
        self.update_layer4_datatable()
        # 防呆反灰設定
        self.layer4_box.setEnabled(False)
        self.layer4_table.setEnabled(False)
        self.layer4_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.layer4_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 觸發layer4 table
        self.updatelayer4ComboBox()  # 初始化
        self.layer4_table.currentIndexChanged.connect(self.updatelayer4ComboBox)
        # 計算觸發連動
        self.layer4_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer4_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer4_table.currentIndexChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.layer4_mode.currentIndexChanged.connect(self.drawciechart)
        self.layer4_box.currentIndexChanged.connect(self.drawciechart)
        self.layer4_table.currentIndexChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.layer4_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer4_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer4_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.layer4_mode.currentIndexChanged.connect(self.set_wave_p)
        self.layer4_box.currentIndexChanged.connect(self.set_wave_p)
        self.layer4_table.currentIndexChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.layer4_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.layer4_box.currentIndexChanged.connect(self.calculate_WPC)
        self.layer4_table.currentIndexChanged.connect(self.calculate_WPC)

        self.layer5_mode = QComboBox()
        layer5_mode_items = ["未選", "自訂"]
        for item in layer5_mode_items:
            self.layer5_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer5_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer5 = QLabel("Layer_5")
        self.label_layer5.setFont(font)
        self.layer5_box = QComboBox()
        # layer5 table
        self.layer5_table = QComboBox()
        self.update_layer5_datatable()
        # 防呆反灰設定
        self.layer5_box.setEnabled(False)
        self.layer5_table.setEnabled(False)
        self.layer5_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.layer5_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 觸發layer5 table
        self.updatelayer5ComboBox()  # 初始化
        self.layer5_table.currentIndexChanged.connect(self.updatelayer5ComboBox)
        # 計算觸發連動
        self.layer5_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer5_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer5_table.currentIndexChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.layer5_mode.currentIndexChanged.connect(self.drawciechart)
        self.layer5_box.currentIndexChanged.connect(self.drawciechart)
        self.layer5_table.currentIndexChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.layer5_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer5_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer5_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.layer5_mode.currentIndexChanged.connect(self.set_wave_p)
        self.layer5_box.currentIndexChanged.connect(self.set_wave_p)
        self.layer5_table.currentIndexChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.layer5_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.layer5_box.currentIndexChanged.connect(self.calculate_WPC)
        self.layer5_table.currentIndexChanged.connect(self.calculate_WPC)

        self.layer6_mode = QComboBox()
        layer6_mode_items = ["未選", "自訂"]
        for item in layer6_mode_items:
            self.layer6_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.layer6_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.label_layer6 = QLabel("Layer_6")
        self.label_layer6.setFont(font)
        self.layer6_box = QComboBox()
        # layer6 table
        self.layer6_table = QComboBox()
        self.update_layer6_datatable()
        # 防呆反灰設定
        self.layer6_box.setEnabled(False)
        self.layer6_table.setEnabled(False)
        self.layer6_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.layer6_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 觸發layer6 table
        self.updatelayer6ComboBox()  # 初始化
        self.layer6_table.currentIndexChanged.connect(self.updatelayer6ComboBox)
        # 計算觸發連動
        self.layer6_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer6_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.layer6_table.currentIndexChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.layer6_mode.currentIndexChanged.connect(self.drawciechart)
        self.layer6_box.currentIndexChanged.connect(self.drawciechart)
        self.layer6_table.currentIndexChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.layer6_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer6_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.layer6_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.layer6_mode.currentIndexChanged.connect(self.set_wave_p)
        self.layer6_box.currentIndexChanged.connect(self.set_wave_p)
        self.layer6_table.currentIndexChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.layer6_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.layer6_box.currentIndexChanged.connect(self.calculate_WPC)
        self.layer6_table.currentIndexChanged.connect(self.calculate_WPC)
        # RGB-Fix區域--------------------------------------------------------------------------------------------------
        self.RGB_fix_label = QLabel("RGB-Fix")
        self.RGB_fix_label.setStyleSheet("color: #5151A2; font-weight: bold; border: 2px solid black;")

        self.R_fix_mode = QComboBox()
        R_fix_mode_items = ["未選", "自訂"]
        for item in R_fix_mode_items:
            self.R_fix_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.R_fix_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.R_fix_label = QLabel("R-CF-Fix")
        self.R_fix_box = QComboBox()
        self.R_fix_table = QComboBox()
        # 防呆反灰設定
        self.R_fix_box.setEnabled(False)
        self.R_fix_table.setEnabled(False)
        self.R_fix_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.R_fix_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 觸發table連動改變
        self.update_RCF_Fix_datatable()
        self.updateRCF_Fix_ComboBox()
        self.R_fix_table.currentIndexChanged.connect(self.updateRCF_Fix_ComboBox)
        self.R_fix_mode.currentIndexChanged.connect(self.update_RCF_Fix_modeclose)

        self.R_TK_edit_label = QLabel("R-Fix-TK")
        # self.R_TK_edit = QLineEdit()
        # self.R_TK_edit.setFixedSize(100, 25)
        self.R_TK_label = QLabel()
        # 觸發連動計算
        self.R_fix_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.R_fix_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.R_fix_table.currentIndexChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.R_fix_mode.currentIndexChanged.connect(self.drawciechart)
        self.R_fix_box.currentIndexChanged.connect(self.drawciechart)
        self.R_fix_table.currentIndexChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.R_fix_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.R_fix_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.R_fix_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發算dialog_table
        self.R_fix_mode.currentIndexChanged.connect(self.set_wave_p)
        self.R_fix_box.currentIndexChanged.connect(self.set_wave_p)
        self.R_fix_table.currentIndexChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.R_fix_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.R_fix_box.currentIndexChanged.connect(self.calculate_WPC)
        self.R_fix_table.currentIndexChanged.connect(self.calculate_WPC)

        # G-Fix
        self.G_fix_mode = QComboBox()
        G_fix_mode_items = ["未選", "自訂"]
        for item in G_fix_mode_items:
            self.G_fix_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.G_fix_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.G_fix_label = QLabel("G-CF-Fix")
        self.G_fix_box = QComboBox()
        self.G_fix_table = QComboBox()
        # 防呆反灰設定
        self.G_fix_box.setEnabled(False)
        self.G_fix_table.setEnabled(False)
        self.G_fix_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.G_fix_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 觸發table連動改變
        self.update_GCF_Fix_datatable()
        self.updateGCF_Fix_ComboBox()
        self.G_fix_table.currentIndexChanged.connect(self.updateGCF_Fix_ComboBox)
        self.G_fix_mode.currentIndexChanged.connect(self.update_GCF_Fix_modeclose)

        self.G_TK_edit_label = QLabel("G-Fix-TK")
        # self.G_TK_edit = QLineEdit()
        # self.G_TK_edit.setFixedSize(100, 25)
        self.G_TK_label = QLabel()
        # 觸發連動計算
        self.G_fix_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.G_fix_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.G_fix_table.currentIndexChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.G_fix_mode.currentIndexChanged.connect(self.drawciechart)
        self.G_fix_box.currentIndexChanged.connect(self.drawciechart)
        self.G_fix_table.currentIndexChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.G_fix_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.G_fix_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.G_fix_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.G_fix_mode.currentIndexChanged.connect(self.set_wave_p)
        self.G_fix_box.currentIndexChanged.connect(self.set_wave_p)
        self.G_fix_table.currentIndexChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.G_fix_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.G_fix_box.currentIndexChanged.connect(self.calculate_WPC)
        self.G_fix_table.currentIndexChanged.connect(self.calculate_WPC)

        # B-Fix
        self.B_fix_mode = QComboBox()
        B_fix_mode_items = ["未選", "自訂"]
        for item in B_fix_mode_items:
            self.B_fix_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.B_fix_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.B_fix_label = QLabel("B-CF-Fix")
        self.B_fix_box = QComboBox()
        self.B_fix_table = QComboBox()
        # 防呆反灰設定
        self.B_fix_box.setEnabled(False)
        self.B_fix_table.setEnabled(False)
        self.B_fix_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.B_fix_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 觸發table連動改變
        self.update_BCF_Fix_datatable()
        self.updateBCF_Fix_ComboBox()
        self.B_fix_table.currentIndexChanged.connect(self.updateBCF_Fix_ComboBox)
        self.B_fix_mode.currentIndexChanged.connect(self.update_BCF_Fix_modeclose)

        self.B_TK_edit_label = QLabel("B-Fix-TK")
        # self.B_TK_edit = QLineEdit()
        # self.B_TK_edit.setFixedSize(100, 25)
        self.B_TK_label = QLabel()
        # 觸發連動計算
        self.B_fix_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.B_fix_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.B_fix_table.currentIndexChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.B_fix_mode.currentIndexChanged.connect(self.drawciechart)
        self.B_fix_box.currentIndexChanged.connect(self.drawciechart)
        self.B_fix_table.currentIndexChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.B_fix_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.B_fix_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.B_fix_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.B_fix_mode.currentIndexChanged.connect(self.set_wave_p)
        self.B_fix_box.currentIndexChanged.connect(self.set_wave_p)
        self.B_fix_table.currentIndexChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.B_fix_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.B_fix_box.currentIndexChanged.connect(self.calculate_WPC)
        self.B_fix_table.currentIndexChanged.connect(self.calculate_WPC)

        # RGB-α,K 區域---------------------------------------------------
        self.RGB_aK_label = QLabel("RGB-α,K")
        self.RGB_aK_label.setStyleSheet("color: #5151A2; font-weight: bold; border: 2px solid black;")
        self.R_aK_mode = QComboBox()
        R_aK_mode_items = ["未選", "自訂"]
        for item in R_aK_mode_items:
            self.R_aK_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.R_aK_mode.setStyleSheet(QCOMBOXMODESETTING)
        # self.R_aK_label = QLabel("R-CF-α,K")
        self.R_aK_box = QComboBox()
        self.R_aK_table = QComboBox()
        # 防呆反灰設定
        self.R_aK_box.setEnabled(False)
        self.R_aK_table.setEnabled(False)
        self.R_aK_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.R_aK_table.setStyleSheet(QCOMBOBOXDISABLE)

        # 觸發table連動改變
        self.update_RCF_Change_datatable()
        self.updateRCF_Change_ComboBox()
        self.R_aK_table.currentIndexChanged.connect(self.updateRCF_Change_ComboBox)
        self.R_aK_mode.currentIndexChanged.connect(self.update_RCF_Change_modeclose)

        self.R_aK_TK_edit_label = QLabel("R-α,K-TK")
        self.R_aK_TK_edit = QLineEdit()
        self.R_aK_TK_edit.setFixedSize(100, 25)
        # 觸發連動計算
        self.R_aK_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.R_aK_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.R_aK_table.currentIndexChanged.connect(self.calculate_color_customize)
        self.R_aK_TK_edit.textChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.R_aK_mode.currentIndexChanged.connect(self.drawciechart)
        self.R_aK_box.currentIndexChanged.connect(self.drawciechart)
        self.R_aK_table.currentIndexChanged.connect(self.drawciechart)
        self.R_aK_TK_edit.textChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.R_aK_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.R_aK_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.R_aK_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.R_aK_TK_edit.textChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.R_aK_mode.currentIndexChanged.connect(self.set_wave_p)
        self.R_aK_box.currentIndexChanged.connect(self.set_wave_p)
        self.R_aK_table.currentIndexChanged.connect(self.set_wave_p)
        self.R_aK_TK_edit.textChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.R_aK_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.R_aK_box.currentIndexChanged.connect(self.calculate_WPC)
        self.R_aK_table.currentIndexChanged.connect(self.calculate_WPC)
        self.R_aK_TK_edit.textChanged.connect(self.calculate_WPC)
        # G-α,K
        self.G_aK_mode = QComboBox()
        G_aK_mode_items = ["未選", "自訂"]
        for item in G_aK_mode_items:
            self.G_aK_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.G_aK_mode.setStyleSheet(QCOMBOXMODESETTING)
        # self.G_aK_label = QLabel("G-F-α,K")
        self.G_aK_box = QComboBox()
        self.G_aK_table = QComboBox()
        # 防呆反灰設定
        self.G_aK_box.setEnabled(False)
        self.G_aK_table.setEnabled(False)
        self.G_aK_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.G_aK_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 觸發table連動改變
        self.update_GCF_Change_datatable()
        self.updateGCF_Change_ComboBox()
        self.G_aK_table.currentIndexChanged.connect(self.updateGCF_Change_ComboBox)
        self.G_aK_mode.currentIndexChanged.connect(self.update_GCF_Change_modeclose)

        self.G_aK_TK_edit_label = QLabel("G-α,K-TK")
        self.G_aK_TK_edit = QLineEdit()
        self.G_aK_TK_edit.setFixedSize(100, 25)
        # 觸發連動計算
        self.G_aK_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.G_aK_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.G_aK_table.currentIndexChanged.connect(self.calculate_color_customize)
        self.G_aK_TK_edit.textChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.G_aK_mode.currentIndexChanged.connect(self.drawciechart)
        self.G_aK_box.currentIndexChanged.connect(self.drawciechart)
        self.G_aK_table.currentIndexChanged.connect(self.drawciechart)
        self.G_aK_TK_edit.textChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.G_aK_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.G_aK_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.G_aK_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.G_aK_TK_edit.textChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.G_aK_mode.currentIndexChanged.connect(self.set_wave_p)
        self.G_aK_box.currentIndexChanged.connect(self.set_wave_p)
        self.G_aK_table.currentIndexChanged.connect(self.set_wave_p)
        self.G_aK_TK_edit.textChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.G_aK_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.G_aK_box.currentIndexChanged.connect(self.calculate_WPC)
        self.G_aK_table.currentIndexChanged.connect(self.calculate_WPC)
        self.G_aK_TK_edit.textChanged.connect(self.calculate_WPC)
        # B-α,K
        self.B_aK_mode = QComboBox()
        B_aK_mode_items = ["未選", "自訂"]
        for item in B_aK_mode_items:
            self.B_aK_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.B_aK_mode.setStyleSheet(QCOMBOXMODESETTING)
        # self.B_aK_label = QLabel("B-CF-α,K")
        self.B_aK_box = QComboBox()
        self.B_aK_table = QComboBox()
        # 防呆反灰設定
        self.B_aK_box.setEnabled(False)
        self.B_aK_table.setEnabled(False)
        self.B_aK_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.B_aK_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 觸發table連動改變
        self.update_BCF_Change_datatable()
        self.updateBCF_Change_ComboBox()
        self.B_aK_table.currentIndexChanged.connect(self.updateBCF_Change_ComboBox)
        self.B_aK_mode.currentIndexChanged.connect(self.update_BCF_Change_modeclose)

        self.B_aK_TK_edit_label = QLabel("B-α,K-TK")
        self.B_aK_TK_edit = QLineEdit()
        self.B_aK_TK_edit.setFixedSize(100, 25)
        # 觸發連動計算
        self.B_aK_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.B_aK_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.B_aK_table.currentIndexChanged.connect(self.calculate_color_customize)
        self.B_aK_TK_edit.textChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.B_aK_mode.currentIndexChanged.connect(self.drawciechart)
        self.B_aK_box.currentIndexChanged.connect(self.drawciechart)
        self.B_aK_table.currentIndexChanged.connect(self.drawciechart)
        self.B_aK_TK_edit.textChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.B_aK_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.B_aK_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.B_aK_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.B_aK_TK_edit.textChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發dialog_table
        self.B_aK_mode.currentIndexChanged.connect(self.set_wave_p)
        self.B_aK_box.currentIndexChanged.connect(self.set_wave_p)
        self.B_aK_table.currentIndexChanged.connect(self.set_wave_p)
        self.B_aK_TK_edit.textChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.B_aK_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.B_aK_box.currentIndexChanged.connect(self.calculate_WPC)
        self.B_aK_table.currentIndexChanged.connect(self.calculate_WPC)
        self.B_aK_TK_edit.textChanged.connect(self.calculate_WPC)

        # RGB-內差法 區域---------------------------------------------------
        self.RGB_differ_label = QLabel("RGB-頻譜內外差法")
        self.RGB_differ_label.setStyleSheet("color: #5151A2; font-weight: bold; border: 2px solid black;")
        self.R_differ_mode = QComboBox()
        R_differ_mode_items = ["未選", "自訂"]
        for item in R_differ_mode_items:
            self.R_differ_mode.addItem(str(item))
        self.R_differ_mode.setStyleSheet(QCOMBOXMODESETTING)
        #self.R_differ_label = QLabel("R-CF-differ")
        self.R_differ_box = QComboBox()
        self.R_differ_box.setStyleSheet(QCOMBOXSETTING)
        self.R_differ_TK_edit_label = QLabel("R-differ-TK")
        self.R_differ_TK_edit = QLineEdit()
        self.R_differ_TK_edit.setFixedSize(100, 25)
        # R_differ_table
        self.R_differ_table = QComboBox()
        # 防呆反灰設定
        self.R_differ_box.setEnabled(False)
        self.R_differ_table.setEnabled(False)
        self.R_differ_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.R_differ_table.setStyleSheet(QCOMBOBOXDISABLE)


        self.update_RCF_Differ_datatable()
        # 觸發R_differ_table
        self.updateRCF_Differ_ComboBox()# 初始化
        self.R_differ_table.currentIndexChanged.connect(self.updateRCF_Differ_ComboBox)
        self.R_differ_mode.currentIndexChanged.connect(self.update_RCF_Differ_modeclose)
        # 觸發連動計算
        self.R_differ_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.R_differ_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.R_differ_table.currentIndexChanged.connect(self.calculate_color_customize)
        self.R_differ_TK_edit.textChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.R_differ_mode.currentIndexChanged.connect(self.drawciechart)
        self.R_differ_box.currentIndexChanged.connect(self.drawciechart)
        self.R_differ_table.currentIndexChanged.connect(self.drawciechart)
        self.R_differ_TK_edit.textChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.R_differ_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.R_differ_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.R_differ_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.R_differ_TK_edit.textChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發算dialog_table
        self.R_differ_mode.currentIndexChanged.connect(self.set_wave_p)
        self.R_differ_box.currentIndexChanged.connect(self.set_wave_p)
        self.R_differ_table.currentIndexChanged.connect(self.set_wave_p)
        self.R_differ_TK_edit.textChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.R_differ_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.R_differ_box.currentIndexChanged.connect(self.calculate_WPC)
        self.R_differ_table.currentIndexChanged.connect(self.calculate_WPC)
        self.R_differ_TK_edit.textChanged.connect(self.calculate_WPC)

        # G-differ
        self.G_differ_mode = QComboBox()
        G_differ_mode_items = ["未選", "自訂"]
        for item in G_differ_mode_items:
            self.G_differ_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.G_differ_mode.setStyleSheet(QCOMBOXMODESETTING)
        #self.G_differ_label = QLabel("G-CF-differ")
        self.G_differ_box = QComboBox()
        # 設定當前選中項目的文字顏色
        self.G_differ_TK_edit_label = QLabel("G-differ-TK")
        self.G_differ_TK_edit = QLineEdit()
        self.G_differ_TK_edit.setFixedSize(100, 25)
        # G_differ_table
        self.G_differ_table = QComboBox()
        self.update_GCF_Differ_datatable()
        # 防呆反灰設定
        self.G_differ_box.setEnabled(False)
        self.G_differ_table.setEnabled(False)
        self.G_differ_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.G_differ_table.setStyleSheet(QCOMBOBOXDISABLE)
        # 觸發R_differ_table
        self.updateGCF_Differ_ComboBox()  # 初始化
        self.G_differ_table.currentIndexChanged.connect(self.updateGCF_Differ_ComboBox)
        self.G_differ_mode.currentIndexChanged.connect(self.update_GCF_Differ_modeclose)
        # 觸發連動計算
        self.G_differ_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.G_differ_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.G_differ_table.currentIndexChanged.connect(self.calculate_color_customize)
        self.G_differ_TK_edit.textChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.G_differ_mode.currentIndexChanged.connect(self.drawciechart)
        self.G_differ_box.currentIndexChanged.connect(self.drawciechart)
        self.G_differ_table.currentIndexChanged.connect(self.drawciechart)
        self.G_differ_TK_edit.textChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.G_differ_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.G_differ_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.G_differ_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.G_differ_TK_edit.textChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發算dialog_table
        self.G_differ_mode.currentIndexChanged.connect(self.set_wave_p)
        self.G_differ_box.currentIndexChanged.connect(self.set_wave_p)
        self.G_differ_table.currentIndexChanged.connect(self.set_wave_p)
        self.G_differ_TK_edit.textChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.G_differ_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.G_differ_box.currentIndexChanged.connect(self.calculate_WPC)
        self.G_differ_table.currentIndexChanged.connect(self.calculate_WPC)
        self.G_differ_TK_edit.textChanged.connect(self.calculate_WPC)

        # B-set
        self.B_differ_mode = QComboBox()
        B_set_mode_items = ["未選", "自訂"]
        for item in B_set_mode_items:
            self.B_differ_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.B_differ_mode.setStyleSheet(QCOMBOXMODESETTING)
        #self.B_differ_label = QLabel("B-CF-differ")
        self.B_differ_box = QComboBox()
        # 設定當前選中項目的文字顏色
        self.B_differ_TK_edit_label = QLabel("B-differ-TK")
        self.B_differ_TK_edit = QLineEdit()
        self.B_differ_TK_edit.setFixedSize(100, 25)
        # B_differ_table
        self.B_differ_table = QComboBox()
        # 防呆反灰設定
        self.B_differ_box.setEnabled(False)
        self.B_differ_table.setEnabled(False)
        self.B_differ_box.setStyleSheet(QCOMBOBOXDISABLE)
        self.B_differ_table.setStyleSheet(QCOMBOBOXDISABLE)
        self.update_BCF_Differ_datatable()
        # 觸發R_differ_table
        self.updateBCF_Differ_ComboBox()  # 初始化
        self.B_differ_table.currentIndexChanged.connect(self.updateBCF_Differ_ComboBox)
        self.B_differ_mode.currentIndexChanged.connect(self.update_BCF_Differ_modeclose)
        # 觸發連動計算
        self.B_differ_mode.currentIndexChanged.connect(self.calculate_color_customize)
        self.B_differ_box.currentIndexChanged.connect(self.calculate_color_customize)
        self.B_differ_table.currentIndexChanged.connect(self.calculate_color_customize)
        self.B_differ_TK_edit.textChanged.connect(self.calculate_color_customize)
        # 觸發CIE圖
        self.B_differ_mode.currentIndexChanged.connect(self.drawciechart)
        self.B_differ_box.currentIndexChanged.connect(self.drawciechart)
        self.B_differ_table.currentIndexChanged.connect(self.drawciechart)
        self.B_differ_TK_edit.textChanged.connect(self.drawciechart)
        # 觸發Sample圖
        self.B_differ_mode.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.B_differ_box.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.B_differ_table.currentIndexChanged.connect(self.drawcie_WRGB_samplechart)
        self.B_differ_TK_edit.textChanged.connect(self.drawcie_WRGB_samplechart)
        # 觸發算dialog_table
        self.B_differ_mode.currentIndexChanged.connect(self.set_wave_p)
        self.B_differ_box.currentIndexChanged.connect(self.set_wave_p)
        self.B_differ_table.currentIndexChanged.connect(self.set_wave_p)
        self.B_differ_TK_edit.textChanged.connect(self.set_wave_p)
        # 觸發WPC
        self.B_differ_mode.currentIndexChanged.connect(self.calculate_WPC)
        self.B_differ_box.currentIndexChanged.connect(self.calculate_WPC)
        self.B_differ_table.currentIndexChanged.connect(self.calculate_WPC)
        self.B_differ_TK_edit.textChanged.connect(self.calculate_WPC)


        # 計算Button
        self.calculate = QPushButton("Jump CIE plot")
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
        self.refresh = QPushButton("refresh_all_data_table")
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

        # 臨時test button
        self.test_button = QPushButton("Test")

        # 設置游標樣式
        self.refresh.setCursor(Qt.PointingHandCursor)  # 手指形狀

        # 按鈕功能設定-----------------------------------------------------------
        self.refresh.clicked.connect(self.updateall)
        self.calculate.clicked.connect(self.picturedialog)
        self.test_button.clicked.connect(self.Color_Simulation)
        self.copy_button.clicked.connect(self.copy_entire_table_content)

        # WPC專區
        self.WPClabel = QLabel("WPC Target")
        self.WPCCombobox = QComboBox()
        WPCcomboboxitems = ["WPC-BM遮擋設計","WPC-不等開口率設計"]
        for item in WPCcomboboxitems:
            self.WPCCombobox.addItem(item)
        self.WPCx = QLineEdit()
        self.WPCx.setPlaceholderText("WPCx_Target")  # 反灰提醒用
        self.WPCy = QLineEdit()
        self.WPCy.setPlaceholderText("WPCy_Target")  # 反灰提醒用
        # 創建table
        self.WPCTable = QTableWidget()
        self.WPCTable.setColumnCount(8)
        self.WPCTable.setRowCount(7)
        self.WPCTable.setFixedSize(620,300)

        # 新增設定顏色的程式碼
        color_items = ["R", "G", "B","T%","開口率修正T%"]
        color_values = [QColor(RESULTRED), QColor(RESULTGREEN), QColor(RESULTBLUE),QColor(RESULTWHITE),QColor(RESULTWHITE)]

        for column, (color_item, color_value) in enumerate(zip(color_items, color_values)):
            item = QTableWidgetItem(color_item)
            item.setBackground(color_value)
            item.setTextAlignment(Qt.AlignCenter)
            self.WPCTable.setItem(0, column+1, item)

        WPCitems = ["Item","遮BMDesign","不等開口率","Item_反推","自行輸入比例","總-原開口率(%)","總-開口率修正(%)"]
        for row ,values in enumerate(WPCitems):
            item = QTableWidgetItem(values)
            self.WPCTable.setItem(row, 0, item)

        # 新增設定顏色的程式碼
        color_items_2 = ["R", "G", "B", "Wx","Wy","T%","開口率修正T%"]
        color_values_2 = [QColor(RESULTRED), QColor(RESULTGREEN), QColor(RESULTBLUE), QColor(RESULTWHITE),QColor(RESULTWHITE),QColor(RESULTWHITE),QColor(RESULTWHITE)]

        for column, (color_item, color_value) in enumerate(zip(color_items_2, color_values_2)):
            item = QTableWidgetItem(color_item)
            item.setBackground(color_value)
            item.setTextAlignment(Qt.AlignCenter)
            self.WPCTable.setItem(3, column + 1, item)

        # 做連動
        self.WPCx.textChanged.connect(self.calculate_WPC)
        self.WPCy.textChanged.connect(self.calculate_WPC)
        self.WPCTable.cellChanged.connect(self.calculate_WPC_reverse)
        self.WPCTable.cellChanged.connect(self.calculate_WPC_trs_ratio)
        self.WPCTable.cellChanged.connect(self.calculate_WPC_trs_ratio_2)


        # widget放置
        # color table
        self.Color_Enter_layout.addWidget(self.color_table, 0, 0, 4, 10)
        # Light source區域---------------------------------------
        self.Color_Enter_layout.addWidget(self.copy_button,4,0,1,1)
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
        self.Color_Enter_layout.addWidget(self.test_button, 5, 8, 1, 1)
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
        self.Color_Enter_layout.addWidget(self.label_layer3, 6, 6, 1, 2)
        self.Color_Enter_layout.addWidget(self.layer3_table, 6, 8, 1, 1)

        self.Color_Enter_layout.addWidget(self.layer4_mode, 9, 0)
        self.Color_Enter_layout.addWidget(self.layer4_box, 9, 1, 1, 2)
        self.Color_Enter_layout.addWidget(self.label_layer4, 8, 0, 1, 2)
        self.Color_Enter_layout.addWidget(self.layer4_table, 8, 2, 1, 1)

        self.Color_Enter_layout.addWidget(self.layer5_mode, 9, 3)
        self.Color_Enter_layout.addWidget(self.layer5_box, 9, 4, 1, 2)
        self.Color_Enter_layout.addWidget(self.label_layer5, 8, 3, 1, 2)
        self.Color_Enter_layout.addWidget(self.layer5_table, 8, 5, 1, 1)

        self.Color_Enter_layout.addWidget(self.layer6_mode, 9, 6)
        self.Color_Enter_layout.addWidget(self.layer6_box, 9, 7, 1, 2)
        self.Color_Enter_layout.addWidget(self.label_layer6, 8, 6, 1, 2)
        self.Color_Enter_layout.addWidget(self.layer6_table, 8, 8, 1, 1)

        # RGB-Fix區域---------------------------------------------------
        self.Color_Enter_layout.addWidget(self.RGB_fix_label, 10, 0)
        self.Color_Enter_layout.addWidget(self.R_fix_mode, 11, 0)
        self.Color_Enter_layout.addWidget(self.R_fix_box, 11, 1)
        self.Color_Enter_layout.addWidget(self.R_fix_table, 10, 1)
        self.Color_Enter_layout.addWidget(self.R_TK_edit_label, 10, 2)
        self.Color_Enter_layout.addWidget(self.R_TK_label, 11, 2)

        self.Color_Enter_layout.addWidget(self.G_fix_mode, 11, 3)
        self.Color_Enter_layout.addWidget(self.G_fix_box, 11, 4)
        self.Color_Enter_layout.addWidget(self.G_fix_table, 10, 4)
        self.Color_Enter_layout.addWidget(self.G_TK_edit_label, 10, 5)
        self.Color_Enter_layout.addWidget(self.G_TK_label, 11, 5)

        self.Color_Enter_layout.addWidget(self.B_fix_mode, 11, 6)
        self.Color_Enter_layout.addWidget(self.B_fix_box, 11, 7)
        self.Color_Enter_layout.addWidget(self.B_fix_table, 10, 7)
        self.Color_Enter_layout.addWidget(self.B_TK_edit_label, 10, 8)
        self.Color_Enter_layout.addWidget(self.B_TK_label, 11, 8)

        # RGB-α,K區域---------------------------------------------------
        self.Color_Enter_layout.addWidget(self.RGB_aK_label, 12, 0)
        self.Color_Enter_layout.addWidget(self.R_aK_mode, 13, 0)
        self.Color_Enter_layout.addWidget(self.R_aK_box, 13, 1)
        self.Color_Enter_layout.addWidget(self.R_aK_table, 12, 1)
        self.Color_Enter_layout.addWidget(self.R_aK_TK_edit_label, 12, 2)
        self.Color_Enter_layout.addWidget(self.R_aK_TK_edit, 13, 2)

        self.Color_Enter_layout.addWidget(self.G_aK_mode, 13, 3)
        self.Color_Enter_layout.addWidget(self.G_aK_box, 13, 4)
        self.Color_Enter_layout.addWidget(self.G_aK_table, 12, 4)
        self.Color_Enter_layout.addWidget(self.G_aK_TK_edit_label, 12, 5)
        self.Color_Enter_layout.addWidget(self.G_aK_TK_edit, 13, 5)

        self.Color_Enter_layout.addWidget(self.B_aK_mode, 13, 6)
        self.Color_Enter_layout.addWidget(self.B_aK_box, 13, 7)
        self.Color_Enter_layout.addWidget(self.B_aK_table, 12, 7)
        self.Color_Enter_layout.addWidget(self.B_aK_TK_edit_label, 12, 8)
        self.Color_Enter_layout.addWidget(self.B_aK_TK_edit, 13, 8)

        # RGB-differ區域---------------------------------------------------
        self.Color_Enter_layout.addWidget(self.RGB_differ_label, 14, 0)
        self.Color_Enter_layout.addWidget(self.R_differ_mode, 15, 0)
        self.Color_Enter_layout.addWidget(self.R_differ_box, 15, 1)
        self.Color_Enter_layout.addWidget(self.R_differ_table, 14, 1)
        self.Color_Enter_layout.addWidget(self.R_differ_TK_edit_label, 14, 2)
        self.Color_Enter_layout.addWidget(self.R_differ_TK_edit, 15, 2)

        self.Color_Enter_layout.addWidget(self.G_differ_mode, 15, 3)
        self.Color_Enter_layout.addWidget(self.G_differ_box, 15, 4)
        self.Color_Enter_layout.addWidget(self.G_differ_table, 14, 4)
        self.Color_Enter_layout.addWidget(self.G_differ_TK_edit_label, 14, 5)
        self.Color_Enter_layout.addWidget(self.G_differ_TK_edit, 15, 5)

        self.Color_Enter_layout.addWidget(self.B_differ_mode, 15, 6)
        self.Color_Enter_layout.addWidget(self.B_differ_box, 15, 7)
        self.Color_Enter_layout.addWidget(self.B_differ_table, 14, 7)
        self.Color_Enter_layout.addWidget(self.B_differ_TK_edit_label, 14, 8)
        self.Color_Enter_layout.addWidget(self.B_differ_TK_edit, 15, 8)

        # WPC區
        self.Color_Enter_layout.addWidget(self.WPClabel,16,0)
        self.Color_Enter_layout.addWidget(self.WPCCombobox,17,5,1,4)
        self.Color_Enter_layout.addWidget(self.WPCx,17,0)
        self.Color_Enter_layout.addWidget(self.WPCy,18,0)
        self.Color_Enter_layout.addWidget(self.WPCTable,16,1,3,4)

        # Combobox style
        self.WPCCombobox.setStyleSheet("background-color: #6495ED;selection-background-color: #4169E1;color: black;  /* Set the text color */")
        # 連動
        self.WPCCombobox.currentIndexChanged.connect(self.calculate_WPC)
        # Table style
        self.WPCTable.setStyleSheet("""
                                    QTableWidget::item:selected {
                                color: blcak; /* 設定文字顏色為黑色 */
                                background-color: #008080; /* 設定背景顏色為藍色，你可以根據需要調整 */
                                    }
                                """)

        # # 創建快捷鍵 Ctrl+C
        # copy_shortcut_2 = QShortcut(QKeySequence.Copy, self.WPCTable)
        # copy_shortcut_2.activated.connect(self.copy_WPCtable_content)
        # 为 BM 方法和 newrgb 方法创建两个不同的布局
        self.bm_layout = QGridLayout()  # 用于 BM 方法的布局
        self.rgb_layout = QGridLayout()  # 用于 newrgb 方法的布局

        # 将两个布局添加到主布局中

        self.Color_Enter_layout.addLayout(self.rgb_layout, 16, 5, 1, 4)  # newrgb 方法图表放在第 16 行
        self.Color_Enter_layout.addLayout(self.bm_layout, 16, 5, 1, 4)  # BM 方法图表放在第 16 行
        # 初始化用于保存图表引用的列表
        self.bm_chart_canvases = [None, None, None]
        self.rgb_chart_canvases = [None, None, None]

        # 初始化时绘制 BM 方法和 newrgb 方法的基础长方图
        base_value = 1.0  # 假设基础值为 1.0，相当于 100%
        colors = ['red', 'green', 'blue']
        for i, color in enumerate(colors):
            self.draw_single_bar_chart(base_value, color, i)
            self.draw_single_bar_chart_rgb(base_value, color, i)

        # layout放置
        self.setLayout(self.Color_Enter_layout)

        # Set Background
        self.setStyleSheet("background-color: lightyellow;")
        # # 連接全局信號到相應的方法
        global_signal_manager.databaseUpdated.connect(self.updateLightSourceComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_light_source_datatable)


        global_signal_manager.databaseUpdated.connect(self.update_source_led_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateSourceledComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_light_source_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateLightSourceComboBox)
        # # layer
        global_signal_manager.databaseUpdated.connect(self.update_layer1_datatable)
        global_signal_manager.databaseUpdated.connect(self.updatelayer1ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_layer2_datatable)
        global_signal_manager.databaseUpdated.connect(self.updatelayer2ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_layer3_datatable)
        global_signal_manager.databaseUpdated.connect(self.updatelayer3ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_layer4_datatable)
        global_signal_manager.databaseUpdated.connect(self.updatelayer4ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_layer5_datatable)
        global_signal_manager.databaseUpdated.connect(self.updatelayer5ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_layer6_datatable)
        global_signal_manager.databaseUpdated.connect(self.updatelayer6ComboBox)
        # Fix
        global_signal_manager.databaseUpdated.connect(self.update_RCF_Fix_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateRCF_Fix_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_GCF_Fix_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateGCF_Fix_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_BCF_Fix_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateBCF_Fix_ComboBox)
        # AK
        global_signal_manager.databaseUpdated.connect(self.update_RCF_Change_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateRCF_Change_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_GCF_Change_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateGCF_Change_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_BCF_Change_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateBCF_Change_ComboBox)
        # Differ
        global_signal_manager.databaseUpdated.connect(self.update_RCF_Differ_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateRCF_Differ_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_GCF_Differ_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateGCF_Differ_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_BCF_Differ_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateBCF_Differ_ComboBox)


    def test(self):
        print("SSS")

    # 複製整個表格的內容
    def copy_entire_table_content(self):
        row_count = self.color_table.rowCount()
        column_count = self.color_table.columnCount()

        copied_data = []
        for row in range(row_count):
            row_data = []
            for col in range(column_count):
                item = self.color_table.item(row, col)
                if item:
                    row_data.append(item.text())
                else:
                    row_data.append("")  # 如果單元格為空，填充空字符串
            copied_data.append('\t'.join(row_data))

        # 將複製的數據組合成字符串，每行以換行符分隔
        copied_text = '\n'.join(copied_data)

        # 複製到剪貼板
        clipboard = QApplication.clipboard()
        clipboard.setText(copied_text)

    # def keyPressEvent(self, event):
    #     super().keyPressEvent(event)
    #
    #     if event.key() == Qt.Key.Key_C and (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
    #         copied_cells = sorted(self.color_table.selectedIndexes())
    #
    #         copy_text = ''
    #         max_column = copied_cells[-1].column()
    #         for c in copied_cells:
    #             copy_text += self.color_table.item(c.row(), c.column()).text()
    #             if c.column() == max_column:
    #                 copy_text += '\n'
    #             else:
    #                 copy_text += '\t'
    #
    #         QApplication.clipboard().setText(copy_text)
    #
    #     if event.key() == Qt.Key_V and (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
    #         selection = self.color_table.selectedIndexes()
    #         if selection:
    #             row_anchor = selection[0].row()
    #             column_anchor = selection[0].column()
    #
    #             clipboard = QApplication.clipboard()
    #             rows = clipboard.text().split('\n')
    #             for index_row, row in enumerate(rows):
    #                 values = row.split('\t')
    #                 for index_col, value in enumerate(values):
    #                     item = QTableWidgetItem(value)
    #                     self.color_table.setItem(row_anchor + index_row, column_anchor + index_col, item)
    #         super().keyPressEvent()

    def deleteSelectedCells(self):
        # 獲取選擇的儲存格範圍
        selected_ranges = self.color_table.selectedRanges()

        # 刪除選擇的儲存格內容
        for selected_range in selected_ranges:
            for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                for column in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                    item = self.color_table.item(row, column)
                    if item is not None:
                        item.setText("")  # 可以根據需要進行其他操作

    def calculate_BLU(self):
        if self.check == "NO":
            print("Please select light source-BLU")
            return

        # BLU
        # 選項關鍵----------------------------------------------
        if self.light_source_mode.currentText() == "未選":
            for row in range(1, self.color_table.rowCount()):
                for col in range(1,self.color_table.columnCount()):
                    item = self.color_table.item(row, col)
                    if item is not None:
                        item.setText("")
            # 防呆反灰設定
            self.light_source.setEnabled(False)
            self.light_source_datatable.setEnabled(False)
            self.light_source.setStyleSheet(QCOMBOBOXDISABLE)
            self.light_source_datatable.setStyleSheet(QCOMBOBOXDISABLE)
            # 防呆反灰設定
            self.light_source_led.setEnabled(False)
            self.light_source_led_datatable.setEnabled(False)
            self.light_source_led.setStyleSheet(QCOMBOBOXDISABLE)
            self.light_source_led_datatable.setStyleSheet(QCOMBOBOXDISABLE)
            # 防呆反灰設定
            self.source_led.setEnabled(False)
            self.source_led_datatable.setEnabled(False)
            self.source_led.setStyleSheet(QCOMBOBOXDISABLE)
            self.source_led_datatable.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer1_box.setEnabled(False)
            self.layer1_table.setEnabled(False)
            self.layer1_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer1_table.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer2_box.setEnabled(False)
            self.layer2_table.setEnabled(False)
            self.layer2_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer2_table.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer3_box.setEnabled(False)
            self.layer3_table.setEnabled(False)
            self.layer3_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer3_table.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer4_box.setEnabled(False)
            self.layer4_table.setEnabled(False)
            self.layer4_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer4_table.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer5_box.setEnabled(False)
            self.layer5_table.setEnabled(False)
            self.layer5_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer5_table.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer6_box.setEnabled(False)
            self.layer6_table.setEnabled(False)
            self.layer6_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer6_table.setStyleSheet(QCOMBOBOXDISABLE)
            # 防呆反灰設定
            self.R_fix_box.setEnabled(False)
            self.R_fix_table.setEnabled(False)
            self.R_fix_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.R_fix_table.setStyleSheet(QCOMBOBOXDISABLE)
            # 防呆反灰設定
            self.G_fix_box.setEnabled(False)
            self.G_fix_table.setEnabled(False)
            self.G_fix_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.G_fix_table.setStyleSheet(QCOMBOBOXDISABLE)
            # 防呆反灰設定
            self.B_fix_box.setEnabled(False)
            self.B_fix_table.setEnabled(False)
            self.B_fix_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.B_fix_table.setStyleSheet(QCOMBOBOXDISABLE)
            # 防呆反灰設定
            self.R_aK_box.setEnabled(False)
            self.R_aK_table.setEnabled(False)
            self.R_aK_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.R_aK_table.setStyleSheet(QCOMBOBOXDISABLE)
            # 防呆反灰設定
            self.G_aK_box.setEnabled(False)
            self.G_aK_table.setEnabled(False)
            self.G_aK_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.G_aK_table.setStyleSheet(QCOMBOBOXDISABLE)
            # 防呆反灰設定
            self.B_aK_box.setEnabled(False)
            self.B_aK_table.setEnabled(False)
            self.B_aK_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.B_aK_table.setStyleSheet(QCOMBOBOXDISABLE)
            # 防呆反灰設定
            self.G_differ_box.setEnabled(False)
            self.G_differ_table.setEnabled(False)
            self.G_differ_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.G_differ_table.setStyleSheet(QCOMBOBOXDISABLE)
            # 防呆反灰設定
            self.R_differ_box.setEnabled(False)
            self.R_differ_table.setEnabled(False)
            self.R_differ_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.R_differ_table.setStyleSheet(QCOMBOBOXDISABLE)
            # 防呆反灰設定
            self.B_differ_box.setEnabled(False)
            self.B_differ_table.setEnabled(False)
            self.B_differ_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.B_differ_table.setStyleSheet(QCOMBOBOXDISABLE)
            # 在Color_Table上顯示
            self.color_table.setItem(4, 7, QTableWidgetItem(""))
            self.color_table.setItem(4, 8, QTableWidgetItem(""))
            self.color_table.setItem(4, 9, QTableWidgetItem(""))
        if self.light_source_mode.currentText() == "自訂":
            # 防呆區
            self.light_source.setEnabled(True)
            self.light_source_datatable.setEnabled(True)
            self.light_source.setStyleSheet(QCOMBOXSETTING)
            self.light_source_datatable.setStyleSheet(QCOMBOBOXTABLESELECT)
            self.light_source_led.setEnabled(False)
            self.light_source_led_datatable.setEnabled(False)
            self.source_led.setEnabled(False)
            self.source_led_datatable.setEnabled(False)
            self.source_led.setStyleSheet(QCOMBOBOXDISABLE)
            self.source_led_datatable.setStyleSheet(QCOMBOBOXDISABLE)
            self.light_source_led.setStyleSheet(QCOMBOBOXDISABLE)
            self.light_source_led_datatable.setStyleSheet(QCOMBOBOXDISABLE)
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
            print("header_BLU", header_BLU)
            column_index_BLU = header_BLU.index(f"{column_name_BLU}")

            # 取得下一個欄位的名稱
            # next_column_name_BLU = header_BLU[column_index_BLU + 1]
            # print("next_column_name_BLU",next_column_name_BLU)

            # 取得指定欄位的數據並轉換為 Series
            BLU_spectrum_Series = pd.Series([row[column_index_BLU] for row in result_BLU])
            print("column_index_BLU", column_index_BLU)
            # BLU_spectrum_Series_test =pd.Series([row[column_index_BLU+1] for row in result_BLU])
            # print("BLU_spectrum_Series_test", BLU_spectrum_Series_test)

            # 將 Series 中的字符串轉換為數值
            BLU_spectrum_Series = pd.to_numeric(BLU_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            BLU_spectrum_Series = BLU_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            BLU_spectrum_Series = BLU_spectrum_Series.dropna()

            # 關閉連線
            connection_BLU.close()
            # 在Color_Table上顯示
            self.color_table.setItem(4, 7, QTableWidgetItem(self.light_source.currentText()))
            self.color_table.setItem(4, 8, QTableWidgetItem(""))
            self.color_table.setItem(4, 9, QTableWidgetItem(""))
            print("BLU_spectrum_Series",BLU_spectrum_Series)
            # 自訂BLU_Spectrum回傳
            return BLU_spectrum_Series
            # 選項關鍵----------------------------------------------------------
        elif self.light_source_mode.currentText() == "替換":
            self.light_source.setEnabled(True)
            self.light_source_datatable.setEnabled(True)
            self.light_source_led.setEnabled(True)
            self.light_source_led_datatable.setEnabled(True)
            self.source_led.setEnabled(True)
            self.source_led_datatable.setEnabled(True)
            self.light_source.setStyleSheet(QCOMBOXSETTING)
            self.light_source_datatable.setStyleSheet(QCOMBOBOXTABLESELECT)
            self.source_led.setStyleSheet(QCOMBOXSETTING)
            self.source_led_datatable.setStyleSheet(QCOMBOBOXTABLESELECT)
            self.light_source_led.setStyleSheet(QCOMBOXSETTING)
            self.light_source_led_datatable.setStyleSheet(QCOMBOBOXTABLESELECT)
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
            LED_light_Source_spectrum_Series = pd.Series(
                [row[column_index_LED_light_Source] for row in result_led_light_source])
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
            # 在Color_Table上顯示
            self.color_table.setItem(4, 7, QTableWidgetItem(self.light_source.currentText()))
            self.color_table.setItem(4, 8, QTableWidgetItem(self.light_source_led.currentText()))
            self.color_table.setItem(4, 9, QTableWidgetItem(self.source_led.currentText()))
            return BLU_spectrum_Series
        else:
            print("BLU_None")
            return None

    # Cell
    def calculate_layer1(self):
        if self.check == "NO":
            print("Please select light source")
            return
        if self.layer1_mode.currentText() == "自訂":
            self.layer1_box.setEnabled(True)
            self.layer1_table.setEnabled(True)
            self.layer1_box.setStyleSheet(QCOMBOXSETTING)
            self.layer1_table.setStyleSheet(QCOMBOBOXTABLESELECT)
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
            # print("layer1_spectrum_Series", layer1_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer1_spectrum_Series = pd.to_numeric(layer1_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer1_spectrum_Series = layer1_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer1_spectrum_Series = layer1_spectrum_Series.dropna()

            # 關閉連線
            connection_layer1.close()
            # Table顯示
            self.color_table.setItem(4, 10, QTableWidgetItem(self.layer1_box.currentText()))
            # 自訂BLU_Spectrum回傳
            return layer1_spectrum_Series
        else:
            self.layer1_box.setEnabled(False)
            self.layer1_table.setEnabled(False)
            self.layer1_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer1_table.setStyleSheet(QCOMBOBOXDISABLE)
            layer1_spectrum_Series = 1
            print("layer1:未選")
            # Table顯示
            self.color_table.setItem(4, 10, QTableWidgetItem(""))
            return layer1_spectrum_Series

    # BSITO
    def calculate_layer2(self):
        if self.check == "NO":
            print("Please select light source")
        if self.layer2_mode.currentText() == "自訂":
            self.layer2_box.setEnabled(True)
            self.layer2_table.setEnabled(True)
            self.layer2_box.setStyleSheet(QCOMBOXSETTING)
            self.layer2_table.setStyleSheet(QCOMBOBOXTABLESELECT)
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
            # print("layer2_spectrum_Series", layer2_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer2_spectrum_Series = pd.to_numeric(layer2_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer2_spectrum_Series = layer2_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer2_spectrum_Series = layer2_spectrum_Series.dropna()

            # 關閉連線
            connection_layer2.close()
            # Table顯示
            self.color_table.setItem(4, 11, QTableWidgetItem(self.layer2_box.currentText()))
            # 自訂BLU_Spectrum回傳
            return layer2_spectrum_Series
        else:
            self.layer2_box.setEnabled(False)
            self.layer2_table.setEnabled(False)
            self.layer2_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer2_table.setStyleSheet(QCOMBOBOXDISABLE)
            layer2_spectrum_Series = 1
            print("layer2:未選")
            # Table顯示
            self.color_table.setItem(4, 11, QTableWidgetItem(""))
            return layer2_spectrum_Series

    def calculate_layer3(self):
        if self.check == "NO":
            print("Please select light source")
        if self.layer3_mode.currentText() == "自訂":
            self.layer3_box.setEnabled(True)
            self.layer3_table.setEnabled(True)
            self.layer3_box.setStyleSheet(QCOMBOXSETTING)
            self.layer3_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            print("in_layer3_自訂")
            connection_layer3 = sqlite3.connect("Layer3_spectrum.db")
            cursor_layer3 = connection_layer3.cursor()
            # 取得BLU資料
            column_name_layer3 = self.layer3_box.currentText()
            table_name_layer3 = self.layer3_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer3 = f"SELECT * FROM '{table_name_layer3}';"
            cursor_layer3.execute(query_layer3)
            result_layer3 = cursor_layer3.fetchall()

            # 找到指定標題的欄位索引
            header_layer3 = [column[0] for column in cursor_layer3.description]
            column_index_layer3 = header_layer3.index(f"{column_name_layer3}")

            # 取得指定欄位的數據並轉換為 Series
            layer3_spectrum_Series = pd.Series([row[column_index_layer3] for row in result_layer3])
            # print("layer2_spectrum_Series", layer2_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer3_spectrum_Series = pd.to_numeric(layer3_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer3_spectrum_Series = layer3_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer3_spectrum_Series = layer3_spectrum_Series.dropna()

            # 關閉連線
            connection_layer3.close()
            # Table顯示
            self.color_table.setItem(4, 12, QTableWidgetItem(self.layer3_box.currentText()))
            # 自訂BLU_Spectrum回傳
            return layer3_spectrum_Series
        else:
            self.layer3_box.setEnabled(False)
            self.layer3_table.setEnabled(False)
            self.layer3_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer3_table.setStyleSheet(QCOMBOBOXDISABLE)
            layer3_spectrum_Series = 1
            print("layer3:未選")
            # Table顯示
            self.color_table.setItem(4, 12, QTableWidgetItem(""))
            return layer3_spectrum_Series

    def calculate_layer4(self):
        if self.check == "NO":
            print("Please select light source")
        if self.layer4_mode.currentText() == "自訂":
            self.layer4_box.setEnabled(True)
            self.layer4_table.setEnabled(True)
            self.layer4_box.setStyleSheet(QCOMBOXSETTING)
            self.layer4_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            print("in_layer4_自訂")
            connection_layer4 = sqlite3.connect("Layer4_spectrum.db")
            cursor_layer4 = connection_layer4.cursor()
            # 取得BLU資料
            column_name_layer4 = self.layer4_box.currentText()
            table_name_layer4 = self.layer4_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer4 = f"SELECT * FROM '{table_name_layer4}';"
            cursor_layer4.execute(query_layer4)
            result_layer4 = cursor_layer4.fetchall()

            # 找到指定標題的欄位索引
            header_layer4 = [column[0] for column in cursor_layer4.description]
            column_index_layer4 = header_layer4.index(f"{column_name_layer4}")

            # 取得指定欄位的數據並轉換為 Series
            layer4_spectrum_Series = pd.Series([row[column_index_layer4] for row in result_layer4])
            # print("layer4_spectrum_Series", layer4_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer4_spectrum_Series = pd.to_numeric(layer4_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer4_spectrum_Series = layer4_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer4_spectrum_Series = layer4_spectrum_Series.dropna()

            # 關閉連線
            connection_layer4.close()
            # Table顯示
            self.color_table.setItem(4, 13, QTableWidgetItem(self.layer4_box.currentText()))
            # 自訂BLU_Spectrum回傳
            return layer4_spectrum_Series
        else:
            self.layer4_box.setEnabled(False)
            self.layer4_table.setEnabled(False)
            self.layer4_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer4_table.setStyleSheet(QCOMBOBOXDISABLE)
            layer4_spectrum_Series = 1
            print("layer4:未選")
            # Table顯示
            self.color_table.setItem(4, 13, QTableWidgetItem(""))
            return layer4_spectrum_Series

    def calculate_layer5(self):
        if self.check == "NO":
            print("Please select light source")
        if self.layer5_mode.currentText() == "自訂":
            self.layer5_box.setEnabled(True)
            self.layer5_table.setEnabled(True)
            self.layer5_box.setStyleSheet(QCOMBOXSETTING)
            self.layer5_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            print("in_layer5_自訂")
            connection_layer5 = sqlite3.connect("Layer5_spectrum.db")
            cursor_layer5 = connection_layer5.cursor()
            # 取得BLU資料
            column_name_layer5 = self.layer5_box.currentText()
            table_name_layer5 = self.layer5_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer5 = f"SELECT * FROM '{table_name_layer5}';"
            cursor_layer5.execute(query_layer5)
            result_layer5 = cursor_layer5.fetchall()

            # 找到指定標題的欄位索引
            header_layer5 = [column[0] for column in cursor_layer5.description]
            column_index_layer5 = header_layer5.index(f"{column_name_layer5}")

            # 取得指定欄位的數據並轉換為 Series
            layer5_spectrum_Series = pd.Series([row[column_index_layer5] for row in result_layer5])
            # print("layer5_spectrum_Series", layer5_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer5_spectrum_Series = pd.to_numeric(layer5_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer5_spectrum_Series = layer5_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer5_spectrum_Series = layer5_spectrum_Series.dropna()

            # 關閉連線
            connection_layer5.close()
            # Table顯示
            self.color_table.setItem(4, 14, QTableWidgetItem(self.layer5_box.currentText()))
            # 自訂BLU_Spectrum回傳
            return layer5_spectrum_Series
        else:
            self.layer5_box.setEnabled(False)
            self.layer5_table.setEnabled(False)
            self.layer5_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer5_table.setStyleSheet(QCOMBOBOXDISABLE)
            layer5_spectrum_Series = 1
            print("layer5:未選")
            # Table顯示
            self.color_table.setItem(4, 14, QTableWidgetItem(""))
            return layer5_spectrum_Series

    def calculate_layer6(self):
        if self.check == "NO":
            print("Please select light source")
        if self.layer6_mode.currentText() == "自訂":
            self.layer6_box.setEnabled(True)
            self.layer6_table.setEnabled(True)
            self.layer6_box.setStyleSheet(QCOMBOXSETTING)
            self.layer6_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            print("in_layer6_自訂")
            connection_layer6 = sqlite3.connect("Layer6_spectrum.db")
            cursor_layer6 = connection_layer6.cursor()
            # 取得BLU資料
            column_name_layer6 = self.layer6_box.currentText()
            table_name_layer6 = self.layer6_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer6 = f"SELECT * FROM '{table_name_layer6}';"
            cursor_layer6.execute(query_layer6)
            result_layer6 = cursor_layer6.fetchall()

            # 找到指定標題的欄位索引
            header_layer6 = [column[0] for column in cursor_layer6.description]
            column_index_layer6 = header_layer6.index(f"{column_name_layer6}")

            # 取得指定欄位的數據並轉換為 Series
            layer6_spectrum_Series = pd.Series([row[column_index_layer6] for row in result_layer6])
            # print("layer6_spectrum_Series", layer6_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer6_spectrum_Series = pd.to_numeric(layer6_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer6_spectrum_Series = layer6_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer6_spectrum_Series = layer6_spectrum_Series.dropna()

            # 關閉連線
            connection_layer6.close()
            # Table顯示
            self.color_table.setItem(4, 15, QTableWidgetItem(self.layer6_box.currentText()))
            # 自訂BLU_Spectrum回傳
            return layer6_spectrum_Series
        else:
            self.layer6_box.setEnabled(False)
            self.layer6_table.setEnabled(False)
            self.layer6_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.layer6_table.setStyleSheet(QCOMBOBOXDISABLE)
            layer6_spectrum_Series = 1
            # Table顯示
            self.color_table.setItem(4, 15, QTableWidgetItem(""))
            print("layer6:未選")
            return layer6_spectrum_Series

    # RCF_Fix
    def calculate_RCF_Fix(self):
        if self.check == "NO":
            print("Please select light source")
        if self.R_fix_mode.currentText() == "自訂":
            # 防呆反灰設定
            self.R_fix_box.setEnabled(True)
            self.R_fix_table.setEnabled(True)
            self.R_fix_box.setStyleSheet(QCOMBOXSETTING)
            self.R_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
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
            # print("RCF_Fix_spectrum_Series", RCF_Fix_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            RCF_Fix_spectrum_Series = pd.to_numeric(RCF_Fix_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            RCF_Fix_spectrum_Series = RCF_Fix_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            RCF_Fix_spectrum_Series = RCF_Fix_spectrum_Series.dropna()

            # 第二欄 新的 Series 除第一列之外的所有數值
            new_RCF_Fix_spectrum_Series = RCF_Fix_spectrum_Series[1:]

            # 厚度
            self.RTK = RCF_Fix_spectrum_Series.iloc[0]
            self.R_TK_label.setText(f"{self.RTK:.3f} um")

            # 在計算完 C_Series 後，加上以下代碼
            new_RCF_Fix_spectrum_Series = new_RCF_Fix_spectrum_Series.reset_index(drop=True)

            # 關閉連線
            connection_RCF_Fix.close()
            # 在Color_Table上顯示
            self.color_table.setItem(4,1,QTableWidgetItem(self.R_fix_box.currentText()))
            self.color_table.setItem(4,2,QTableWidgetItem(f"{self.RTK:.3f}"))
            # 自訂RCF_Fix_Spectrum回傳
            return new_RCF_Fix_spectrum_Series
        elif self.R_fix_mode.currentText() == "未選":
            # 防呆反灰設定
            self.R_fix_box.setEnabled(False)
            self.R_fix_table.setEnabled(False)
            self.R_fix_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.R_fix_table.setStyleSheet(QCOMBOBOXDISABLE)
            new_RCF_Fix_spectrum_Series = 1
            print("RCF_Fix:未選")
            return new_RCF_Fix_spectrum_Series
        elif self.R_fix_mode.currentText() == "模擬":
            self.R_fix_box.setEnabled(True)
            self.R_fix_table.setEnabled(True)
            self.R_fix_box.setStyleSheet(QCOMBOXSETTING)
            self.R_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            new_RCF_Fix_spectrum_Series = 1
            print("RCF_Fix:模擬,默認1")
            return new_RCF_Fix_spectrum_Series

    # GCF_Fix
    def calculate_GCF_Fix(self):
        if self.check == "NO":
            print("Please select light source")
        if self.G_fix_mode.currentText() == "自訂":
            # 防呆反灰設定
            self.G_fix_box.setEnabled(True)
            self.G_fix_table.setEnabled(True)
            self.G_fix_box.setStyleSheet(QCOMBOXSETTING)
            self.G_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
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
            # print("GCF_Fix_spectrum_Series", GCF_Fix_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            GCF_Fix_spectrum_Series = pd.to_numeric(GCF_Fix_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            GCF_Fix_spectrum_Series = GCF_Fix_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            GCF_Fix_spectrum_Series = GCF_Fix_spectrum_Series.dropna()

            # 第二欄 新的 Series 除第一列之外的所有數值
            new_GCF_Fix_spectrum_Series = GCF_Fix_spectrum_Series[1:]

            # 厚度
            self.GTK = GCF_Fix_spectrum_Series.iloc[0]
            self.G_TK_label.setText(f"{self.GTK:.3f} um")

            # 在計算完 C_Series 後，加上以下代碼
            new_GCF_Fix_spectrum_Series = new_GCF_Fix_spectrum_Series.reset_index(drop=True)

            # 關閉連線
            connection_GCF_Fix.close()
            # 在Color_Table上顯示
            self.color_table.setItem(4, 3, QTableWidgetItem(self.G_fix_box.currentText()))
            self.color_table.setItem(4, 4, QTableWidgetItem(f"{self.GTK:.3f}"))
            # 自訂GCF_Fix_Spectrum回傳
            return new_GCF_Fix_spectrum_Series
        elif self.G_fix_mode.currentText() == "未選":
            # 防呆反灰設定
            self.G_fix_box.setEnabled(False)
            self.G_fix_table.setEnabled(False)
            self.G_fix_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.G_fix_table.setStyleSheet(QCOMBOBOXDISABLE)
            new_GCF_Fix_spectrum_Series = 1
            print("GCF_Fix:未選")
            return new_GCF_Fix_spectrum_Series
        elif self.G_fix_mode.currentText() == "模擬":
            self.G_fix_box.setEnabled(True)
            self.G_fix_table.setEnabled(True)
            self.G_fix_box.setStyleSheet(QCOMBOXSETTING)
            self.G_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            new_GCF_Fix_spectrum_Series = 1
            print("GCF_Fix:模擬,默認1")
            return new_GCF_Fix_spectrum_Series

    # BCF_Fix
    def calculate_BCF_Fix(self):
        if self.check == "NO":
            print("Please select light source")
        if self.B_fix_mode.currentText() == "自訂":
            # 防呆反灰設定
            self.B_fix_box.setEnabled(True)
            self.B_fix_table.setEnabled(True)
            self.B_fix_box.setStyleSheet(QCOMBOXSETTING)
            self.B_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
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
            # print("BCF_Fix_spectrum_Series", BCF_Fix_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            BCF_Fix_spectrum_Series = pd.to_numeric(BCF_Fix_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            BCF_Fix_spectrum_Series = BCF_Fix_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            BCF_Fix_spectrum_Series = BCF_Fix_spectrum_Series.dropna()

            # 第二欄 新的 Series 除第一列之外的所有數值
            new_BCF_Fix_spectrum_Series = BCF_Fix_spectrum_Series[1:]

            # 厚度
            self.BTK = BCF_Fix_spectrum_Series.iloc[0]
            self.B_TK_label.setText(f"{self.BTK:.3f} um")

            # 在計算完 C_Series 後，加上以下代碼
            new_BCF_Fix_spectrum_Series = new_BCF_Fix_spectrum_Series.reset_index(drop=True)

            # 關閉連線
            connection_BCF_Fix.close()
            # 在Color_Table上顯示
            self.color_table.setItem(4, 5, QTableWidgetItem(self.B_fix_box.currentText()))
            self.color_table.setItem(4, 6, QTableWidgetItem(f"{self.BTK:.3f}"))
            # 自訂BCF_Fix_Spectrum回傳
            return new_BCF_Fix_spectrum_Series
        elif self.B_fix_mode.currentText() == "未選":
            # 防呆反灰設定
            self.B_fix_box.setEnabled(False)
            self.B_fix_table.setEnabled(False)
            self.B_fix_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.B_fix_table.setStyleSheet(QCOMBOBOXDISABLE)
            new_BCF_Fix_spectrum_Series = 1
            print("BCF_Fix:未選")
            return new_BCF_Fix_spectrum_Series
        elif self.B_fix_mode.currentText() == "模擬":
            self.B_fix_box.setEnabled(True)
            self.B_fix_table.setEnabled(True)
            self.B_fix_box.setStyleSheet(QCOMBOXSETTING)
            self.B_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            new_BCF_Fix_spectrum_Series = 1
            print("BCF_Fix:模擬,默認1")
            return new_BCF_Fix_spectrum_Series

    # RCF_Change
    def calculate_RCF_Change(self):
        if self.check == "NO":
            print("Please select light source")
        if self.R_aK_mode.currentText() == "自訂":
            # 防呆反灰設定
            self.R_aK_box.setEnabled(True)
            self.R_aK_table.setEnabled(True)
            self.R_aK_box.setStyleSheet(QCOMBOXSETTING)
            self.R_aK_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            print("in_RCF_Change自訂")
            connection_RCF_Change = sqlite3.connect("RCF_Change_spectrum.db")
            cursor_RCF_Change = connection_RCF_Change.cursor()
            # 取得Rdiffer資料
            column_name_RCF_Change = self.R_aK_box.currentText()
            # print("column_name_RCF_Change",column_name_RCF_Change)
            table_name_RCF_Change = self.R_aK_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_RCF_Change = f"SELECT * FROM '{table_name_RCF_Change}';"
            cursor_RCF_Change.execute(query_RCF_Change)
            result_RCF_Change = cursor_RCF_Change.fetchall()

            # 找到指定標題的欄位索引
            header_RCF_Change = [column[0] for column in cursor_RCF_Change.description]

            # 移除所有標題中的編號
            header_RCF_Change = [re.sub(r'_\d+$', '', col) for col in header_RCF_Change]
            print("header_RCF_Change", header_RCF_Change)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_RCF_Change) if col == column_name_RCF_Change]
            print("indices_without_number", indices_without_number)

            # 使用 pd.to_numeric 转换为数值类型，并处理无法转换的值
            SP_series_list = [
                pd.to_numeric(pd.Series([row[i] for row in result_RCF_Change]), errors='coerce').fillna(0)
                for i in indices_without_number
            ]
            print("SP_series_list",SP_series_list)

            # TK_SP準備分厚度順序
            TK_SP_list = []
            for i in range(len(indices_without_number)):
                TK_SP_list.append(SP_series_list[i].iloc[0])
            print("TK_SP_list",TK_SP_list)
            Sort_TK_SP_list = sorted(TK_SP_list,reverse=True)
            print("Sort_TK_SP_list",Sort_TK_SP_list)
            self.R_aK_TK_edit_label.setText(f"TKrange: {min(TK_SP_list):.2f}~{max(TK_SP_list):.2f}")

            # 反轉找回原本的順序list
            Re_list = []
            for i in range(len(indices_without_number)):
                Re_list.append(TK_SP_list.index(Sort_TK_SP_list[i]))
            print("Re_list",Re_list)

            A_list = []
            for i in range(len(indices_without_number)-1):
                A_list.append((-1/(float(Sort_TK_SP_list[i+1]) - float(Sort_TK_SP_list[i])) * np.log(SP_series_list[Re_list[i+1]][1:]/SP_series_list[Re_list[i]][1:])).reset_index(drop=True))
            print("A_list",A_list)
            K_list = []
            for i in range(len(indices_without_number)-1):
                K_list.append(SP_series_list[Re_list[i]][1:].reset_index(drop=True)/np.exp((-1 * A_list[i])* Sort_TK_SP_list[i]))
            print("K_list",K_list)
            if self.R_aK_TK_edit.text() == "":
                AK_RCF_Change_spectrum_Series = 1
                print("AK_RCF_Change_spectrum_Series",AK_RCF_Change_spectrum_Series)
                return AK_RCF_Change_spectrum_Series
            elif self.R_aK_TK_edit.text() != "":

                R_aK_TK = float(self.R_aK_TK_edit.text())
                # 先将 R_aK_TK 添加到 Sort_TK_SP_list
                Sort_TK_SP_list.append(R_aK_TK)
                # 然后对更新后的列表进行排序
                Sort_TK_SP_AK_list = sorted(Sort_TK_SP_list, reverse=True)
                print("Sort_TK_SP_AK_list", Sort_TK_SP_AK_list)

                # 4種情況
                if R_aK_TK == max(Sort_TK_SP_AK_list):
                    print("max")
                    AK_RCF_Change_spectrum_Series = K_list[0] * np.exp(-1 * A_list[0] * R_aK_TK)
                    print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 1, QTableWidgetItem(self.R_aK_box.currentText()))
                    self.color_table.setItem(4, 2, QTableWidgetItem(f"{R_aK_TK}"))
                    return AK_RCF_Change_spectrum_Series
                elif R_aK_TK == min(Sort_TK_SP_AK_list):
                    print("min")
                    AK_RCF_Change_spectrum_Series = K_list[len(K_list)-1] * np.exp(-1 * A_list[len(A_list)-1] * R_aK_TK)
                    print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 1, QTableWidgetItem(self.R_aK_box.currentText()))
                    self.color_table.setItem(4, 2, QTableWidgetItem(f"{R_aK_TK}"))
                    return AK_RCF_Change_spectrum_Series
                elif R_aK_TK in TK_SP_list:
                    print("equal")
                    AK_RCF_Change_spectrum_Series = SP_series_list[TK_SP_list.index(R_aK_TK)][1:].reset_index(drop=True)
                    print("AK_RCF_Change_spectrum_Series",AK_RCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 1, QTableWidgetItem(self.R_aK_box.currentText()))
                    self.color_table.setItem(4, 2, QTableWidgetItem(f"{R_aK_TK}"))
                    return AK_RCF_Change_spectrum_Series
                else:
                    print("mid")
                    position = Sort_TK_SP_AK_list.index(R_aK_TK)
                    AK_RCF_Change_spectrum_Series = K_list[position-1] * np.exp(-1 * A_list[position-1] * R_aK_TK)
                    print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 1, QTableWidgetItem(self.R_aK_box.currentText()))
                    self.color_table.setItem(4, 2, QTableWidgetItem(f"{R_aK_TK}"))
                    return AK_RCF_Change_spectrum_Series

            # 關閉連線
            connection_RCF_Change.close()
        elif self.R_aK_mode.currentText() == "未選":
            # 防呆反灰設定
            self.R_aK_box.setEnabled(False)
            self.R_aK_table.setEnabled(False)
            self.R_aK_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.R_aK_table.setStyleSheet(QCOMBOBOXDISABLE)
            AK_RCF_Change_spectrum_Series = 1
            print("RCF_Change:未選,默認1")
            return AK_RCF_Change_spectrum_Series
        elif self.R_aK_mode.currentText() == "模擬":
            self.R_aK_box.setEnabled(True)
            self.R_aK_table.setEnabled(True)
            self.R_aK_box.setStyleSheet(QCOMBOXSETTING)
            self.R_aK_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            AK_RCF_Change_spectrum_Series = 1
            print("RCF_Change:模擬,默認1")
            return AK_RCF_Change_spectrum_Series

    # GCF_Change
    def calculate_GCF_Change(self):
        if self.check == "NO":
            print("Please select light source")
        if self.G_aK_mode.currentText() == "自訂":
            # 防呆反灰設定
            self.G_aK_box.setEnabled(True)
            self.G_aK_table.setEnabled(True)
            self.G_aK_box.setStyleSheet(QCOMBOXSETTING)
            self.G_aK_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            print("in_GCF_Change自訂")
            connection_GCF_Change = sqlite3.connect("GCF_Change_spectrum.db")
            cursor_GCF_Change = connection_GCF_Change.cursor()
            # 取得Gchange資料
            column_name_GCF_Change = self.G_aK_box.currentText()
            # print("column_name_GCF_Change",column_name_GCF_Change)
            table_name_GCF_Change = self.G_aK_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_GCF_Change = f"SELECT * FROM '{table_name_GCF_Change}';"
            cursor_GCF_Change.execute(query_GCF_Change)
            result_GCF_Change = cursor_GCF_Change.fetchall()

            # 找到指定標題的欄位索引
            header_GCF_Change = [column[0] for column in cursor_GCF_Change.description]

            # 移除所有標題中的編號
            header_GCF_Change = [re.sub(r'_\d+$', '', col) for col in header_GCF_Change]
            print("header_GCF_Change", header_GCF_Change)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_GCF_Change) if col == column_name_GCF_Change]
            print("indices_without_number", indices_without_number)

            # 使用 pd.to_numeric 转换为数值类型，并处理无法转换的值
            SP_series_list = [
                pd.to_numeric(pd.Series([row[i] for row in result_GCF_Change]), errors='coerce').fillna(0)
                for i in indices_without_number
            ]
            print("SP_series_list", SP_series_list)

            # TK_SP準備分厚度順序
            TK_SP_list = []
            for i in range(len(indices_without_number)):
                TK_SP_list.append(SP_series_list[i].iloc[0])
            print("TK_SP_list", TK_SP_list)
            Sort_TK_SP_list = sorted(TK_SP_list, reverse=True)
            print("Sort_TK_SP_list", Sort_TK_SP_list)
            self.G_aK_TK_edit_label.setText(f"TKrange: {min(TK_SP_list):.2f}~{max(TK_SP_list):.2f}")

            # 反轉找回原本的順序list
            Re_list = []
            for i in range(len(indices_without_number)):
                Re_list.append(TK_SP_list.index(Sort_TK_SP_list[i]))
            print("Re_list", Re_list)

            A_list = []
            for i in range(len(indices_without_number) - 1):
                A_list.append((-1 / (float(Sort_TK_SP_list[i + 1]) - float(Sort_TK_SP_list[i])) * np.log(
                    SP_series_list[Re_list[i + 1]][1:] / SP_series_list[Re_list[i]][1:])).reset_index(drop=True))
            print("A_list", A_list)
            K_list = []
            for i in range(len(indices_without_number) - 1):
                K_list.append(SP_series_list[Re_list[i]][1:].reset_index(drop=True) / np.exp(
                    (-1 * A_list[i]) * Sort_TK_SP_list[i]))
            print("K_list", K_list)
            if self.G_aK_TK_edit.text() == "":
                AK_GCF_Change_spectrum_Series = 1
                print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                return AK_GCF_Change_spectrum_Series
            elif self.G_aK_TK_edit.text() != "":

                G_aK_TK = float(self.G_aK_TK_edit.text())
                # 先将 G_aK_TK 添加到 Sort_TK_SP_list
                Sort_TK_SP_list.append(G_aK_TK)
                # 然后对更新后的列表进行排序
                Sort_TK_SP_AK_list = sorted(Sort_TK_SP_list, reverse=True)
                print("Sort_TK_SP_AK_list", Sort_TK_SP_AK_list)

                # 4種情況
                if G_aK_TK == max(Sort_TK_SP_AK_list):
                    print("max")
                    AK_GCF_Change_spectrum_Series = K_list[0] * np.exp(-1 * A_list[0] * G_aK_TK)
                    print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 3, QTableWidgetItem(self.G_aK_box.currentText()))
                    self.color_table.setItem(4, 4, QTableWidgetItem(f"{G_aK_TK}"))
                    return AK_GCF_Change_spectrum_Series
                elif G_aK_TK == min(Sort_TK_SP_AK_list):
                    print("min")
                    AK_GCF_Change_spectrum_Series = K_list[len(K_list) - 1] * np.exp(
                        -1 * A_list[len(A_list) - 1] * G_aK_TK)
                    print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 3, QTableWidgetItem(self.G_aK_box.currentText()))
                    self.color_table.setItem(4, 4, QTableWidgetItem(f"{G_aK_TK}"))
                    return AK_GCF_Change_spectrum_Series
                elif G_aK_TK in TK_SP_list:
                    print("equal")
                    AK_GCF_Change_spectrum_Series = SP_series_list[TK_SP_list.index(G_aK_TK)][1:].reset_index(drop=True)
                    print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 3, QTableWidgetItem(self.G_aK_box.currentText()))
                    self.color_table.setItem(4, 4, QTableWidgetItem(f"{G_aK_TK}"))
                    return AK_GCF_Change_spectrum_Series
                else:
                    print("mid")
                    position = Sort_TK_SP_AK_list.index(G_aK_TK)
                    AK_GCF_Change_spectrum_Series = K_list[position - 1] * np.exp(-1 * A_list[position - 1] * G_aK_TK)
                    print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 3, QTableWidgetItem(self.G_aK_box.currentText()))
                    self.color_table.setItem(4, 4, QTableWidgetItem(f"{G_aK_TK}"))
                    return AK_GCF_Change_spectrum_Series
            # 關閉連線
            connection_GCF_Change.close()
        elif self.G_aK_mode.currentText() == "未選":
            # 防呆反灰設定
            self.G_aK_box.setEnabled(False)
            self.G_aK_table.setEnabled(False)
            self.G_aK_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.G_aK_table.setStyleSheet(QCOMBOBOXDISABLE)
            AK_GCF_Change_spectrum_Series = 1
            print("GCF_Change:未選,默認1")
            return AK_GCF_Change_spectrum_Series
        elif self.G_aK_mode.currentText() == "模擬":
            self.G_aK_box.setEnabled(True)
            self.G_aK_table.setEnabled(True)
            self.G_aK_box.setStyleSheet(QCOMBOXSETTING)
            self.G_aK_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            AK_GCF_Change_spectrum_Series = 1
            print("GCF_Change:模擬,默認1")
            return AK_GCF_Change_spectrum_Series

    # BCF_Change
    def calculate_BCF_Change(self):
        if self.check == "NO":
            print("Please select light source")
        if self.B_aK_mode.currentText() == "自訂":
            # 防呆反灰設定
            self.B_aK_box.setEnabled(True)
            self.B_aK_table.setEnabled(True)
            self.B_aK_box.setStyleSheet(QCOMBOXSETTING)
            self.B_aK_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            print("in_BCF_Change自訂")
            connection_BCF_Change = sqlite3.connect("BCF_Change_spectrum.db")
            cursor_BCF_Change = connection_BCF_Change.cursor()
            # 取得Bchange資料
            column_name_BCF_Change = self.B_aK_box.currentText()
            # print("column_name_BCF_Change",column_name_BCF_Change)
            table_name_BCF_Change = self.B_aK_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_BCF_Change = f"SELECT * FROM '{table_name_BCF_Change}';"
            cursor_BCF_Change.execute(query_BCF_Change)
            result_BCF_Change = cursor_BCF_Change.fetchall()

            # 找到指定標題的欄位索引
            header_BCF_Change = [column[0] for column in cursor_BCF_Change.description]

            # 移除所有標題中的編號
            header_BCF_Change = [re.sub(r'_\d+$', '', col) for col in header_BCF_Change]
            print("header_BCF_Change", header_BCF_Change)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_BCF_Change) if col == column_name_BCF_Change]
            print("indices_without_number", indices_without_number)

            # 使用 pd.to_numeric 转换为数值类型，并处理无法转换的值
            SP_series_list = [
                pd.to_numeric(pd.Series([row[i] for row in result_BCF_Change]), errors='coerce').fillna(0)
                for i in indices_without_number
            ]
            print("SP_series_list", SP_series_list)

            # TK_SP準備分厚度順序
            TK_SP_list = []
            for i in range(len(indices_without_number)):
                TK_SP_list.append(SP_series_list[i].iloc[0])
            print("TK_SP_list", TK_SP_list)
            Sort_TK_SP_list = sorted(TK_SP_list, reverse=True)
            print("Sort_TK_SP_list", Sort_TK_SP_list)
            self.B_aK_TK_edit_label.setText(f"TKrange: {min(TK_SP_list):.2f}~{max(TK_SP_list):.2f}")

            # 反轉找回原本的順序list
            Re_list = []
            for i in range(len(indices_without_number)):
                Re_list.append(TK_SP_list.index(Sort_TK_SP_list[i]))
            print("Re_list", Re_list)

            A_list = []
            for i in range(len(indices_without_number) - 1):
                A_list.append((-1 / (float(Sort_TK_SP_list[i + 1]) - float(Sort_TK_SP_list[i])) * np.log(
                    SP_series_list[Re_list[i + 1]][1:] / SP_series_list[Re_list[i]][1:])).reset_index(drop=True))
            print("A_list", A_list)
            K_list = []
            for i in range(len(indices_without_number) - 1):
                K_list.append(SP_series_list[Re_list[i]][1:].reset_index(drop=True) / np.exp(
                    (-1 * A_list[i]) * Sort_TK_SP_list[i]))
            print("K_list", K_list)
            if self.B_aK_TK_edit.text() == "":
                AK_BCF_Change_spectrum_Series = 1
                print("AK_BCF_Change_spectrum_Series", AK_BCF_Change_spectrum_Series)
                return AK_BCF_Change_spectrum_Series
            elif self.B_aK_TK_edit.text() != "":

                B_aK_TK = float(self.B_aK_TK_edit.text())
                # 先将 B_aK_TK 添加到 Sort_TK_SP_list
                Sort_TK_SP_list.append(B_aK_TK)
                # 然后对更新后的列表进行排序
                Sort_TK_SP_AK_list = sorted(Sort_TK_SP_list, reverse=True)
                print("Sort_TK_SP_AK_list", Sort_TK_SP_AK_list)

                # 4種情況
                if B_aK_TK == max(Sort_TK_SP_AK_list):
                    print("max")
                    AK_BCF_Change_spectrum_Series = K_list[0] * np.exp(-1 * A_list[0] * B_aK_TK)
                    print("AK_BCF_Change_spectrum_Series", AK_BCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 5, QTableWidgetItem(self.B_aK_box.currentText()))
                    self.color_table.setItem(4, 6, QTableWidgetItem(f"{B_aK_TK}"))
                    return AK_BCF_Change_spectrum_Series
                elif B_aK_TK == min(Sort_TK_SP_AK_list):
                    print("min")
                    AK_BCF_Change_spectrum_Series = K_list[len(K_list) - 1] * np.exp(
                        -1 * A_list[len(A_list) - 1] * B_aK_TK)
                    print("AK_BCF_Change_spectrum_Series", AK_BCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 5, QTableWidgetItem(self.B_aK_box.currentText()))
                    self.color_table.setItem(4, 6, QTableWidgetItem(f"{B_aK_TK}"))
                    return AK_BCF_Change_spectrum_Series
                elif B_aK_TK in TK_SP_list:
                    print("equal")
                    AK_BCF_Change_spectrum_Series = SP_series_list[TK_SP_list.index(B_aK_TK)][1:].reset_index(drop=True)
                    print("AK_BCF_Change_spectrum_Series", AK_BCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 5, QTableWidgetItem(self.B_aK_box.currentText()))
                    self.color_table.setItem(4, 6, QTableWidgetItem(f"{B_aK_TK}"))
                    return AK_BCF_Change_spectrum_Series
                else:
                    print("mid")
                    position = Sort_TK_SP_AK_list.index(B_aK_TK)
                    AK_BCF_Change_spectrum_Series = K_list[position - 1] * np.exp(-1 * A_list[position - 1] * B_aK_TK)
                    print("AK_BCF_Change_spectrum_Series", AK_BCF_Change_spectrum_Series)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 5, QTableWidgetItem(self.B_aK_box.currentText()))
                    self.color_table.setItem(4, 6, QTableWidgetItem(f"{B_aK_TK}"))
                    return AK_BCF_Change_spectrum_Series
            # 關閉連線
            connection_BCF_Change.close()
        elif self.B_aK_mode.currentText() == "未選":
            # 防呆反灰設定
            self.B_aK_box.setEnabled(False)
            self.B_aK_table.setEnabled(False)
            self.B_aK_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.B_aK_table.setStyleSheet(QCOMBOBOXDISABLE)
            AK_BCF_Change_spectrum_Series = 1
            print("BCF_Change:未選,默認1")
            return AK_BCF_Change_spectrum_Series
        elif self.B_aK_mode.currentText() == "模擬":
            self.B_aK_box.setEnabled(True)
            self.B_aK_table.setEnabled(True)
            self.B_aK_box.setStyleSheet(QCOMBOXSETTING)
            self.B_aK_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            AK_BCF_Change_spectrum_Series = 1
            print("BCF_Change:模擬,默認1")
            return AK_BCF_Change_spectrum_Series

    def calculate_RCF_Differ(self):
        if self.check == "NO":
            print("Please select light source")
        if self.R_differ_mode.currentText() == "自訂":
            # 防呆反灰設定
            self.R_differ_box.setEnabled(True)
            self.R_differ_table.setEnabled(True)
            self.R_differ_box.setStyleSheet(QCOMBOXSETTING)
            self.R_differ_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            print("in_RCF_Differ自訂")
            connection_RCF_Differ = sqlite3.connect("RCF_Differ_spectrum.db")
            cursor_RCF_Differ = connection_RCF_Differ.cursor()
            # 取得Rdiffer資料
            column_name_RCF_Differ = self.R_differ_box.currentText()
            #print("column_name_RCF_Differ",column_name_RCF_Differ)
            table_name_RCF_Differ = self.R_differ_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_RCF_Differ = f"SELECT * FROM '{table_name_RCF_Differ}';"
            cursor_RCF_Differ.execute(query_RCF_Differ)
            result_RCF_Differ = cursor_RCF_Differ.fetchall()

            # 找到指定標題的欄位索引
            header_RCF_Differ = [column[0] for column in cursor_RCF_Differ.description]

            # 移除所有標題中的編號
            header_RCF_Differ = [re.sub(r'_\d+$', '', col) for col in header_RCF_Differ]
            #print("header_RCF_Differ", header_RCF_Differ)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_RCF_Differ) if col == column_name_RCF_Differ]
            #print("indices_without_number", indices_without_number)

            # 使用 pandas Series 存儲結果
            series_list = [pd.Series([row[i] for row in result_RCF_Differ]) for i in indices_without_number]

            #print("series_list", series_list)
            TK_record = []
            series_list_length = len(series_list)
            for i in range(series_list_length):
                try:
                    TK_record.append(float(series_list[i].iloc[0]))
                except ValueError:
                    C_Series = 1
                    # 發生錯誤時，忽略並繼續執行
                    return C_Series
            self.R_differ_TK_edit_label.setText(f"TKrange: {min(TK_record):.2f}~{max(TK_record):.2f}")
            if self.R_differ_TK_edit.text() == "":
                C_Series = 1
                return C_Series
            elif self.R_differ_TK_edit.text() != "":
                R_differ_TK = float(self.R_differ_TK_edit.text())
                Thickness_list = [R_differ_TK]
                for i in range(series_list_length):
                    Thickness_list.append(float(series_list[i].iloc[0]))
                    TK_record.append(float(series_list[i].iloc[0]))
                    #print(series_list[i].iloc[0])
                #print("Thickness_list",Thickness_list)
                #print(max(Thickness_list))
                # print(sorted(Thickness_list))
                Thickness_list_2 = sorted(Thickness_list)
                #print("Thickness_list_2", Thickness_list_2)
                index_R_differ_TK = Thickness_list_2.index(R_differ_TK)
                #print(Thickness_list_2.index(R_differ_TK))

                # 輸入厚度的三種情況
                if R_differ_TK == min(Thickness_list):
                    print("min")
                    A_index_value = Thickness_list_2[index_R_differ_TK + 1]
                    #print("A_index_value", A_index_value)
                    B_index_value = Thickness_list_2[index_R_differ_TK + 2]
                    #print("B_index_value", B_index_value)
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    #print("A_index", A_index)
                    B_index = Thickness_list.index(B_index_value)
                    #print("B_index", B_index)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()

                    # Final R Series
                    C_Series = A_Series + (R_differ_TK - A_index_value) * (B_Series - A_Series) / (
                                B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 1, QTableWidgetItem(self.R_differ_box.currentText()))
                    self.color_table.setItem(4, 2, QTableWidgetItem(f"{R_differ_TK}"))
                elif R_differ_TK == max(Thickness_list):
                    print("max")
                    A_index_value = Thickness_list_2[index_R_differ_TK - 2]
                    B_index_value = Thickness_list_2[index_R_differ_TK - 1]
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    B_index = Thickness_list.index(B_index_value)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()
                    # Final R Series
                    C_Series = A_Series + (R_differ_TK - A_index_value) * (B_Series - A_Series) / (
                                B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 1, QTableWidgetItem(self.R_differ_box.currentText()))
                    self.color_table.setItem(4, 2, QTableWidgetItem(f"{R_differ_TK}"))
                elif R_differ_TK in (Thickness_list[1:]):
                    print("equal")
                    # 取得對應的索引
                    Thickness_list.remove(Thickness_list[0])
                    equal_index = Thickness_list.index(R_differ_TK)
                    #print("Equal Index:", equal_index)
                    #print("Equal Value:", R_differ_TK)
                    C_Series = series_list[equal_index][1:]
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 1, QTableWidgetItem(self.R_differ_box.currentText()))
                    self.color_table.setItem(4, 2, QTableWidgetItem(f"{R_differ_TK}"))
                else:
                    print("mid")
                    #print("R_differ_TK",R_differ_TK)
                    #print("Thickness_list[1:]",Thickness_list[1:])
                    A_index_value = Thickness_list_2[index_R_differ_TK - 1]
                    #print("A_index_value",A_index_value)
                    B_index_value = Thickness_list_2[index_R_differ_TK + 1]
                    #print("B_index_value", B_index_value)
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    #print("A_index",A_index)
                    #print("Thickness_list",Thickness_list)
                    B_index = Thickness_list.index(B_index_value)
                    #print("B_index", B_index)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    #print("A_Series",A_Series)
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()
                    #print("B_Series", B_Series)

                    # Final R Series
                    C_Series = A_Series + (R_differ_TK - A_index_value) * (B_Series - A_Series) / (B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 1, QTableWidgetItem(self.R_differ_box.currentText()))
                    self.color_table.setItem(4, 2, QTableWidgetItem(f"{R_differ_TK}"))
                # 關閉連線
                connection_RCF_Differ.close()
                print("C_Series",C_Series)
                return C_Series
        elif self.R_differ_mode.currentText() == "未選":
            # 防呆反灰設定
            self.R_differ_box.setEnabled(False)
            self.R_differ_table.setEnabled(False)
            self.R_differ_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.R_differ_table.setStyleSheet(QCOMBOBOXDISABLE)
            C_Series = 1
            print("RCF_Differ:未選,默認1")
            return C_Series
        elif self.R_differ_mode.currentText() == "模擬":
            self.R_differ_box.setEnabled(True)
            self.R_differ_table.setEnabled(True)
            self.R_differ_box.setStyleSheet(QCOMBOXSETTING)
            self.R_differ_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            C_Series = 1
            print("RCF_Differ:模擬,默認1")
            return C_Series

    def calculate_GCF_Differ(self):
        if self.check == "NO":
            print("Please select light source")
        if self.G_differ_mode.currentText() == "自訂":
            # 防呆反灰設定
            self.G_differ_box.setEnabled(True)
            self.G_differ_table.setEnabled(True)
            self.G_differ_box.setStyleSheet(QCOMBOXSETTING)
            self.G_differ_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            print("in_GCF_Differ自訂")
            connection_GCF_Differ = sqlite3.connect("GCF_Differ_spectrum.db")
            cursor_GCF_Differ = connection_GCF_Differ.cursor()
            # 取得Rdiffer資料
            column_name_GCF_Differ = self.G_differ_box.currentText()
            #print("column_name_GCF_Differ",column_name_GCF_Differ)
            table_name_GCF_Differ = self.G_differ_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_GCF_Differ = f"SELECT * FROM '{table_name_GCF_Differ}';"
            cursor_GCF_Differ.execute(query_GCF_Differ)
            result_GCF_Differ = cursor_GCF_Differ.fetchall()

            # 找到指定標題的欄位索引
            header_GCF_Differ = [column[0] for column in cursor_GCF_Differ.description]

            # 移除所有標題中的編號
            header_GCF_Differ = [re.sub(r'_\d+$', '', col) for col in header_GCF_Differ]
            #print("header_GCF_Differ", header_GCF_Differ)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_GCF_Differ) if col == column_name_GCF_Differ]
            #print("indices_without_number", indices_without_number)

            # 使用 pandas Series 存儲結果
            series_list = [pd.Series([row[i] for row in result_GCF_Differ]) for i in indices_without_number]
            TK_record = []
            series_list_length = len(series_list)
            for i in range(series_list_length):
                try:
                    TK_record.append(float(series_list[i].iloc[0]))
                except ValueError:
                    C_Series = 1
                    # 發生錯誤時，忽略並繼續執行
                    return C_Series
            self.G_differ_TK_edit_label.setText(f"TKrange: {min(TK_record):.2f}~{max(TK_record):.2f}")

            #print("series_list", series_list)
            if self.G_differ_TK_edit.text() == "":
                C_Series = 1
                return C_Series
            elif self.G_differ_TK_edit.text() != "":
                G_differ_TK = float(self.G_differ_TK_edit.text())
                Thickness_list = [G_differ_TK]
                for i in range(series_list_length):
                    Thickness_list.append(float(series_list[i].iloc[0]))
                    print(series_list[i].iloc[0])
                #print("Thickness_list",Thickness_list)
                #print(max(Thickness_list))
                # print(sorted(Thickness_list))
                Thickness_list_2 = sorted(Thickness_list)
                #print("Thickness_list_2", Thickness_list_2)
                index_G_differ_TK = Thickness_list_2.index(G_differ_TK)
                #print(Thickness_list_2.index(G_differ_TK))

                # 輸入厚度的三種情況
                if G_differ_TK == min(Thickness_list):
                    print("min")
                    A_index_value = Thickness_list_2[index_G_differ_TK + 1]
                    #print("A_index_value", A_index_value)
                    B_index_value = Thickness_list_2[index_G_differ_TK + 2]
                    #print("B_index_value", B_index_value)
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    #print("A_index", A_index)
                    B_index = Thickness_list.index(B_index_value)
                    #print("B_index", B_index)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()

                    # Final G Series
                    C_Series = A_Series + (G_differ_TK - A_index_value) * (B_Series - A_Series) / (
                                B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 3, QTableWidgetItem(self.G_differ_box.currentText()))
                    self.color_table.setItem(4, 4, QTableWidgetItem(f"{G_differ_TK}"))
                elif G_differ_TK == max(Thickness_list):
                    print("max")
                    A_index_value = Thickness_list_2[index_G_differ_TK - 2]
                    B_index_value = Thickness_list_2[index_G_differ_TK - 1]
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    B_index = Thickness_list.index(B_index_value)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()
                    # Final R Series
                    C_Series = A_Series + (G_differ_TK - A_index_value) * (B_Series - A_Series) / (
                                B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 3, QTableWidgetItem(self.G_differ_box.currentText()))
                    self.color_table.setItem(4, 4, QTableWidgetItem(f"{G_differ_TK}"))
                elif G_differ_TK in (Thickness_list[1:]):
                    print("equal")
                    # 取得對應的索引
                    Thickness_list.remove(Thickness_list[0])
                    equal_index = Thickness_list.index(G_differ_TK)
                    #print("Equal Index:", equal_index)
                    #print("Equal Value:", R_differ_TK)
                    C_Series = series_list[equal_index][1:]
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 3, QTableWidgetItem(self.G_differ_box.currentText()))
                    self.color_table.setItem(4, 4, QTableWidgetItem(f"{G_differ_TK}"))
                else:
                    print("mid")
                    #print("R_differ_TK",R_differ_TK)
                    #print("Thickness_list[1:]",Thickness_list[1:])
                    A_index_value = Thickness_list_2[index_G_differ_TK - 1]
                    #print("A_index_value",A_index_value)
                    B_index_value = Thickness_list_2[index_G_differ_TK + 1]
                    #print("B_index_value", B_index_value)
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    #print("A_index",A_index)
                    #print("Thickness_list",Thickness_list)
                    B_index = Thickness_list.index(B_index_value)
                    #print("B_index", B_index)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    #print("A_Series",A_Series)
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()
                    #print("B_Series", B_Series)

                    # Final R Series
                    C_Series = A_Series + (G_differ_TK - A_index_value) * (B_Series - A_Series) / (B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 3, QTableWidgetItem(self.G_differ_box.currentText()))
                    self.color_table.setItem(4, 4, QTableWidgetItem(f"{G_differ_TK}"))
                # 關閉連線
                connection_GCF_Differ.close()
                print("C_Series",C_Series)
                return C_Series
        elif self.G_differ_mode.currentText() == "未選":
            # 防呆反灰設定
            self.G_differ_box.setEnabled(False)
            self.G_differ_table.setEnabled(False)
            self.G_differ_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.G_differ_table.setStyleSheet(QCOMBOBOXDISABLE)
            C_Series = 1
            print("GCF_Differ:未選,默認1")
            return C_Series
        elif self.G_differ_mode.currentText() == "模擬":
            self.G_differ_box.setEnabled(True)
            self.G_differ_table.setEnabled(True)
            self.G_differ_box.setStyleSheet(QCOMBOXSETTING)
            self.G_differ_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            C_Series = 1
            print("GCF_Differ:模擬,默認1")
            return C_Series

    def calculate_BCF_Differ(self):
        if self.check == "NO":
            print("Please select light source")
        if self.B_differ_mode.currentText() == "自訂":
            # 防呆反灰設定
            self.B_differ_box.setEnabled(True)
            self.B_differ_table.setEnabled(True)
            self.B_differ_box.setStyleSheet(QCOMBOXSETTING)
            self.B_differ_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            print("in_BCF_Differ自訂")
            connection_BCF_Differ = sqlite3.connect("BCF_Differ_spectrum.db")
            cursor_BCF_Differ = connection_BCF_Differ.cursor()
            # 取得Rdiffer資料
            column_name_BCF_Differ = self.B_differ_box.currentText()
            print("column_name_BCF_Differ",column_name_BCF_Differ)
            table_name_BCF_Differ = self.B_differ_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_BCF_Differ = f"SELECT * FROM '{table_name_BCF_Differ}';"
            cursor_BCF_Differ.execute(query_BCF_Differ)
            result_BCF_Differ = cursor_BCF_Differ.fetchall()

            # 找到指定標題的欄位索引
            header_BCF_Differ = [column[0] for column in cursor_BCF_Differ.description]

            # 移除所有標題中的編號
            header_BCF_Differ = [re.sub(r'_\d+$', '', col) for col in header_BCF_Differ]
            #print("header_BCF_Differ", header_BCF_Differ)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_BCF_Differ) if col == column_name_BCF_Differ]
            #print("indices_without_number", indices_without_number)

            # 使用 pandas Series 存儲結果
            series_list = [pd.Series([row[i] for row in result_BCF_Differ]) for i in indices_without_number]
            TK_record = []
            series_list_length = len(series_list)
            for i in range(series_list_length):
                try:
                    TK_record.append(float(series_list[i].iloc[0]))
                except ValueError:
                    C_Series = 1
                    # 發生錯誤時，忽略並繼續執行
                    return C_Series
            self.B_differ_TK_edit_label.setText(f"TKrange: {min(TK_record):.2f}~{max(TK_record):.2f}")

            #print("series_list", series_list)
            if self.B_differ_TK_edit.text() == "":
                C_Series = 1
                return C_Series
            elif self.B_differ_TK_edit.text() !="":
                B_differ_TK = float(self.B_differ_TK_edit.text())
                Thickness_list = [B_differ_TK]
                for i in range(series_list_length):
                    Thickness_list.append(float(series_list[i].iloc[0]))
                    print(series_list[i].iloc[0])
                print("Thickness_list",Thickness_list)
                #print(max(Thickness_list))
                # print(sorted(Thickness_list))
                Thickness_list_2 = sorted(Thickness_list)
                print("Thickness_list_2", Thickness_list_2)
                index_B_differ_TK = Thickness_list_2.index(B_differ_TK)
                #print(Thickness_list_2.index(B_differ_TK))
                # 輸入厚度的三種情況
                if B_differ_TK == min(Thickness_list):
                    print("min")
                    A_index_value = Thickness_list_2[index_B_differ_TK + 1]
                    #print("A_index_value", A_index_value)
                    B_index_value = Thickness_list_2[index_B_differ_TK + 2]
                    #print("B_index_value", B_index_value)
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    #print("A_index", A_index)
                    B_index = Thickness_list.index(B_index_value)
                    #print("B_index", B_index)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()

                    # Final B Series
                    C_Series = A_Series + (B_differ_TK - A_index_value) * (B_Series - A_Series) / (
                                B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 5, QTableWidgetItem(self.B_differ_box.currentText()))
                    self.color_table.setItem(4, 6, QTableWidgetItem(f"{B_differ_TK}"))
                elif B_differ_TK == max(Thickness_list):
                    print("max")
                    A_index_value = Thickness_list_2[index_B_differ_TK - 2]
                    B_index_value = Thickness_list_2[index_B_differ_TK - 1]
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    B_index = Thickness_list.index(B_index_value)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()
                    # Final R Series
                    C_Series = A_Series + (B_differ_TK - A_index_value) * (B_Series - A_Series) / (
                                B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 5, QTableWidgetItem(self.B_differ_box.currentText()))
                    self.color_table.setItem(4, 6, QTableWidgetItem(f"{B_differ_TK}"))
                elif B_differ_TK in (Thickness_list[1:]):
                    print("equal")
                    # 取得對應的索引
                    Thickness_list.remove(Thickness_list[0])
                    equal_index = Thickness_list.index(B_differ_TK)
                    #print("Equal Index:", equal_index)
                    #print("Equal Value:", R_differ_TK)
                    C_Series = series_list[equal_index][1:]
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 5, QTableWidgetItem(self.B_differ_box.currentText()))
                    self.color_table.setItem(4, 6, QTableWidgetItem(f"{B_differ_TK}"))
                else:
                    print("mid")
                    #print("R_differ_TK",R_differ_TK)
                    #print("Thickness_list[1:]",Thickness_list[1:])
                    A_index_value = Thickness_list_2[index_B_differ_TK - 1]
                    #print("A_index_value",A_index_value)
                    B_index_value = Thickness_list_2[index_B_differ_TK + 1]
                    #print("B_index_value", B_index_value)
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    #print("A_index",A_index)
                    #print("Thickness_list",Thickness_list)
                    B_index = Thickness_list.index(B_index_value)
                    #print("B_index", B_index)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    #print("A_Series",A_Series)
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()
                    #print("B_Series", B_Series)

                    # Final R Series
                    C_Series = A_Series + (B_differ_TK - A_index_value) * (B_Series - A_Series) / (B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    # 在Color_Table上顯示
                    self.color_table.setItem(4, 5, QTableWidgetItem(self.B_differ_box.currentText()))
                    self.color_table.setItem(4, 6, QTableWidgetItem(f"{B_differ_TK}"))
                # 關閉連線
                connection_BCF_Differ.close()
                print("C_Series",C_Series)
                return C_Series
        elif self.B_differ_mode.currentText() == "未選":
            # 防呆反灰設定
            self.B_differ_box.setEnabled(False)
            self.B_differ_table.setEnabled(False)
            self.B_differ_box.setStyleSheet(QCOMBOBOXDISABLE)
            self.B_differ_table.setStyleSheet(QCOMBOBOXDISABLE)
            C_Series = 1
            print("BCF_Differ:未選,默認1")
            return C_Series
        elif self.B_differ_mode.currentText() == "模擬":
            self.B_differ_box.setEnabled(True)
            self.B_differ_table.setEnabled(True)
            self.B_differ_box.setStyleSheet(QCOMBOXSETTING)
            self.B_differ_table.setStyleSheet(QCOMBOBOXTABLESELECT)
            C_Series = 1
            print("BCF_Differ:模擬,默認1")
            return C_Series

    def calculate_color_customize(self):
        if self.check == "NO":
            self.BLUcheck = "NO"
            print("Please select light source")
            print("self.BLUcheck",self.BLUcheck)
            return
        self.light_source_mode.setEnabled(True)
        self.BLUcheck = "OK"
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
        RxSxxl_C = CIE_spectrum_Series_X * C_spectrum_Series
        RxSxyl_C = CIE_spectrum_Series_Y * C_spectrum_Series
        RxSxzl_C = CIE_spectrum_Series_Z * C_spectrum_Series

        C_spectrum_Series_sum = C_spectrum_Series.sum()
        k = 100 / C_spectrum_Series_sum
        RxSxxl_C_sum = RxSxxl_C.sum()
        RxSxyl_C_sum = RxSxyl_C.sum()
        RxSxzl_C_sum = RxSxzl_C.sum()
        RxSxxl_C_sum_k = RxSxxl_C_sum * k
        RxSxyl_C_sum_k = RxSxyl_C_sum * k
        RxSxzl_C_sum_k = RxSxzl_C_sum * k
        BLU_C_x = RxSxxl_C_sum_k / (RxSxxl_C_sum_k + RxSxyl_C_sum_k + RxSxzl_C_sum_k)
        BLU_C_y = RxSxyl_C_sum_k / (RxSxxl_C_sum_k + RxSxyl_C_sum_k + RxSxzl_C_sum_k)
        self.color_table.setItem(2, 20, QTableWidgetItem(f"{BLU_C_x:.3f}"))
        self.color_table.setItem(2, 21, QTableWidgetItem(f"{BLU_C_y:.3f}"))

        if self.calculate_BLU() is not None:
            # 計算BLU------------------------------------------------------------------
            self.calculate_BLU()
            # print("self.calculate_BLU()",self.calculate_BLU())

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
            self.color_table.setItem(1, 20, QTableWidgetItem(f"{BLU_x:.3f}"))
            self.color_table.setItem(1, 21, QTableWidgetItem(f"{BLU_y:.3f}"))

            # if self.calculate_BLU() is not None and self.calculate_RCF_Fix() is not None\
            #         and self.calculate_GCF_Fix() is not None and self.calculate_BCF_Fix() is not None:
            #     self.calculate_BLU()
            #     self.calculate_layer1()
            #     self.calculate_layer2()
            #     self.calculate_RCF_Fix()
            #     self.calculate_GCF_Fix()
            #     self.calculate_BCF_Fix()
            # BLU +Cell part
            cell_blu_total_spectrum = self.calculate_BLU() * self.calculate_layer1() * self.calculate_layer2() \
                                      * self.calculate_layer3() * self.calculate_layer4() * self.calculate_layer5() \
                                      * self.calculate_layer6()
            # print("self.calculate_BLU()",self.calculate_BLU())
            # print("cell_blu_total_spectrum",cell_blu_total_spectrum)
            # R_Fix & R_AK
            R = cell_blu_total_spectrum * self.calculate_RCF_Fix() * self.calculate_RCF_Change() * self.calculate_RCF_Differ()
            print("cell_blu_total_spectrum",cell_blu_total_spectrum)
            print("self.calculate_RCF_Fix()",self.calculate_RCF_Fix())
            print("self.calculate_RCF_Change()",self.calculate_RCF_Change())
            print("self.calculate_RCF_Differ()",self.calculate_RCF_Differ())
            print("R", R)
            R_X = R * CIE_spectrum_Series_X
            #print("R_X", R_X)
            R_Y = R * CIE_spectrum_Series_Y
            #print("R_Y", R_Y)
            R_Z = R * CIE_spectrum_Series_Z
            #print("R_Z", R_Z)
            self.R_X_sum = R_X.sum()
            #print("R_X_sum", self.R_X_sum)
            self.R_Y_sum = R_Y.sum()
            #print("R_Y_sum", self.R_Y_sum)
            self.R_Z_sum = R_Z.sum()
            #print("R_Z_sum", self.R_Z_sum)
            self.R_x = self.R_X_sum / (self.R_X_sum + self.R_Y_sum + self.R_Z_sum)
            self.R_y = self.R_Y_sum / (self.R_X_sum + self.R_Y_sum + self.R_Z_sum)
            self.R_T = ((((R / self.calculate_BLU()) * CIE_spectrum_Series_Y).sum())/ (self.calculate_BLU() * CIE_spectrum_Series_Y).sum())*100
            print("self.R_T",self.R_T)

            self.color_table.setItem(1, 4, QTableWidgetItem(f"{self.R_x:.4f}"))
            self.color_table.setItem(1, 5, QTableWidgetItem(f"{self.R_y:.4f}"))
            self.color_table.setItem(1, 6, QTableWidgetItem(f"{self.R_T:.3f}%"))

            # G_Fix & GAK
            G = cell_blu_total_spectrum * self.calculate_GCF_Fix() * self.calculate_GCF_Change() * self.calculate_GCF_Differ()
            G_X = G * CIE_spectrum_Series_X
            G_Y = G * CIE_spectrum_Series_Y
            G_Z = G * CIE_spectrum_Series_Z
            self.G_X_sum = G_X.sum()
            self.G_Y_sum = G_Y.sum()
            self.G_Z_sum = G_Z.sum()
            self.G_x = self.G_X_sum / (self.G_X_sum + self.G_Y_sum + self.G_Z_sum)
            self.G_y = self.G_Y_sum / (self.G_X_sum + self.G_Y_sum + self.G_Z_sum)
            self.G_T = ((((G / self.calculate_BLU()) * CIE_spectrum_Series_Y).sum())/ (self.calculate_BLU() * CIE_spectrum_Series_Y).sum())*100
            self.color_table.setItem(1, 9, QTableWidgetItem(f"{self.G_x:.4f}"))
            self.color_table.setItem(1, 10, QTableWidgetItem(f"{self.G_y:.4f}"))
            self.color_table.setItem(1, 11, QTableWidgetItem(f"{self.G_T:.3f}%"))

            # B_Fix & BAK
            B = cell_blu_total_spectrum * self.calculate_BCF_Fix() * self.calculate_BCF_Change() * self.calculate_BCF_Differ()
            B_X = B * CIE_spectrum_Series_X
            B_Y = B * CIE_spectrum_Series_Y
            B_Z = B * CIE_spectrum_Series_Z
            self.B_X_sum = B_X.sum()
            self.B_Y_sum = B_Y.sum()
            self.B_Z_sum = B_Z.sum()
            self.B_x = self.B_X_sum / (self.B_X_sum + self.B_Y_sum + self.B_Z_sum)
            self.B_y = self.B_Y_sum / (self.B_X_sum + self.B_Y_sum + self.B_Z_sum)
            self.B_T = ((((B / self.calculate_BLU()) * CIE_spectrum_Series_Y).sum())/ (self.calculate_BLU() * CIE_spectrum_Series_Y).sum())*100
            self.color_table.setItem(1, 14, QTableWidgetItem(f"{self.B_x:.4f}"))
            self.color_table.setItem(1, 15, QTableWidgetItem(f"{self.B_y:.4f}"))
            self.color_table.setItem(1, 16, QTableWidgetItem(f"{self.B_T:.3f}%"))

            # W
            W_X = R_X + G_X + B_X
            W_Y = R_Y + G_Y + B_Y
            W_Z = R_Z + G_Z + B_Z
            W_X_sum = W_X.sum()
            W_Y_sum = W_Y.sum()
            W_Z_sum = W_Z.sum()
            self.W_x = W_X_sum / (W_X_sum + W_Y_sum + W_Z_sum)
            self.W_y = W_Y_sum / (W_X_sum + W_Y_sum + W_Z_sum)
            self.W_T = (self.R_T + self.G_T + self.B_T) / 3
            self.color_table.setItem(1, 1, QTableWidgetItem(f"{self.W_x:.4f}"))
            self.color_table.setItem(1, 2, QTableWidgetItem(f"{self.W_y:.4f}"))
            self.color_table.setItem(1, 3, QTableWidgetItem(f"{self.W_T:.3f}%"))

            # NTSC
            NTSC = 100 * 0.5 * abs((self.R_x * self.G_y + self.G_x * self.B_y + self.B_x * self.R_y - (
                    self.G_x * self.R_y) - (self.B_x * self.G_y) - (self.R_x * self.B_y))) / 0.1582
            self.color_table.setItem(1, 19, QTableWidgetItem(f"{NTSC:.3f}%"))

            # wave purity
            self.calculate_All_wave_P()

            # Clight +Cell part(只算Clight其他layer沒有)
            cell_C_total_spectrum = C_spectrum_Series
            # print("self.calculate_BLU()",self.calculate_BLU())
            #print("cell_C_total_spectrum",cell_C_total_spectrum)
            # RC
            RC = cell_C_total_spectrum * self.calculate_RCF_Fix() * self.calculate_RCF_Change() * self.calculate_RCF_Differ()
            print("RC",RC)
            RC_X = RC * CIE_spectrum_Series_X
            RC_Y = RC * CIE_spectrum_Series_Y
            RC_Z = RC * CIE_spectrum_Series_Z
            RC_X_sum = RC_X.sum()
            RC_Y_sum = RC_Y.sum()
            RC_Z_sum = RC_Z.sum()
            self.RC_x = RC_X_sum / (RC_X_sum + RC_Y_sum + RC_Z_sum)
            self.RC_y = RC_Y_sum / (RC_X_sum + RC_Y_sum + RC_Z_sum)
            # 能量積分再相除
            RC_T = ((((RC / cell_C_total_spectrum) * CIE_spectrum_Series_Y).sum())/ (cell_C_total_spectrum * CIE_spectrum_Series_Y).sum())*100
            print("RC_Y",RC_Y)
            print("RC_T",RC_T)

            self.color_table.setItem(2, 4, QTableWidgetItem(f"{self.RC_x:.4f}"))
            self.color_table.setItem(2, 5, QTableWidgetItem(f"{self.RC_y:.4f}"))
            self.color_table.setItem(2, 6, QTableWidgetItem(f"{RC_T:.3f}%"))

            # GC
            GC = cell_C_total_spectrum * self.calculate_GCF_Fix() * self.calculate_GCF_Change() * self.calculate_GCF_Differ()
            GC_X = GC * CIE_spectrum_Series_X
            GC_Y = GC * CIE_spectrum_Series_Y
            GC_Z = GC * CIE_spectrum_Series_Z
            GC_X_sum = GC_X.sum()
            GC_Y_sum = GC_Y.sum()
            GC_Z_sum = GC_Z.sum()
            self.GC_x = GC_X_sum / (GC_X_sum + GC_Y_sum + GC_Z_sum)
            self.GC_y = GC_Y_sum / (GC_X_sum + GC_Y_sum + GC_Z_sum)
            GC_T = ((((GC / cell_C_total_spectrum) * CIE_spectrum_Series_Y).sum())/ (cell_C_total_spectrum * CIE_spectrum_Series_Y).sum())*100
            self.color_table.setItem(2, 9, QTableWidgetItem(f"{self.GC_x:.4f}"))
            self.color_table.setItem(2, 10, QTableWidgetItem(f"{self.GC_y:.4f}"))
            self.color_table.setItem(2, 11, QTableWidgetItem(f"{GC_T:.3f}%"))

            # BC
            BC = cell_C_total_spectrum * self.calculate_BCF_Fix() * self.calculate_BCF_Change() * self.calculate_BCF_Differ()
            BC_X = BC * CIE_spectrum_Series_X
            BC_Y = BC * CIE_spectrum_Series_Y
            BC_Z = BC * CIE_spectrum_Series_Z
            BC_X_sum = BC_X.sum()
            BC_Y_sum = BC_Y.sum()
            BC_Z_sum = BC_Z.sum()
            self.BC_x = BC_X_sum / (BC_X_sum + BC_Y_sum + BC_Z_sum)
            self.BC_y = BC_Y_sum / (BC_X_sum + BC_Y_sum + BC_Z_sum)
            BC_T = ((((BC / cell_C_total_spectrum) * CIE_spectrum_Series_Y).sum())/ (cell_C_total_spectrum * CIE_spectrum_Series_Y).sum())*100
            self.color_table.setItem(2, 14, QTableWidgetItem(f"{self.BC_x:.4f}"))
            self.color_table.setItem(2, 15, QTableWidgetItem(f"{self.BC_y:.4f}"))
            self.color_table.setItem(2, 16, QTableWidgetItem(f"{BC_T:.3f}%"))

            # W
            WC_X = RC_X + GC_X + BC_X
            WC_Y = RC_Y + GC_Y + BC_Y
            WC_Z = RC_Z + GC_Z + BC_Z
            WC_X_sum = WC_X.sum()
            WC_Y_sum = WC_Y.sum()
            WC_Z_sum = WC_Z.sum()
            WC_x = WC_X_sum / (WC_X_sum + WC_Y_sum + WC_Z_sum)
            WC_y = WC_Y_sum / (WC_X_sum + WC_Y_sum + WC_Z_sum)
            WC_T = (RC_T + GC_T + BC_T) / 3
            self.color_table.setItem(2, 1, QTableWidgetItem(f"{WC_x:.4f}"))
            self.color_table.setItem(2, 2, QTableWidgetItem(f"{WC_y:.4f}"))
            self.color_table.setItem(2, 3, QTableWidgetItem(f"{WC_T:.3f}%"))

            # wave p
            self.calculate_All_2_wave_P()

            # NTSCC
            NTSCC = 100 * 0.5 * abs((self.RC_x * self.GC_y + self.GC_x * self.BC_y + self.BC_x * self.RC_y - (self.GC_x * self.RC_y) - (
                    self.BC_x * self.GC_y) - (self.RC_x * self.BC_y))) / 0.1582
            self.color_table.setItem(2, 19, QTableWidgetItem(f"{NTSCC:.3f}%"))
            # 關閉連線
            connection_CIE.close()

    def calculate_WPC(self):
        if self.check == "NO":
            self.BLUcheck = "NO"
            print("Please select light source")
            print("self.BLUcheck",self.BLUcheck)
            return
        self.light_source_mode.setEnabled(True)
        self.BLUcheck = "OK"
        if self.WPCx.text() == '':
            return
        target_W_x = float(self.WPCx.text())
        target_W_y = float(self.WPCy.text())

        Matrix_R = np.array([[self.R_x / self.R_y], [1], [(1 - self.R_x - self.R_y) / self.R_y]])
        print("self.R_x",self.R_x)
        print("self.R_y",self.R_y)
        Matrix_G = np.array([[self.G_x / self.G_y], [1], [(1 - self.G_x - self.G_y) / self.G_y]])
        print("self.G_x", self.G_x)
        print("self.G_y", self.G_y)
        Matrix_B = np.array([[self.B_x / self.B_y], [1], [(1 - self.B_x - self.B_y) / self.B_y]])
        print("self.B_x", self.B_x)
        print("self.B_y", self.B_y)
        Matrix_W = np.array([[self.W_x / self.W_y], [1], [(1 - self.W_x - self.W_y) / self.W_y]])
        print("self.W_x", self.W_x)
        print("self.W_y", self.W_y)
        print("Matrix_W",Matrix_W)
        Matrix_W_target = np.array([[target_W_x / target_W_y], [1], [(1 - target_W_x - target_W_y) / target_W_y]])
        print("Matrix_W_target",Matrix_W_target)

        Matrix_total = np.hstack((Matrix_R, Matrix_G, Matrix_B))
        print("Matrix_total", Matrix_total)
        Matrix_total_inverse = np.linalg.inv(Matrix_total)
        print("Matrix_total_inverse",Matrix_total_inverse)
        F_Matrix = Matrix_total_inverse.dot(Matrix_W)
        print("F_Matrix",F_Matrix)

        #print(F_Matrix[0][0])
        Matrix_PR = Matrix_R * F_Matrix[0][0]
        Matrix_PG = Matrix_G * F_Matrix[1][0]
        Matrix_PB = Matrix_B * F_Matrix[2][0]

        self.Matrix_Ptotal = np.hstack((Matrix_PR, Matrix_PG, Matrix_PB))
        Matrix_Ptotal_inverse = np.linalg.inv(self.Matrix_Ptotal)

        print("self.Matrix_Ptotal",self.Matrix_Ptotal)


        Matrixrgb = Matrix_Ptotal_inverse.dot(Matrix_W_target)

        print("Matrixrgb",Matrixrgb)

        Matrixrgb_norm = Matrixrgb / max(Matrixrgb)
        print("Matrixrgb_norm",Matrixrgb_norm)

        ratio_total = Matrixrgb_norm[0][0] + Matrixrgb_norm[1][0] + Matrixrgb_norm[2][0]
        print("Matrixrgb_norm[0][0]",Matrixrgb_norm[0][0])
        print("Matrixrgb_norm[1][0]",Matrixrgb_norm[1][0])
        print("Matrixrgb_norm[2][0]",Matrixrgb_norm[2][0])
        print("ratio_total",ratio_total)

        ratio_r = 3 / ratio_total * Matrixrgb_norm[0][0]
        ratio_g = 3 / ratio_total * Matrixrgb_norm[1][0]
        ratio_b = 3 / ratio_total * Matrixrgb_norm[2][0]
        print("ratio_r",ratio_r)
        print("ratio_g",ratio_g)
        print("ratio_b",ratio_b)
        Matrix_newrgb = [[ratio_r],[ratio_g],[ratio_b]]

        W_T_Matrix_BM = self.Matrix_Ptotal.dot(Matrixrgb_norm)
        print("W_T_Matrix_BM",W_T_Matrix_BM)
        W_T_Matrix_newrgb = self.Matrix_Ptotal.dot(Matrix_newrgb)
        BMitems = [str(Matrixrgb_norm[0][0]),str(Matrixrgb_norm[1][0]),str(Matrixrgb_norm[2][0]),str(W_T_Matrix_BM[1][0]*self.W_T)]
        for column,BMitem in enumerate(BMitems):
            # 首先格式化數值，然後創建 QTableWidgetItem
            formatted_item = f"{float(BMitem):.3f}"
            item = QTableWidgetItem(formatted_item)
            self.WPCTable.setItem(1, column + 1, item)
        newrgbs = [str(ratio_r),str(ratio_g),str(ratio_b),str(W_T_Matrix_newrgb[1][0]*self.W_T)]
        for column,newrgb in enumerate(newrgbs):
            # 首先格式化數值，然後創建 QTableWidgetItem
            formatted_item = f"{float(newrgb):.3f}"
            item = QTableWidgetItem(formatted_item)
            self.WPCTable.setItem(2, column + 1, item)
        # 为 BM 法繪製三個獨立的長方圖
        colors = ['red', 'green', 'blue']
        if self.WPCCombobox.currentText() == "WPC-BM遮擋設計":
            for i, color in enumerate(colors):
                self.draw_single_bar_chart(Matrixrgb_norm[i][0], color, i)
        else:
            # 计算 ratio_r, ratio_g, ratio_b 并绘制 newrgb 方法的长方图
            newrgbs = [ratio_r, ratio_g, ratio_b]
            for i, value in enumerate(newrgbs):
                self.draw_single_bar_chart_rgb(value, colors[i], i)

    def calculate_WPC_reverse(self):
        # 斷開信號以避免遞歸呼叫
        self.WPCTable.cellChanged.disconnect(self.calculate_WPC_reverse)

        try:
            if self.WPCTable.item(4, 1).text() and self.WPCTable.item(4, 2).text() and self.WPCTable.item(4,
                                                                                                          3).text() != "":
                R_ratio = float(self.WPCTable.item(4, 1).text())
                G_ratio = float(self.WPCTable.item(4, 2).text())
                B_ratio = float(self.WPCTable.item(4, 3).text())
                Matrix_RGB = [[R_ratio],[G_ratio],[B_ratio]]
                RX = self.R_X_sum * R_ratio
                RY = self.R_Y_sum * R_ratio
                RZ = self.R_Z_sum * R_ratio
                GX = self.G_X_sum * G_ratio
                GY = self.G_Y_sum * G_ratio
                GZ = self.G_Z_sum * G_ratio
                BX = self.B_X_sum * B_ratio
                BY = self.B_Y_sum * B_ratio
                BZ = self.B_Z_sum * B_ratio
                SUM = RX + RY + RZ + GX + GY + GZ + BX + BY + BZ
                Wx = (RX + GX + BX) / SUM
                Wy = (RY + GY + BY) / SUM
                WT = self.Matrix_Ptotal.dot(Matrix_RGB)
                self.WPCTable.setItem(4, 4, QTableWidgetItem(f"{Wx:.3f}"))
                self.WPCTable.setItem(4, 5, QTableWidgetItem(f"{Wy:.3f}"))
                self.WPCTable.setItem(4, 6, QTableWidgetItem(f"{WT[1][0]*self.W_T:.3f}"))
        except Exception as e:
            print(f"Error: {e}")
        finally:
            # 再次連接信號
            self.WPCTable.cellChanged.connect(self.calculate_WPC_reverse)

    def calculate_WPC_trs_ratio(self):
        # 斷開信號以避免遞歸呼叫
        self.WPCTable.cellChanged.disconnect(self.calculate_WPC_trs_ratio)

        try:
            if self.WPCTable.item(5, 1).text() and self.WPCTable.item(6, 1).text() \
                    and self.WPCTable.item(1, 4).text() and self.WPCTable.item(2, 4).text() != "":
                original_ratio = float(self.WPCTable.item(5, 1).text())
                modify_ratio = float(self.WPCTable.item(6, 1).text())
                BM_T = float(self.WPCTable.item(1, 4).text())
                rgb_T = float(self.WPCTable.item(2, 4).text())
                change_ratio = modify_ratio / original_ratio
                new_BM_T = BM_T * change_ratio
                mew_rgb_T = rgb_T * change_ratio
                self.WPCTable.setItem(1, 5, QTableWidgetItem(f"{new_BM_T:.3f}"))
                self.WPCTable.setItem(2, 5, QTableWidgetItem(f"{mew_rgb_T:.3f}"))
        except Exception as e:
            print(f"Error: {e}")
        finally:
            # 再次連接信號
            self.WPCTable.cellChanged.connect(self.calculate_WPC_trs_ratio)

    def calculate_WPC_trs_ratio_2(self):
        # 斷開信號以避免遞歸呼叫
        self.WPCTable.cellChanged.disconnect(self.calculate_WPC_trs_ratio_2)

        try:
            if self.WPCTable.item(5, 1).text() and self.WPCTable.item(6, 1).text() \
                    and self.WPCTable.item(4, 6).text()!= "":
                original_ratio = float(self.WPCTable.item(5, 1).text())
                modify_ratio = float(self.WPCTable.item(6, 1).text())
                self_T = float(self.WPCTable.item(4, 6).text())
                change_ratio = modify_ratio / original_ratio
                new_self_T = self_T * change_ratio

                self.WPCTable.setItem(4, 7, QTableWidgetItem(f"{new_self_T:.3f}"))
        except Exception as e:
            print(f"Error: {e}")
        finally:
            # 再次連接信號
            self.WPCTable.cellChanged.connect(self.calculate_WPC_trs_ratio_2)


    def draw_single_bar_chart(self, value, color, index):
        # 检查是否已有图表，如果有则移除
        if self.bm_chart_canvases[index]:
            self.bm_layout.removeWidget(self.bm_chart_canvases[index])
            self.bm_chart_canvases[index].deleteLater()

        # 创建一个 matplotlib 图形和轴
        fig = Figure(figsize=(2,2))
        ax = fig.add_subplot(111)

        base_width = 20  # 长方图的基础宽度
        bar_height = 100 * value  # 根据比例值调整高度

        rect = patches.Rectangle((0, 0), base_width, bar_height, color=color)
        ax.add_patch(rect)

        max_height = max(bar_height, 100)  # 长方图的最大高度

        ax.set_xlim(0, base_width)
        ax.set_ylim(0, max_height)

        # 设置 Y 轴的刻度
        ax.set_yticks(range(0, int(max_height) + 1, 10))
        ax.yaxis.set_visible(True)  # 显示 Y 轴
        ax.xaxis.set_visible(False)  # 隐藏 X 轴
        # 设置坐标轴刻度字体大小
        ax.tick_params(axis='y', labelsize=8)  # 设置 Y 轴刻度字体为 10
        # 使用 tight_layout 自动调整图表布局
        fig.tight_layout()

        # 将图形添加到 FigureCanvas，并将它添加到布局中
        canvas = FigureCanvas(fig)
        self.bm_layout.addWidget(canvas, 0, index, 1, 1)  # 放置在 bm_layout 的指定位置

        # 保存图表的引用
        self.bm_chart_canvases[index] = canvas
        # 打印长方图高度值
        print(f"{color.capitalize()} 长方图的高度值: {bar_height}")

    def draw_single_bar_chart_rgb(self, value, color, index):
        # 检查是否已有图表，如果有则移除
        if self.rgb_chart_canvases[index]:
            self.rgb_layout.removeWidget(self.rgb_chart_canvases[index])
            self.rgb_chart_canvases[index].deleteLater()

        # 创建一个 matplotlib 图形和轴
        fig = Figure(figsize=(2, 2))
        ax = fig.add_subplot(111)

        bar_width = 100 * value  # 根据比例值调整宽度
        base_height = 20

        rect = patches.Rectangle((0, 0), bar_width, base_height, color=color)
        ax.add_patch(rect)

        ax.set_xlim(0, max(bar_width, 150))  # 确保至少显示出 100 单位宽度
        ax.set_ylim(0, base_height)

        ax.yaxis.set_visible(False)  # 隐藏 Y 轴
        ax.xaxis.set_visible(True)  # 显示 X 轴
        ax.set_xticks(range(0, 150, 20))
        ax.tick_params(axis='x', labelsize=8)
        fig.tight_layout()

        # 将图形添加到 FigureCanvas，并将它添加到布局中
        canvas = FigureCanvas(fig)
        self.rgb_layout.addWidget(canvas, 0, index, 1, 1)  # 放置在 rgb_layout 的指定位置

        # 保存图表的引用
        self.rgb_chart_canvases[index] = canvas

    # Plot CIE
    def drawciechart(self):
        if self.check == "NO":
            self.BLUcheck = "NO"
            print("Please select light source")
            print("self.BLUcheck",self.BLUcheck)
            return
        self.light_source_mode.setEnabled(True)
        self.BLUcheck = "OK"
        # 检查是否已定义 'W_x' 属性
        if not hasattr(self, 'W_x') or self.W_x is None:
            return
        # 清空畫布
        self.CIEchart.clear()
        RGB_values = []
        # 4組 x 和 y 色座標
        x_coords = [self.W_x, self.R_x, self.G_x, self.B_x]
        y_coords = [self.W_y, self.R_y, self.G_y, self.B_y]

        for x, y in zip(x_coords, y_coords):
            XYZ = xy_to_XYZ(np.array([x, y]))
            RGB = XYZ_to_sRGB(XYZ)
            RGB_values.append(RGB)

            # 转换为 numpy 数组
        RGB_array = np.array(RGB_values)

        # 顯示對應的顏色在 CIE 1931 色度圖上
        plotting.plot_chromaticity_diagram_CIE1931(axes=self.CIEchart, show=False)

        # 使用 self.CIEchart.scatter 添加自定義的標記
        markers = ['o', 's', '^', '*']  # 可以根據需要更改標記形狀
        labels = ['W', 'R', 'G', 'B']
        # 設定外框的顏色為黑色
        edgecolors = 'black'
        # # 定義每個標記的顏色，這裡使用RGB表示，你可以根據需要調整顏色值
        # colors = ["white", (1, 0, 0), "#019858", "#2894FF"]
        for x, y, marker, label, RGB in zip(x_coords, y_coords, markers, labels, RGB_array):
            # 确保颜色值在 0 到 1 的范围内
            RGB_clipped = np.clip(RGB, 0, 1)
            self.CIEchart.scatter(x, y, marker=marker, label=f'{label}:({x:.3f}, {y:.3f})', color=[RGB_clipped],
                                  edgecolors=edgecolors)
        plt.tight_layout()

        # 顯示圖例
        self.CIEchart.legend(loc='upper right', bbox_to_anchor=(1.5, 0.6), fontsize='small')
        self.customcanvas_CIE.draw()

        # 连接滚轮事件
        self.customcanvas_CIE.mpl_connect('scroll_event', self.zoom_CIE_on_scroll)
        self.customcanvas_CIE.mpl_connect('button_press_event', self.on_CIE_press)
        self.customcanvas_CIE.mpl_connect('motion_notify_event', self.on_CIE_drag)
        self.customcanvas_CIE.mpl_connect('button_release_event', self.on_CIE_release)
        self.customcanvas_CIE.mpl_connect('motion_notify_event', self.hover_CIE)

        self.dragging = False

        # 创建注释，用于显示信息
        self.annot_CIE = self.CIEchart.annotate("", xy=(0, 0), xytext=(-20, 20),
                                                 textcoords="offset points",
                                                 bbox=dict(boxstyle="round", fc="w"),
                                                 arrowprops=dict(arrowstyle="->"))
        self.annot_CIE.set_visible(False)

    # FUNCTION
    def zoom_CIE_on_scroll(self, event: MouseEvent):
        # 缩放函数
        base_scale = 1.1  # 缩放基数
        cur_xlim = self.CIEchart.get_xlim()
        cur_ylim = self.CIEchart.get_ylim()
        cur_xrange = (cur_xlim[1] - cur_xlim[0]) * .5
        cur_yrange = (cur_ylim[1] - cur_ylim[0]) * .5

        xdata = event.xdata  # 获取事件发生时的 X 坐标
        ydata = event.ydata  # 获取事件发生时的 Y 坐标

        if event.button == 'up':  # 向上滚动放大
            scale_factor = 1 / base_scale
        elif event.button == 'down':  # 向下滚动缩小
            scale_factor = base_scale
        else:
            scale_factor = 1

        # 设置新的限制
        self.CIEchart.set_xlim([xdata - cur_xrange * scale_factor,
                                   xdata + cur_xrange * scale_factor])
        self.CIEchart.set_ylim([ydata - cur_yrange * scale_factor,
                                   ydata + cur_yrange * scale_factor])

        self.customcanvas_CIE.draw()

    def on_CIE_press(self, event: MouseEvent):
        if event.button == 1:  # 检查是否为鼠标左键
            self.dragging = True
            self.drag_start_x = event.xdata
            self.drag_start_y = event.ydata

    def on_CIE_drag(self, event: MouseEvent):
        if self.dragging:
            dx = event.xdata - self.drag_start_x
            dy = event.ydata - self.drag_start_y
            self.drag_start_x = event.xdata
            self.drag_start_y = event.ydata
            self.pan_CIE(dx, dy)

    def on_CIE_release(self, event: MouseEvent):
        self.dragging = False

    def pan_CIE(self, dx, dy):
        cur_xlim = self.CIEchart.get_xlim()
        cur_ylim = self.CIEchart.get_ylim()

        self.CIEchart.set_xlim(cur_xlim[0] - dx, cur_xlim[1] - dx)
        self.CIEchart.set_ylim(cur_ylim[0] - dy, cur_ylim[1] - dy)

        self.customcanvas_CIE.draw()

    def hover_CIE(self, event):
        # 鼠标悬停事件处理函数
        vis = self.annot_CIE.get_visible()
        if event.inaxes == self.CIEchart:
            for scatter in self.CIEchart.collections:
                cont, ind = scatter.contains(event)
                if cont:
                    pos = scatter.get_offsets()[ind["ind"][0]]
                    self.annot_CIE.xy = pos
                    text = f"{pos[0]:.3f}, {pos[1]:.3f}"
                    self.annot_CIE.set_text(text)
                    self.annot_CIE.set_visible(True)
                    self.customcanvas_CIE.draw_idle()
                    return
        if vis:
            self.annot_CIE.set_visible(False)
            self.customcanvas_CIE.draw_idle()

    def drawcie_W_samplechart(self):
        # 4組 x 和 y 色座標
        x_coords = [self.W_x]
        y_coords = [self.W_y]

        # 轉換為 numpy array
        x_coords_array = np.array(x_coords)
        y_coords_array = np.array(y_coords)

        # 將 xy 座標轉換為 CIE 1931 XYZ 三刺激值
        XYZ = xy_to_XYZ(np.column_stack((x_coords_array, y_coords_array)))

        # 將 XYZ 三刺激值轉換為 sRGB 顏色表示
        RGB = XYZ_to_sRGB(XYZ)

        # 顯示座標的顏色
        # show=False 控制是否顯示即時的繪圖，當你設置 show=False 時，繪圖將不會立即顯示在畫面上，你可以在需要顯示的時候調用 plt.show()。
        # 這樣可以讓你在繪製多個圖形時先完成所有的繪製操作再顯示，以提高效率。
        # 顯示座標的顏色
        for i in range(len(x_coords)):
            XYZ = xy_to_XYZ(np.array([x_coords[i], y_coords[i]]))
            RGB = XYZ_to_sRGB(XYZ)
            plot_single_colour_swatch(RGB, swatch_name=f"Color {i + 1}", xy=(x_coords[i], y_coords[i]),
                                      axes=self.CIEsamplechart_W, show=False)

        # 設定框線樣式
        for spine in self.CIEsamplechart_W.spines.values():
            spine.set_visible(True)
            spine.set_color('black')
            spine.set_linewidth(1)

        self.CIEsamplechart_W.axis('off')  # 關閉座標軸

        # # 顯示對應的顏色在 CIE 1931 色度圖上
        # plotting.models.plot_RGB_chromaticities_in_chromaticity_diagram(RGB, axes=self.CIEsamplechart, show=False)
        # 顯示圖例
        self.CIEsamplechart_W.legend()
        self.customcanvas_W.draw()

    def drawcie_R_samplechart(self):
        # 4組 x 和 y 色座標
        x_coords = [self.R_x]
        y_coords = [self.R_y]

        # 轉換為 numpy array
        x_coords_array = np.array(x_coords)
        y_coords_array = np.array(y_coords)

        # 將 xy 座標轉換為 CIE 1931 XYZ 三刺激值
        XYZ = xy_to_XYZ(np.column_stack((x_coords_array, y_coords_array)))

        # 將 XYZ 三刺激值轉換為 sRGB 顏色表示
        RGB = XYZ_to_sRGB(XYZ)

        # 顯示座標的顏色
        # show=False 控制是否顯示即時的繪圖，當你設置 show=False 時，繪圖將不會立即顯示在畫面上，你可以在需要顯示的時候調用 plt.show()。
        # 這樣可以讓你在繪製多個圖形時先完成所有的繪製操作再顯示，以提高效率。
        # 顯示座標的顏色
        for i in range(len(x_coords)):
            XYZ = xy_to_XYZ(np.array([x_coords[i], y_coords[i]]))
            RGB = XYZ_to_sRGB(XYZ)
            plot_single_colour_swatch(RGB, swatch_name=f"Color {i + 1}", xy=(x_coords[i], y_coords[i]),
                                      axes=self.CIEsamplechart_R, show=False)

        # 設定框線樣式
        for spine in self.CIEsamplechart_R.spines.values():
            spine.set_visible(True)
            spine.set_color('black')
            spine.set_linewidth(1)

        self.CIEsamplechart_R.axis('off')  # 關閉座標軸

        # # 顯示對應的顏色在 CIE 1931 色度圖上
        # plotting.models.plot_RGB_chromaticities_in_chromaticity_diagram(RGB, axes=self.CIEsamplechart, show=False)
        # 顯示圖例
        self.CIEsamplechart_R.legend()
        self.customcanvas_R.draw()

    def drawcie_G_samplechart(self):
        # 4組 x 和 y 色座標
        x_coords = [self.G_x]
        y_coords = [self.G_y]

        # 轉換為 numpy array
        x_coords_array = np.array(x_coords)
        y_coords_array = np.array(y_coords)

        # 將 xy 座標轉換為 CIE 1931 XYZ 三刺激值
        XYZ = xy_to_XYZ(np.column_stack((x_coords_array, y_coords_array)))

        # 將 XYZ 三刺激值轉換為 sRGB 顏色表示
        RGB = XYZ_to_sRGB(XYZ)

        # 顯示座標的顏色
        # show=False 控制是否顯示即時的繪圖，當你設置 show=False 時，繪圖將不會立即顯示在畫面上，你可以在需要顯示的時候調用 plt.show()。
        # 這樣可以讓你在繪製多個圖形時先完成所有的繪製操作再顯示，以提高效率。
        # 顯示座標的顏色
        for i in range(len(x_coords)):
            XYZ = xy_to_XYZ(np.array([x_coords[i], y_coords[i]]))
            RGB = XYZ_to_sRGB(XYZ)
            plot_single_colour_swatch(RGB, swatch_name=f"Color {i + 1}", xy=(x_coords[i], y_coords[i]),
                                      axes=self.CIEsamplechart_G, show=False)

        # 設定框線樣式
        for spine in self.CIEsamplechart_G.spines.values():
            spine.set_visible(True)
            spine.set_color('black')
            spine.set_linewidth(1)

        self.CIEsamplechart_G.axis('off')  # 關閉座標軸

        # # 顯示對應的顏色在 CIE 1931 色度圖上
        # plotting.models.plot_RGB_chromaticities_in_chromaticity_diagram(RGB, axes=self.CIEsamplechart, show=False)
        # 顯示圖例
        self.CIEsamplechart_G.legend()
        self.customcanvas_G.draw()

    def drawcie_B_samplechart(self):
        # 4組 x 和 y 色座標
        x_coords = [self.B_x]
        y_coords = [self.B_y]

        # 轉換為 numpy array
        x_coords_array = np.array(x_coords)
        y_coords_array = np.array(y_coords)

        # 將 xy 座標轉換為 CIE 1931 XYZ 三刺激值
        XYZ = xy_to_XYZ(np.column_stack((x_coords_array, y_coords_array)))

        # 將 XYZ 三刺激值轉換為 sRGB 顏色表示
        RGB = XYZ_to_sRGB(XYZ)

        # 顯示座標的顏色
        # show=False 控制是否顯示即時的繪圖，當你設置 show=False 時，繪圖將不會立即顯示在畫面上，你可以在需要顯示的時候調用 plt.show()。
        # 這樣可以讓你在繪製多個圖形時先完成所有的繪製操作再顯示，以提高效率。
        # 顯示座標的顏色
        for i in range(len(x_coords)):
            XYZ = xy_to_XYZ(np.array([x_coords[i], y_coords[i]]))
            RGB = XYZ_to_sRGB(XYZ)
            plot_single_colour_swatch(RGB, swatch_name=f"Color {i + 1}", xy=(x_coords[i], y_coords[i]),
                                      axes=self.CIEsamplechart_B, show=False)

        # 設定框線樣式
        for spine in self.CIEsamplechart_B.spines.values():
            spine.set_visible(True)
            spine.set_color('black')
            spine.set_linewidth(1)

        self.CIEsamplechart_B.axis('off')  # 關閉座標軸

        # # 顯示對應的顏色在 CIE 1931 色度圖上
        # plotting.models.plot_RGB_chromaticities_in_chromaticity_diagram(RGB, axes=self.CIEsamplechart, show=False)
        # 顯示圖例
        self.CIEsamplechart_B.legend()
        self.customcanvas_B.draw()

    def drawcie_WRGB_samplechart(self):
        if self.check == "NO":
            self.BLUcheck = "NO"
            print("Please select light source")
            print("self.BLUcheck",self.BLUcheck)
            return
        self.light_source_mode.setEnabled(True)
        self.BLUcheck = "OK"
        # 检查是否已定义 'W_x' 属性
        if not hasattr(self, 'W_x') or self.W_x is None:
            return
        self.drawcie_W_samplechart()
        self.drawcie_R_samplechart()
        self.drawcie_G_samplechart()
        self.drawcie_B_samplechart()

    def picturedialog(self):
        # 檢查是否已經存在對話框
        if hasattr(self, "customdialog") and self.customdialog.isVisible():
            # 如果存在且可見，則關閉舊的對話框
            self.customdialog.close()
        # 創建新的對話框
        self.customdialog = QDialog(None, Qt.Window)  # 這樣才有縮小鍵
        self.customdialog.setWindowTitle("Picture")
        self.customdialog.setStyleSheet("background-color: lightgrey;")
        self.customdialog.resize(1000, 800)
        # Apply styles to the dialog
        self.customdialog.setStyleSheet("""
                    QDialog {
                        background-color: lightyellow;
                    }

                    QTableWidget {
                        background-color: white;
                        alternate-background-color: lightgray;
                        selection-background-color: lightgreen;
                    }

                    QComboBox {
                        background-color: #6495ED;
                        selection-background-color: #4169E1;
                        color: black;  /* Set the text color */
                    }
                    QTableWidget::item:selected {
                color: blcak; /* 設定文字顏色為黑色 */
                background-color: #008080; /* 設定背景顏色為藍色，你可以根據需要調整 */
                    }
                """)

        # dialog layout
        self.customdialog_layout = QGridLayout(self.customdialog)
        # 設定purity_table所在的區域在Grid中的大小
        # self.customdialog_layout.setColumnStretch(1, 2)
        # self.customdialog_layout.setRowStretch(2, 6)
        # 創建參考者選單
        self.observer = QComboBox()
        observer_box_items = ["參考者光源", "A", "B", "C", "D65", "E"]
        for item in observer_box_items:
            self.observer.addItem(str(item))
        self.observer.currentIndexChanged.connect(self.set_wave_p)
        # 創建table
        self.purity_table = QTableWidget()
        self.purity_table.setColumnCount(5)
        self.purity_table.setHorizontalHeaderLabels(["Item", "x", "y", "Wavelength", "Purity"])
        self.purity_table.setRowCount(6)
        # 設定purity_table的最小大小
        self.purity_table.setMinimumSize(550, 150)  # 這裡設定寬度為 300，高度為 200，你可以根據需要調整數值

        # 新增設定顏色的程式碼
        color_items = ["W", "R", "G", "B"]
        color_values = [QColor(RESULTWHITE), QColor(RESULTRED), QColor(RESULTGREEN), QColor(RESULTBLUE)]

        for row, (color_item, color_value) in enumerate(zip(color_items, color_values)):
            item = QTableWidgetItem(color_item)
            item.setBackground(color_value)
            item.setTextAlignment(Qt.AlignCenter)
            self.purity_table.setItem(row, 0, item)

        # 為purity_table添加複製快捷鍵
        copy_purity_shortcut = QShortcut(QKeySequence.Copy, self.purity_table)
        copy_purity_shortcut.activated.connect(self.copy_purity_table_content)


        # label
        font = QFont()
        font.setPointSize(12)  # 設置字型大小
        font.setBold(True)  # 設置為粗體（可選）
        self.label_W = QLabel("W-Sample")
        self.label_W.setFont(font)
        self.label_R = QLabel("R-Sample")
        self.label_R.setFont(font)
        self.label_G = QLabel("G-Sample")
        self.label_G.setFont(font)
        self.label_B = QLabel("B-Sample")
        self.label_B.setFont(font)
        # 圖窗創建
        self.customfigure_CIE = Figure(figsize=(8, 8))
        self.customfigure_W = Figure(figsize=(7, 5))
        self.customfigure_R = Figure(figsize=(7, 5))
        self.customfigure_G = Figure(figsize=(7, 5))
        self.customfigure_B = Figure(figsize=(7, 5))
        # 創立CIE畫布
        self.customcanvas_CIE = FigureCanvas(self.customfigure_CIE)
        self.customcanvas_W = FigureCanvas(self.customfigure_W)
        self.customcanvas_R = FigureCanvas(self.customfigure_R)
        self.customcanvas_G = FigureCanvas(self.customfigure_G)
        self.customcanvas_B = FigureCanvas(self.customfigure_B)
        # 畫布放置 & Lable & table
        self.customdialog_layout.addWidget(self.customcanvas_CIE, 0, 0, 1, 4)
        self.customdialog_layout.addWidget(self.observer, 1, 0, 1, 1)
        self.customdialog_layout.addWidget(self.purity_table, 1, 1, 1, 2)
        self.customdialog_layout.addWidget(self.label_W, 2, 0)
        self.customdialog_layout.addWidget(self.customcanvas_W, 3, 0)
        self.customdialog_layout.addWidget(self.label_R, 2, 1)
        self.customdialog_layout.addWidget(self.customcanvas_R, 3, 1)
        self.customdialog_layout.addWidget(self.label_G, 2, 2)
        self.customdialog_layout.addWidget(self.customcanvas_G, 3, 2)
        self.customdialog_layout.addWidget(self.label_B, 2, 3)
        self.customdialog_layout.addWidget(self.customcanvas_B, 3, 3)
        # self.customcanvas_CIE.figure.clf()
        # 創立建圖物件
        self.CIEchart = self.customcanvas_CIE.figure.add_subplot(111)
        self.CIEsamplechart_W = self.customcanvas_W.figure.add_subplot(111)
        self.CIEsamplechart_R = self.customcanvas_R.figure.add_subplot(111)
        self.CIEsamplechart_G = self.customcanvas_G.figure.add_subplot(111)
        self.CIEsamplechart_B = self.customcanvas_B.figure.add_subplot(111)
        # 顯示新的對話框
        self.customdialog.show()

    # 通用的複製函數
    def copy_table_content(self):
        active_widget = QApplication.focusWidget()
        if isinstance(active_widget, QTableWidget):
            selected_ranges = active_widget.selectedRanges()
            if not selected_ranges:
                return

            copied_data = []
            for selected_range in selected_ranges:
                top_row = selected_range.topRow()
                bottom_row = selected_range.bottomRow() + 1
                left_column = selected_range.leftColumn()
                right_column = selected_range.rightColumn() + 1

                for row in range(top_row, bottom_row):
                    row_data = []
                    for col in range(left_column, right_column):
                        item = active_widget.item(row, col)
                        if item:
                            row_data.append(item.text())
                        else:
                            row_data.append("")  # 如果單元格為空，填充空字符串
                    copied_data.append('\t'.join(row_data))

            copied_text = '\n'.join(copied_data)

            clipboard = QApplication.clipboard()
            clipboard.setText(copied_text)

    # 複製purity_table的內容
    def copy_purity_table_content(self):
        selected_ranges = self.purity_table.selectedRanges()
        if not selected_ranges:
            return

        copied_data = []
        for selected_range in selected_ranges:
            top_row = selected_range.topRow()
            bottom_row = selected_range.bottomRow() + 1
            left_column = selected_range.leftColumn()
            right_column = selected_range.rightColumn() + 1

            for row in range(top_row, bottom_row):
                row_data = []
                for col in range(left_column, right_column):
                    item = self.purity_table.item(row, col)
                    if item:
                        row_data.append(item.text())
                    else:
                        row_data.append("")  # 如果單元格為空，填充空字符串
                copied_data.append('\t'.join(row_data))

        copied_text = '\n'.join(copied_data)

        clipboard = QApplication.clipboard()
        clipboard.setText(copied_text)

    def set_wave_p(self):
        if self.check == "NO":
            self.BLUcheck = "NO"
            print("Please select light source")
            print("self.BLUcheck",self.BLUcheck)
            return
        self.light_source_mode.setEnabled(True)
        self.BLUcheck = "OK"
        # 检查是否已定义 'W_x' 属性
        if not hasattr(self, 'W_x') or self.W_x is None:
            return
        if self.observer.currentText() == "D65":
            self.calculate_W_wave_P(xy_n=[0.3127, 0.329])
            self.calculate_R_wave_P(xy_n=[0.3127, 0.329])
            self.calculate_G_wave_P(xy_n=[0.3127, 0.329])
            self.calculate_B_wave_P(xy_n=[0.3127, 0.329])
        elif self.observer.currentText() == "A":
            self.calculate_W_wave_P(xy_n=[0.448, 0.407])
            self.calculate_R_wave_P(xy_n=[0.448, 0.407])
            self.calculate_G_wave_P(xy_n=[0.448, 0.407])
            self.calculate_B_wave_P(xy_n=[0.448, 0.407])
        elif self.observer.currentText() == "B":
            self.calculate_W_wave_P(xy_n=[0.348, 0.352])
            self.calculate_R_wave_P(xy_n=[0.348, 0.352])
            self.calculate_G_wave_P(xy_n=[0.348, 0.352])
            self.calculate_B_wave_P(xy_n=[0.348, 0.352])
        elif self.observer.currentText() == "C":
            self.calculate_W_wave_P(xy_n=[0.310, 0.316])
            self.calculate_R_wave_P(xy_n=[0.310, 0.316])
            self.calculate_G_wave_P(xy_n=[0.310, 0.316])
            self.calculate_B_wave_P(xy_n=[0.310, 0.316])
        elif self.observer.currentText() == "E":
            self.calculate_W_wave_P(xy_n=[0.333, 0.333])
            self.calculate_R_wave_P(xy_n=[0.333, 0.333])
            self.calculate_G_wave_P(xy_n=[0.333, 0.333])
            self.calculate_B_wave_P(xy_n=[0.333, 0.333])
        item_W_x = self.W_x
        item_W_y = self.W_y
        self.purity_table.setItem(0, 1, QTableWidgetItem(f"{item_W_x:.3f}"))
        self.purity_table.setItem(0, 2, QTableWidgetItem(f"{item_W_y:.3f}"))
        item_R_x = self.R_x
        item_R_y = self.R_y
        self.purity_table.setItem(1, 1, QTableWidgetItem(f"{item_R_x:.3f}"))
        self.purity_table.setItem(1, 2, QTableWidgetItem(f"{item_R_y:.3f}"))
        item_G_x = self.G_x
        item_G_y = self.G_y
        self.purity_table.setItem(2, 1, QTableWidgetItem(f"{item_G_x:.3f}"))
        self.purity_table.setItem(2, 2, QTableWidgetItem(f"{item_G_y:.3f}"))
        item_B_x = self.B_x
        item_B_y = self.B_y
        self.purity_table.setItem(3, 1, QTableWidgetItem(f"{item_B_x:.3f}"))
        self.purity_table.setItem(3, 2, QTableWidgetItem(f"{item_B_y:.3f}"))

    def calculate_W_wave_P(self, xy_n):
        xy = [self.W_x, self.W_y]
        # 計算主波長(
        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
        dominant_wavelength_value = dominant_wavelength
        CIEcoordinate_value_1 = CIEcoordinate[0]
        CIEcoordinate_value_2 = CIEcoordinate[1]
        # 計算距離
        distance_from_white = math.sqrt((self.W_x - xy_n[0]) ** 2 + (self.W_y - xy_n[1]) ** 2)
        distance_from_black = math.sqrt(
            (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
        # 計算色純度
        purity = distance_from_white / distance_from_black * 100
        wave = dominant_wavelength_value
        self.purity_table.setItem(0, 3, QTableWidgetItem(str(wave)))
        self.purity_table.setItem(0, 4, QTableWidgetItem(f"{purity:.2f}"))

    def calculate_R_wave_P(self, xy_n):
        xy = [self.R_x, self.R_y]
        # 計算主波長(
        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
        dominant_wavelength_value = dominant_wavelength
        CIEcoordinate_value_1 = CIEcoordinate[0]
        CIEcoordinate_value_2 = CIEcoordinate[1]
        # 計算距離
        distance_from_white = math.sqrt((self.R_x - xy_n[0]) ** 2 + (self.R_y - xy_n[1]) ** 2)
        distance_from_black = math.sqrt(
            (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
        # 計算色純度
        purity = distance_from_white / distance_from_black * 100
        wave = dominant_wavelength_value
        self.purity_table.setItem(1, 3, QTableWidgetItem(str(wave)))
        self.purity_table.setItem(1, 4, QTableWidgetItem(f"{purity:.2f}"))

    def calculate_G_wave_P(self, xy_n):
        xy = [self.G_x, self.G_y]
        # 計算主波長(
        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
        dominant_wavelength_value = dominant_wavelength
        CIEcoordinate_value_1 = CIEcoordinate[0]
        CIEcoordinate_value_2 = CIEcoordinate[1]
        # 計算距離
        distance_from_white = math.sqrt((self.G_x - xy_n[0]) ** 2 + (self.G_y - xy_n[1]) ** 2)
        distance_from_black = math.sqrt(
            (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
        # 計算色純度
        purity = distance_from_white / distance_from_black * 100
        wave = dominant_wavelength_value
        self.purity_table.setItem(2, 3, QTableWidgetItem(str(wave)))
        self.purity_table.setItem(2, 4, QTableWidgetItem(f"{purity:.2f}"))

    def calculate_B_wave_P(self, xy_n):
        xy = [self.B_x, self.B_y]
        # 計算主波長(
        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
        dominant_wavelength_value = dominant_wavelength
        CIEcoordinate_value_1 = CIEcoordinate[0]
        CIEcoordinate_value_2 = CIEcoordinate[1]
        # 計算距離
        distance_from_white = math.sqrt((self.B_x - xy_n[0]) ** 2 + (self.B_y - xy_n[1]) ** 2)
        distance_from_black = math.sqrt(
            (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
        # 計算色純度
        purity = distance_from_white / distance_from_black * 100
        wave = dominant_wavelength_value
        self.purity_table.setItem(3, 3, QTableWidgetItem(str(wave)))
        self.purity_table.setItem(3, 4, QTableWidgetItem(f"{purity:.2f}"))


    def calculate_R2_wave_P(self):
        xy_n = self.observer_D65
        xy = [self.R_x, self.R_y]
        # 計算主波長(
        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
        dominant_wavelength_value = dominant_wavelength
        CIEcoordinate_value_1 = CIEcoordinate[0]
        CIEcoordinate_value_2 = CIEcoordinate[1]
        # 計算距離
        distance_from_white = math.sqrt((self.R_x - xy_n[0]) ** 2 + (self.R_y - xy_n[1]) ** 2)
        distance_from_black = math.sqrt(
            (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
        # 計算色純度
        purity = distance_from_white / distance_from_black * 100
        wave = dominant_wavelength_value
        self.color_table.setItem(1, 7, QTableWidgetItem(str(wave)))
        self.color_table.setItem(1, 8, QTableWidgetItem(f"{purity:.2f}"))

    def calculate_G2_wave_P(self):
        xy_n = self.observer_D65
        xy = [self.G_x, self.G_y]
        # 計算主波長(
        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
        dominant_wavelength_value = dominant_wavelength
        CIEcoordinate_value_1 = CIEcoordinate[0]
        CIEcoordinate_value_2 = CIEcoordinate[1]
        # 計算距離
        distance_from_white = math.sqrt((self.G_x - xy_n[0]) ** 2 + (self.G_y - xy_n[1]) ** 2)
        distance_from_black = math.sqrt(
            (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
        # 計算色純度
        purity = distance_from_white / distance_from_black * 100
        wave = dominant_wavelength_value
        self.color_table.setItem(1, 12, QTableWidgetItem(str(wave)))
        self.color_table.setItem(1, 13, QTableWidgetItem(f"{purity:.2f}"))

    def calculate_B2_wave_P(self):
        xy_n = self.observer_D65
        xy = [self.B_x, self.B_y]
        # 計算主波長(
        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
        dominant_wavelength_value = dominant_wavelength
        CIEcoordinate_value_1 = CIEcoordinate[0]
        CIEcoordinate_value_2 = CIEcoordinate[1]
        # 計算距離
        distance_from_white = math.sqrt((self.B_x - xy_n[0]) ** 2 + (self.B_y - xy_n[1]) ** 2)
        distance_from_black = math.sqrt(
            (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
        # 計算色純度
        purity = distance_from_white / distance_from_black * 100
        wave = dominant_wavelength_value
        self.color_table.setItem(1, 17, QTableWidgetItem(str(wave)))
        self.color_table.setItem(1, 18, QTableWidgetItem(f"{purity:.2f}"))

    def calculate_All_wave_P(self):
        self.calculate_B2_wave_P()
        self.calculate_R2_wave_P()
        self.calculate_G2_wave_P()

    def calculate_R3_wave_P(self):
        xy_n = self.observer_D65
        xy = [self.RC_x, self.RC_y]
        # 計算主波長(
        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
        dominant_wavelength_value = dominant_wavelength
        CIEcoordinate_value_1 = CIEcoordinate[0]
        CIEcoordinate_value_2 = CIEcoordinate[1]
        # 計算距離
        distance_from_white = math.sqrt((self.RC_x - xy_n[0]) ** 2 + (self.RC_y - xy_n[1]) ** 2)
        distance_from_black = math.sqrt(
            (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
        # 計算色純度
        purity = distance_from_white / distance_from_black * 100
        wave = dominant_wavelength_value
        self.color_table.setItem(2, 7, QTableWidgetItem(str(wave)))
        self.color_table.setItem(2, 8, QTableWidgetItem(f"{purity:.2f}"))

    def calculate_G3_wave_P(self):
        xy_n = self.observer_D65
        xy = [self.GC_x, self.GC_y]
        # 計算主波長(
        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
        dominant_wavelength_value = dominant_wavelength
        CIEcoordinate_value_1 = CIEcoordinate[0]
        CIEcoordinate_value_2 = CIEcoordinate[1]
        # 計算距離
        distance_from_white = math.sqrt((self.GC_x - xy_n[0]) ** 2 + (self.GC_y - xy_n[1]) ** 2)
        distance_from_black = math.sqrt(
            (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
        # 計算色純度
        purity = distance_from_white / distance_from_black * 100
        wave = dominant_wavelength_value
        self.color_table.setItem(2, 12, QTableWidgetItem(str(wave)))
        self.color_table.setItem(2, 13, QTableWidgetItem(f"{purity:.2f}"))

    def calculate_B3_wave_P(self):
        xy_n = self.observer_D65
        xy = [self.BC_x, self.BC_y]
        # 計算主波長(
        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
        dominant_wavelength_value = dominant_wavelength
        CIEcoordinate_value_1 = CIEcoordinate[0]
        CIEcoordinate_value_2 = CIEcoordinate[1]
        # 計算距離
        distance_from_white = math.sqrt((self.BC_x - xy_n[0]) ** 2 + (self.BC_y - xy_n[1]) ** 2)
        distance_from_black = math.sqrt(
            (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
        # 計算色純度
        purity = distance_from_white / distance_from_black * 100
        wave = dominant_wavelength_value
        self.color_table.setItem(2, 17, QTableWidgetItem(str(wave)))
        self.color_table.setItem(2, 18, QTableWidgetItem(f"{purity:.2f}"))

    def calculate_All_2_wave_P(self):
        self.calculate_B3_wave_P()
        self.calculate_R3_wave_P()
        self.calculate_G3_wave_P()

    # Plot CIE jump
    def drawcie(self):
        # 4组 x 和 y 色坐标
        x_coords = [self.W_x, self.R_x, self.G_x, self.B_x]
        y_coords = [self.W_y, self.R_y, self.G_y, self.B_y]

        # 将 xy 坐标转换为 CIE 1931 XYZ 三刺激值，并将它们转换为 sRGB 值
        RGB_values = []
        for x, y in zip(x_coords, y_coords):
            XYZ = xy_to_XYZ(np.array([x, y]))
            RGB = XYZ_to_sRGB(XYZ)
            RGB_values.append(RGB)

        # 转换为 numpy 数组
        RGB_array = np.array(RGB_values)

        # 显示座标的颜色
        for i, (RGB, x, y) in enumerate(zip(RGB_array, x_coords, y_coords)):
            plot_single_colour_swatch(RGB, swatch_name=f"Color {i + 1}", xy=(x, y), show=False)

        # 显示对应的颜色在 CIE 1931 色度图上
        # 注意：这里可能需要调整以适应您的具体需求
        plotting.plot_chromaticity_diagram_CIE1931(show=False)

        # 使用 plt.scatter 添加自定义的标记
        markers = ['o', 's', '^', '*']  # 可以根据需要更改标记形状
        labels = ['W', 'R', 'G', 'B']
        edgecolors = 'black'
        for x, y, marker, label, RGB in zip(x_coords, y_coords, markers, labels, RGB_array):
            # 确保颜色值在 0 到 1 的范围内
            RGB_clipped = np.clip(RGB, 0, 1)
            plt.scatter(x, y, marker=marker, label=f'{label}:({x:.2f}, {y:.2f})', c=[RGB_clipped],
                        edgecolors=edgecolors)

        # 显示图例
        plt.legend()
        plt.show()  # 显示最后一次更新的图形

    # table更新function區--------------------------------------------------------------------------------------

    def update_light_source_datatable(self):
        # 防止更新不停Reload

        self.check = "NO"

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
        self.check = "OK"
        self.BLUcheck = "NO"





    def updateLightSourceComboBox(self):
        # 防止更新不停Reload

        self.check = "NO"

        connection = sqlite3.connect("blu_database.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.light_source_datatable.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.light_source.clear()
        for item in header_labels:
            self.light_source.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"




    def update_light_source_led_datatable(self):
        # 防止更新不停Reload

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
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
        self.check = "OK"
        self.BLUcheck = "NO"


    def updateLightSourceledComboBox(self):
        self.check = "NO"
        # 防止更新不停Reload

        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("led_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        # cursor.execute(f"PRAGMA table_info({self.light_source_led_datatable.currentText()});")
        cursor.execute(f"PRAGMA table_info('{self.light_source_led_datatable.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.light_source_led.clear()
        for item in header_labels:
            self.light_source_led.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def update_source_led_datatable(self):
        # 防止更新不停Reload

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
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
        self.check = "OK"
        self.BLUcheck = "NO"


    def updateSourceledComboBox(self):
        # 防止更新不停Reload

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("led_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.source_led_datatable.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.source_led.clear()
        for item in header_labels:
            self.source_led.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"



    # cell
    def update_layer1_datatable(self):
        # 防止更新不停Reload

        self.check = "NO"
        print("self.check",self.check)
        # # 防呆反灰設定
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
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
        self.check = "OK"
        self.BLUcheck = "NO"



    def updatelayer1ComboBox(self):
        # 防止更新不停Reload

        self.check = "NO"
        # # 防呆反灰設定
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("cell_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.layer1_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.layer1_box.clear()
        for item in header_labels:
            self.layer1_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    # BSITO
    def update_layer2_datatable(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source_mode.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
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
        self.check = "OK"
        self.BLUcheck = "NO"


    def updatelayer2ComboBox(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source_mode.setEnabled(False)
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("BSITO_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.layer2_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.layer2_box.clear()
        for item in header_labels:
            self.layer2_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def update_layer3_datatable(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("Layer3_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.layer3_table.clear()
        for table in tables:
            self.layer3_table.addItem(table[0])

        # 關閉連線
        conn.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def updatelayer3ComboBox(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("Layer3_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.layer3_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.layer3_box.clear()
        for item in header_labels:
            self.layer3_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def update_layer4_datatable(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("Layer4_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.layer4_table.clear()
        for table in tables:
            self.layer4_table.addItem(table[0])

        # 關閉連線
        conn.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def updatelayer4ComboBox(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("Layer4_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.layer4_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.layer4_box.clear()
        for item in header_labels:
            self.layer4_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def update_layer5_datatable(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("Layer5_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.layer5_table.clear()
        for table in tables:
            self.layer5_table.addItem(table[0])

        # 關閉連線
        conn.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def updatelayer5ComboBox(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("Layer5_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.layer5_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.layer5_box.clear()
        for item in header_labels:
            self.layer5_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def update_layer6_datatable(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("Layer6_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.layer6_table.clear()
        for table in tables:
            self.layer6_table.addItem(table[0])

        # 關閉連線
        conn.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def updatelayer6ComboBox(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("Layer5_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.layer6_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.layer6_box.clear()
        for item in header_labels:
            self.layer6_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    # RCF_Fix
    def update_RCF_Fix_datatable(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
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
        self.check = "OK"
        self.BLUcheck = "NO"


    def updateRCF_Fix_ComboBox(self):

        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("RCF_Fix_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.R_fix_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.R_fix_box.clear()
        for item in header_labels:
            self.R_fix_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def update_RCF_Fix_modeclose(self):


        if self.R_fix_mode.currentText() == "自訂" or self.R_fix_mode.currentText() == "模擬":
            self.R_aK_mode.setCurrentText("未選")
            self.R_differ_mode.setCurrentText("未選")
        else:
            # 在Color_Table上顯示
            self.color_table.setItem(4, 1, QTableWidgetItem(""))
            self.color_table.setItem(4, 2, QTableWidgetItem(""))

    # GCF_Fix
    def update_GCF_Fix_datatable(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
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
        self.check = "OK"
        self.BLUcheck = "NO"


    def updateGCF_Fix_ComboBox(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
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
        self.check = "OK"
        self.BLUcheck = "NO"


    def update_GCF_Fix_modeclose(self):

        if self.G_fix_mode.currentText() == "自訂" or self.G_fix_mode.currentText() == "模擬":
            self.G_aK_mode.setCurrentText("未選")
            self.G_differ_mode.setCurrentText("未選")
        else:
            # 在Color_Table上顯示
            self.color_table.setItem(4, 3, QTableWidgetItem(""))
            self.color_table.setItem(4, 4, QTableWidgetItem(""))

    # BCF_Fix
    def update_BCF_Fix_datatable(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
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
        self.check = "OK"
        self.BLUcheck = "NO"


    def updateBCF_Fix_ComboBox(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
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
        self.check = "OK"
        self.BLUcheck = "NO"


    def update_BCF_Fix_modeclose(self):


        if self.B_fix_mode.currentText() == "自訂" or self.B_fix_mode.currentText() == "模擬":
            self.B_aK_mode.setCurrentText("未選")
            self.B_differ_mode.setCurrentText("未選")
        else:
            # 在Color_Table上顯示
            self.color_table.setItem(4, 5, QTableWidgetItem(""))
            self.color_table.setItem(4, 6, QTableWidgetItem(""))

    # RCF_Change
    def update_RCF_Change_datatable(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("RCF_Change_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.R_aK_table.clear()
        for table in tables:
            self.R_aK_table.addItem(table[0])

        # 關閉連線
        conn.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def updateRCF_Change_ComboBox(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("RCF_Change_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.R_aK_table.currentText()}');")
        header_data = cursor.fetchall()
        print("header_data",header_data)
        header_labels = [column[1] for column in header_data]
        print("headerlabels-from source", header_labels)
        self.R_aK_box.clear()
        added_items = set()  # 使用集合來追踪已添加的項目
        # 只添加尚未添加的項目
        for index, column in enumerate(header_data):
            header_name = column[1]
            print("header_name", header_name)

            # 尋找最後一個 "_" 的位置
            underscore_index = header_name.rfind('_')
            # 如果存在 "_" 且不在開頭，則截取到這個位置
            if underscore_index > 0:
                header_name = header_name[:underscore_index]

            if header_name not in added_items:
                self.R_aK_box.addItem(str(header_name))
                added_items.add(header_name)
                print("item", str(header_name))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"


    def update_RCF_Change_modeclose(self):
        if self.R_aK_mode.currentText() == "自訂" or self.R_aK_mode.currentText() == "模擬":
            self.R_fix_mode.setCurrentText("未選")
            self.R_differ_mode.setCurrentText("未選")
        else:
            # 在Color_Table上顯示
            self.color_table.setItem(4, 1, QTableWidgetItem(""))
            self.color_table.setItem(4, 2, QTableWidgetItem(""))

    # GCF_Change
    def update_GCF_Change_datatable(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("GCF_Change_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.G_aK_table.clear()
        for table in tables:
            self.G_aK_table.addItem(table[0])

        # 關閉連線
        conn.close()
        self.check = "OK"
        self.BLUcheck = "NO"

    def updateGCF_Change_ComboBox(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("GCF_Change_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.G_aK_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.G_aK_box.clear()
        added_items = set()  # 使用集合來追踪已添加的項目
        # 只添加尚未添加的項目
        for index, column in enumerate(header_data):
            header_name = column[1]
            print("header_name", header_name)

            # 尋找最後一個 "_" 的位置
            underscore_index = header_name.rfind('_')
            # 如果存在 "_" 且不在開頭，則截取到這個位置
            if underscore_index > 0:
                header_name = header_name[:underscore_index]

            if header_name not in added_items:
                self.G_aK_box.addItem(str(header_name))
                added_items.add(header_name)
                print("item", str(header_name))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"

    def update_GCF_Change_modeclose(self):
        if self.G_aK_mode.currentText() == "自訂" or self.G_aK_mode.currentText() == "模擬":
            self.G_fix_mode.setCurrentText("未選")
            self.G_differ_mode.setCurrentText("未選")
        else:
            # 在Color_Table上顯示
            self.color_table.setItem(4, 3, QTableWidgetItem(""))
            self.color_table.setItem(4, 4, QTableWidgetItem(""))

    # BCF_Change
    def update_BCF_Change_datatable(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("BCF_Change_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.B_aK_table.clear()
        for table in tables:
            self.B_aK_table.addItem(table[0])

        # 關閉連線
        conn.close()
        self.check = "OK"
        self.BLUcheck = "NO"

    def updateBCF_Change_ComboBox(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("BCF_Change_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.B_aK_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.B_aK_box.clear()
        added_items = set()  # 使用集合來追踪已添加的項目
        # 只添加尚未添加的項目
        for index, column in enumerate(header_data):
            header_name = column[1]
            print("header_name", header_name)

            # 尋找最後一個 "_" 的位置
            underscore_index = header_name.rfind('_')
            # 如果存在 "_" 且不在開頭，則截取到這個位置
            if underscore_index > 0:
                header_name = header_name[:underscore_index]

            if header_name not in added_items:
                self.B_aK_box.addItem(str(header_name))
                added_items.add(header_name)
                print("item", str(header_name))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"

    def update_BCF_Change_modeclose(self):
        if self.B_aK_mode.currentText() == "自訂" or self.B_aK_mode.currentText() == "模擬":
            self.B_fix_mode.setCurrentText("未選")
            self.B_differ_mode.setCurrentText("未選")
        else:
            # 在Color_Table上顯示
            self.color_table.setItem(4, 5, QTableWidgetItem(""))
            self.color_table.setItem(4, 6, QTableWidgetItem(""))
    # RCF_Differ
    def update_RCF_Differ_datatable(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("RCF_Differ_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.R_differ_table.clear()
        for table in tables:
            self.R_differ_table.addItem(table[0])

        # 關閉連線
        conn.close()
        self.check = "OK"
        self.BLUcheck = "NO"

    def updateRCF_Differ_ComboBox(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("RCF_Differ_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.R_differ_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.R_differ_box.clear()
        added_items = set()  # 使用集合來追踪已添加的項目
        # 只添加尚未添加的項目
        for index, column in enumerate(header_data):
            header_name = column[1]
            print("header_name", header_name)

            # 尋找最後一個 "_" 的位置
            underscore_index = header_name.rfind('_')
            # 如果存在 "_" 且不在開頭，則截取到這個位置
            if underscore_index > 0:
                header_name = header_name[:underscore_index]

            if header_name not in added_items:
                self.R_differ_box.addItem(str(header_name))
                added_items.add(header_name)
                print("item", str(header_name))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"

    def update_RCF_Differ_modeclose(self):
        if self.R_differ_mode.currentText() == "自訂" or self.R_differ_mode.currentText() == "模擬":
            self.R_fix_mode.setCurrentText("未選")
            self.R_aK_mode.setCurrentText("未選")
        else:
            # 在Color_Table上顯示
            self.color_table.setItem(4, 1, QTableWidgetItem(""))
            self.color_table.setItem(4, 2, QTableWidgetItem(""))

    # GCF_Differ
    def update_GCF_Differ_datatable(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("GCF_Differ_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.G_differ_table.clear()
        for table in tables:
            self.G_differ_table.addItem(table[0])

        # 關閉連線
        conn.close()
        self.check = "OK"
        self.BLUcheck = "NO"

    def updateGCF_Differ_ComboBox(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("GCF_Differ_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.G_differ_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        added_items = set()  # 使用集合來追踪已添加的項目
        self.G_differ_box.clear()
        # 只添加尚未添加的項目
        for index, column in enumerate(header_data):
            header_name = column[1]
            print("header_name", header_name)

            # 尋找最後一個 "_" 的位置
            underscore_index = header_name.rfind('_')
            # 如果存在 "_" 且不在開頭，則截取到這個位置
            if underscore_index > 0:
                header_name = header_name[:underscore_index]
            if header_name not in added_items:
                self.G_differ_box.addItem(str(header_name))
                added_items.add(header_name)
                print("item", str(header_name))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"

    def update_GCF_Differ_modeclose(self):
        if self.G_differ_mode.currentText() == "自訂":
            self.G_fix_mode.setCurrentText("未選")
            self.G_aK_mode.setCurrentText("未選")
        else:
            # 在Color_Table上顯示
            self.color_table.setItem(4, 3, QTableWidgetItem(""))
            self.color_table.setItem(4, 4, QTableWidgetItem(""))

    # BCF_Differ
    def update_BCF_Differ_datatable(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("BCF_Differ_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.B_differ_table.clear()
        for table in tables:
            self.B_differ_table.addItem(table[0])

        # 關閉連線
        conn.close()
        self.check = "OK"
        self.BLUcheck = "NO"

    def updateBCF_Differ_ComboBox(self):
        # self.light_source_mode.setCurrentText("未選")
        # self.light_source_mode.setEnabled(False)
        self.check = "NO"
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("BCF_Differ_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.B_differ_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)

        added_items = set()  # 使用集合來追踪已添加的項目
        self.B_differ_box.clear()
        # 只添加尚未添加的項目
        for index, column in enumerate(header_data):
            header_name = column[1]
            print("header_name", header_name)

            # 尋找最後一個 "_" 的位置
            underscore_index = header_name.rfind('_')
            # 如果存在 "_" 且不在開頭，則截取到這個位置
            if underscore_index > 0:
                header_name = header_name[:underscore_index]

            if header_name not in added_items:
                self.B_differ_box.addItem(str(header_name))
                added_items.add(header_name)
                print("item", str(header_name))
        # 關閉連線
        connection.close()
        self.check = "OK"
        self.BLUcheck = "NO"

    def update_BCF_Differ_modeclose(self):
        if self.B_differ_mode.currentText() == "自訂":
            self.B_fix_mode.setCurrentText("未選")
            self.B_aK_mode.setCurrentText("未選")
        else:
            # 在Color_Table上顯示
            self.color_table.setItem(4, 5, QTableWidgetItem(""))
            self.color_table.setItem(4, 6, QTableWidgetItem(""))

#-------------TK區------------------------------------------------------------------------------------------
    def update_TK_light_source_datatable(self):

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
        self.TK_light_source_datatable.clear()
        for table in tables:
            self.TK_light_source_datatable.addItem(table[0])

        # 關閉連線
        conn.close()

    def updateTK_light_sourceComboBox(self):

        connection = sqlite3.connect("blu_database.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.TK_light_source_datatable.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.TK_light_source.clear()
        for item in header_labels:
            self.TK_light_source.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_TK_layer1_datatable(self):
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
        self.TK_layer1_table.clear()
        for table in tables:
            self.TK_layer1_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def updateTKlayer1ComboBox(self):
        # # 防呆反灰設定
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("cell_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.TK_layer1_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.TK_layer1_box.clear()
        for item in header_labels:
            self.TK_layer1_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_TK_layer2_datatable(self):
        # 防止更新不停Reload
        self.light_source.setEnabled(False)
        self.light_source_datatable.setEnabled(False)
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
        self.TK_layer2_table.clear()
        for table in tables:
            self.TK_layer2_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def updateTKlayer2ComboBox(self):
        # 防止更新不停Reload
        self.light_source.setEnabled(False)
        self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("BSITO_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.TK_layer2_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.TK_layer2_box.clear()
        for item in header_labels:
            self.TK_layer2_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_TK_layer3_datatable(self):
        # 防止更新不停Reload
        self.light_source.setEnabled(False)
        self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("Layer3_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.TK_layer3_table.clear()
        for table in tables:
            self.TK_layer3_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def updateTKlayer3ComboBox(self):
        # 防止更新不停Reload
        self.light_source.setEnabled(False)
        self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("Layer3_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.TK_layer3_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.TK_layer3_box.clear()
        for item in header_labels:
            self.TK_layer3_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_TK_layer4_datatable(self):
        # 防止更新不停Reload
        self.light_source.setEnabled(False)
        self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("Layer4_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.TK_layer4_table.clear()
        for table in tables:
            self.TK_layer4_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def updateTKlayer4ComboBox(self):
        # 防止更新不停Reload
        self.light_source.setEnabled(False)
        self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("Layer4_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.TK_layer4_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.TK_layer4_box.clear()
        for item in header_labels:
            self.TK_layer4_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_TK_layer5_datatable(self):
        # 防止更新不停Reload
        self.light_source.setEnabled(False)
        self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("Layer5_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.TK_layer5_table.clear()
        for table in tables:
            self.TK_layer5_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def updateTKlayer5ComboBox(self):
        # 防止更新不停Reload
        self.light_source.setEnabled(False)
        self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("Layer5_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.TK_layer5_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.TK_layer5_box.clear()
        for item in header_labels:
            self.TK_layer5_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_TK_layer6_datatable(self):
        # 防止更新不停Reload
        self.light_source.setEnabled(False)
        self.light_source_datatable.setEnabled(False)
        # 更新 ComboBox 的選項
        # 在需要更新 ComboBox 的地方呼叫這個函數
        # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
        # 連接到 SQLite 資料庫
        conn = sqlite3.connect("Layer6_spectrum.db")
        cursor = conn.cursor()

        # 取得所有的 table 名稱
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()

        # 更新 ComboBox 的選項
        self.TK_layer6_table.clear()
        for table in tables:
            self.TK_layer6_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def updateTKlayer6ComboBox(self):
        # 防止更新不停Reload
        self.light_source.setEnabled(False)
        self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("Layer5_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.TK_layer6_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.TK_layer6_box.clear()
        for item in header_labels:
            self.TK_layer6_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    # TK RCF
    def update_TK_RCF_datatable(self):
        if self.TK_RCF_mode_option.currentText() == "RCF_Change":
            # 更新 ComboBox 的選項
            # 在需要更新 ComboBox 的地方呼叫這個函數
            # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
            # 連接到 SQLite 資料庫
            conn = sqlite3.connect("RCF_Change_spectrum.db")
            cursor = conn.cursor()

            # 取得所有的 table 名稱
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()

            # 更新 ComboBox 的選項
            self.TK_RCF_table.clear()
            for table in tables:
                self.TK_RCF_table.addItem(table[0])

            # 關閉連線
            conn.close()
        else:
            # 更新 ComboBox 的選項
            # 在需要更新 ComboBox 的地方呼叫這個函數
            # 例如，當你新增了新的 table 時，呼叫 updateTableComboBox() 以更新 ComboBox
            # 連接到 SQLite 資料庫
            conn = sqlite3.connect("RCF_Differ_spectrum.db")
            cursor = conn.cursor()

            # 取得所有的 table 名稱
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
            tables = cursor.fetchall()

            # 更新 ComboBox 的選項
            self.TK_RCF_table.clear()
            for table in tables:
                self.TK_RCF_table.addItem(table[0])

            # 關閉連線
            conn.close()

    def update_TK_RCF_ComboBox(self):
        if self.TK_RCF_mode_option.currentText() == "RCF_Change":
            connection = sqlite3.connect("RCF_Change_spectrum.db")
            cursor = connection.cursor()

            # 获取表格的標題
            cursor.execute(f"PRAGMA table_info('{self.TK_RCF_table.currentText()}');")
            header_data = cursor.fetchall()
            header_labels = [column[1] for column in header_data]
            # print("headerlabels-from source", header_labels)
            self.TK_RCF_Combobox.clear()
            added_items = set()  # 使用集合來追踪已添加的項目
            # 只添加尚未添加的項目
            for index, column in enumerate(header_data):
                header_name = column[1]
                # 檢查是否存在 "_" 並且 "_" 不在開頭
                if '_' in header_name and not header_name.startswith('_'):
                    header_name = header_name[:-2]

                if header_name not in added_items:
                    self.TK_RCF_Combobox.addItem(str(header_name))
                    added_items.add(header_name)
                    print("item", str(header_name))
            # 關閉連線
            connection.close()
        else:
            connection = sqlite3.connect("RCF_Differ_spectrum.db")
            cursor = connection.cursor()

            # 获取表格的標題
            cursor.execute(f"PRAGMA table_info('{self.TK_RCF_table.currentText()}');")
            header_data = cursor.fetchall()
            header_labels = [column[1] for column in header_data]
            # print("headerlabels-from source", header_labels)
            self.TK_RCF_Combobox.clear()
            added_items = set()  # 使用集合來追踪已添加的項目
            # 只添加尚未添加的項目
            for index, column in enumerate(header_data):
                header_name = column[1]
                # 檢查是否存在 "_" 並且 "_" 不在開頭
                if '_' in header_name and not header_name.startswith('_'):
                    header_name = header_name[:-2]

                if header_name not in added_items:
                    self.TK_RCF_Combobox.addItem(str(header_name))
                    added_items.add(header_name)
                    print("item", str(header_name))
            # 關閉連線
            connection.close()

    def calculate_TK_BLU(self):

        connection_BLU = sqlite3.connect("blu_database.db")
        cursor_BLU = connection_BLU.cursor()
        # 取得BLU資料
        column_name_BLU = self.TK_light_source.currentText()
        table_name_BLU = self.TK_light_source_datatable.currentText()
        # 使用正確的引號包裹表名和列名
        query_BLU = f"SELECT * FROM '{table_name_BLU}';"
        cursor_BLU.execute(query_BLU)
        result_BLU = cursor_BLU.fetchall()
        # 找到指定標題的欄位索引
        header_BLU = [column[0] for column in cursor_BLU.description]
        print("header_BLU", header_BLU)
        column_index_BLU = header_BLU.index(f"{column_name_BLU}")
        # 取得指定欄位的數據並轉換為 Series
        BLU_spectrum_Series = pd.Series([row[column_index_BLU] for row in result_BLU])
        print("column_index_BLU", column_index_BLU)
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

    def calculate_TK_layer1(self):
        if self.TK_layer1_mode.currentText() == "Layer1_自訂":
            connection_layer1 = sqlite3.connect("cell_spectrum.db")
            cursor_layer1 = connection_layer1.cursor()
            # 取得BLU資料
            column_name_layer1 = self.TK_layer1_box.currentText()
            table_name_layer1 = self.TK_layer1_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer1 = f"SELECT * FROM '{table_name_layer1}';"
            cursor_layer1.execute(query_layer1)
            result_layer1 = cursor_layer1.fetchall()

            # 找到指定標題的欄位索引
            header_layer1 = [column[0] for column in cursor_layer1.description]
            column_index_layer1 = header_layer1.index(f"{column_name_layer1}")

            # 取得指定欄位的數據並轉換為 Series
            layer1_spectrum_Series = pd.Series([row[column_index_layer1] for row in result_layer1])

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
            return layer1_spectrum_Series

    def calculate_TK_layer2(self):
        if self.TK_layer2_mode.currentText() == "Layer2_自訂":
            connection_layer2 = sqlite3.connect("BSITO_spectrum.db")
            cursor_layer2 = connection_layer2.cursor()
            # 取得BLU資料
            column_name_layer2 = self.TK_layer2_box.currentText()
            table_name_layer2 = self.TK_layer2_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer2 = f"SELECT * FROM '{table_name_layer2}';"
            cursor_layer2.execute(query_layer2)
            result_layer2 = cursor_layer2.fetchall()

            # 找到指定標題的欄位索引
            header_layer2 = [column[0] for column in cursor_layer2.description]
            column_index_layer2 = header_layer2.index(f"{column_name_layer2}")

            # 取得指定欄位的數據並轉換為 Series
            layer2_spectrum_Series = pd.Series([row[column_index_layer2] for row in result_layer2])
            # print("layer2_spectrum_Series", layer2_spectrum_Series)

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

    def calculate_TK_layer3(self):
        if self.TK_layer3_mode.currentText() == "Layer3_自訂":
            connection_layer3 = sqlite3.connect("Layer3_spectrum.db")
            cursor_layer3 = connection_layer3.cursor()
            # 取得BLU資料
            column_name_layer3 = self.TK_layer3_box.currentText()
            table_name_layer3 = self.TK_layer3_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer3 = f"SELECT * FROM '{table_name_layer3}';"
            cursor_layer3.execute(query_layer3)
            result_layer3 = cursor_layer3.fetchall()

            # 找到指定標題的欄位索引
            header_layer3 = [column[0] for column in cursor_layer3.description]
            column_index_layer3 = header_layer3.index(f"{column_name_layer3}")

            # 取得指定欄位的數據並轉換為 Series
            layer3_spectrum_Series = pd.Series([row[column_index_layer3] for row in result_layer3])
            # print("layer2_spectrum_Series", layer2_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer3_spectrum_Series = pd.to_numeric(layer3_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer3_spectrum_Series = layer3_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer3_spectrum_Series = layer3_spectrum_Series.dropna()

            # 關閉連線
            connection_layer3.close()
            # 自訂BLU_Spectrum回傳
            return layer3_spectrum_Series
        else:
            layer3_spectrum_Series = 1
            print("layer3:未選")
            return layer3_spectrum_Series

    def calculate_TK_layer4(self):
        if self.TK_layer4_mode.currentText() == "Layer4_自訂":
            print("in_layer4_自訂")
            connection_layer4 = sqlite3.connect("Layer4_spectrum.db")
            cursor_layer4 = connection_layer4.cursor()
            # 取得BLU資料
            column_name_layer4 = self.TK_layer4_box.currentText()
            table_name_layer4 = self.TK_layer4_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer4 = f"SELECT * FROM '{table_name_layer4}';"
            cursor_layer4.execute(query_layer4)
            result_layer4 = cursor_layer4.fetchall()

            # 找到指定標題的欄位索引
            header_layer4 = [column[0] for column in cursor_layer4.description]
            column_index_layer4 = header_layer4.index(f"{column_name_layer4}")

            # 取得指定欄位的數據並轉換為 Series
            layer4_spectrum_Series = pd.Series([row[column_index_layer4] for row in result_layer4])
            # print("layer4_spectrum_Series", layer4_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer4_spectrum_Series = pd.to_numeric(layer4_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer4_spectrum_Series = layer4_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer4_spectrum_Series = layer4_spectrum_Series.dropna()

            # 關閉連線
            connection_layer4.close()
            # 自訂BLU_Spectrum回傳
            return layer4_spectrum_Series
        else:
            layer4_spectrum_Series = 1
            print("layer4:未選")
            return layer4_spectrum_Series

    def calculate_TK_layer5(self):
        if self.TK_layer5_mode.currentText() == "Layer5_自訂":
            print("in_layer5_自訂")
            connection_layer5 = sqlite3.connect("Layer5_spectrum.db")
            cursor_layer5 = connection_layer5.cursor()
            # 取得BLU資料
            column_name_layer5 = self.TK_layer5_box.currentText()
            table_name_layer5 = self.TK_layer5_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer5 = f"SELECT * FROM '{table_name_layer5}';"
            cursor_layer5.execute(query_layer5)
            result_layer5 = cursor_layer5.fetchall()

            # 找到指定標題的欄位索引
            header_layer5 = [column[0] for column in cursor_layer5.description]
            column_index_layer5 = header_layer5.index(f"{column_name_layer5}")

            # 取得指定欄位的數據並轉換為 Series
            layer5_spectrum_Series = pd.Series([row[column_index_layer5] for row in result_layer5])
            # print("layer5_spectrum_Series", layer5_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer5_spectrum_Series = pd.to_numeric(layer5_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer5_spectrum_Series = layer5_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer5_spectrum_Series = layer5_spectrum_Series.dropna()

            # 關閉連線
            connection_layer5.close()
            # 自訂BLU_Spectrum回傳
            return layer5_spectrum_Series
        else:
            layer5_spectrum_Series = 1
            print("layer5:未選")
            return layer5_spectrum_Series

    def calculate_TK_layer6(self):
        if self.TK_layer6_mode.currentText() == "自訂":
            print("in_layer6_自訂")
            connection_layer6 = sqlite3.connect("Layer6_spectrum.db")
            cursor_layer6 = connection_layer6.cursor()
            # 取得BLU資料
            column_name_layer6 = self.TK_layer6_box.currentText()
            table_name_layer6 = self.TK_layer6_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_layer6 = f"SELECT * FROM '{table_name_layer6}';"
            cursor_layer6.execute(query_layer6)
            result_layer6 = cursor_layer6.fetchall()

            # 找到指定標題的欄位索引
            header_layer6 = [column[0] for column in cursor_layer6.description]
            column_index_layer6 = header_layer6.index(f"{column_name_layer6}")

            # 取得指定欄位的數據並轉換為 Series
            layer6_spectrum_Series = pd.Series([row[column_index_layer6] for row in result_layer6])
            # print("layer6_spectrum_Series", layer6_spectrum_Series)

            # 將 Series 中的字符串轉換為數值
            layer6_spectrum_Series = pd.to_numeric(layer6_spectrum_Series, errors='coerce')

            # 將 NaN 值替換為 0，或者根據實際需求替換為其他值
            layer6_spectrum_Series = layer6_spectrum_Series.fillna(0)

            # 刪除包含空值的行
            layer6_spectrum_Series = layer6_spectrum_Series.dropna()

            # 關閉連線
            connection_layer6.close()
            # 自訂BLU_Spectrum回傳
            return layer6_spectrum_Series
        else:
            layer6_spectrum_Series = 1
            print("layer6:未選")
            return layer6_spectrum_Series

    def calculate_TK_RCF(self):
        if self.TK_RCF_mode_option.currentText() == "RCF_Change":
            print("in_RCF_Change自訂")
            connection_RCF_Change = sqlite3.connect("RCF_Change_spectrum.db")
            cursor_RCF_Change = connection_RCF_Change.cursor()
            # 取得Rdiffer資料
            column_name_RCF_Change = self.TK_RCF_Combobox.currentText()
            print("column_name_RCF_Change",column_name_RCF_Change)
            table_name_RCF_Change = self.TK_RCF_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_RCF_Change = f"SELECT * FROM '{table_name_RCF_Change}';"
            cursor_RCF_Change.execute(query_RCF_Change)
            result_RCF_Change = cursor_RCF_Change.fetchall()

            # 找到指定標題的欄位索引
            header_RCF_Change = [column[0] for column in cursor_RCF_Change.description]

            # 移除所有標題中的編號
            header_RCF_Change = [re.sub(r'_\d+$', '', col) for col in header_RCF_Change]
            print("header_RCF_Change", header_RCF_Change)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_RCF_Change) if col == column_name_RCF_Change]
            print("indices_without_number", indices_without_number)

            # 使用 pd.to_numeric 转换为数值类型，并处理无法转换的值
            SP_series_list = [
                pd.to_numeric(pd.Series([row[i] for row in result_RCF_Change]), errors='coerce').fillna(0)
                for i in indices_without_number
            ]
            print("SP_series_list",SP_series_list)

            # TK_SP準備分厚度順序
            TK_SP_list = []
            for i in range(len(indices_without_number)):
                TK_SP_list.append(SP_series_list[i].iloc[0])
            print("TK_SP_list",TK_SP_list)
            Sort_TK_SP_list = sorted(TK_SP_list,reverse=True)
            print("Sort_TK_SP_list",Sort_TK_SP_list)
            self.TK_RCF_Range.setText(f"TKrange: {min(TK_SP_list):.2f}~{max(TK_SP_list):.2f}")

            # 反轉找回原本的順序list
            Re_list = []
            for i in range(len(indices_without_number)):
                Re_list.append(TK_SP_list.index(Sort_TK_SP_list[i]))
            #print("Re_list",Re_list)

            A_list = []
            for i in range(len(indices_without_number)-1):
                A_list.append((-1/(float(Sort_TK_SP_list[i+1]) - float(Sort_TK_SP_list[i])) * np.log(SP_series_list[Re_list[i+1]][1:]/SP_series_list[Re_list[i]][1:])).reset_index(drop=True))
            #print("A_list",A_list)
            K_list = []
            for i in range(len(indices_without_number)-1):
                K_list.append(SP_series_list[Re_list[i]][1:].reset_index(drop=True)/np.exp((-1 * A_list[i])* Sort_TK_SP_list[i]))
            #print("K_list",K_list)
            self.RCF_TK_list = []
            if float(self.TK_RCF_Strat.text()) < float(self.TK_RCF_End.text()):
                i = 0
                while float(self.TK_RCF_Strat.text()) + float(self.TK_RCF_interval.text()) * i < float(self.TK_RCF_End.text()):
                    self.RCF_TK_list.append(float(self.TK_RCF_Strat.text()) + float(self.TK_RCF_interval.text()) * i)
                    i += 1
                self.RCF_TK_list.append(float(self.TK_RCF_End.text()))

            elif float(self.TK_RCF_End.text()) < float(self.TK_RCF_Strat.text()):
                i = 0
                while float(self.TK_RCF_End.text()) + float(self.TK_RCF_interval.text()) * i < float(self.TK_RCF_Strat.text()):
                    self.RCF_TK_list.append(float(self.TK_RCF_End.text()) + float(self.TK_RCF_interval.text()) * i)
                    i += 1
                self.RCF_TK_list.append(float(self.TK_RCF_Strat.text()))

            print("self.RCF_TK_list",self.RCF_TK_list)

            AK_RCF_Change_spectrum_Series_list = []
            for RCF_TK in self.RCF_TK_list:
                # 先将 R_aK_TK 添加到 Sort_TK_SP_list
                Sort_TK_SP_list.append(RCF_TK)
                # 然后对更新后的列表进行排序
                Sort_TK_SP_AK_list = sorted(Sort_TK_SP_list, reverse=True)
                print("Sort_TK_SP_AK_list", Sort_TK_SP_AK_list)
                # 4種情況
                if RCF_TK == max(Sort_TK_SP_AK_list):
                    print("max")
                    AK_RCF_Change_spectrum_Series = K_list[0] * np.exp(-1 * A_list[0] * RCF_TK)
                    AK_RCF_Change_spectrum_Series_list.append(AK_RCF_Change_spectrum_Series)
                    # print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # return AK_RCF_Change_spectrum_Series
                elif RCF_TK == min(Sort_TK_SP_AK_list):
                    print("min")
                    AK_RCF_Change_spectrum_Series = K_list[len(K_list)-1] * np.exp(-1 * A_list[len(A_list)-1] * RCF_TK)
                    AK_RCF_Change_spectrum_Series_list.append(AK_RCF_Change_spectrum_Series)
                    # print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # return AK_RCF_Change_spectrum_Series
                elif RCF_TK in TK_SP_list:
                    print("equal")
                    AK_RCF_Change_spectrum_Series = SP_series_list[TK_SP_list.index(RCF_TK)][1:].reset_index(drop=True)
                    AK_RCF_Change_spectrum_Series_list.append(AK_RCF_Change_spectrum_Series)
                    # print("AK_RCF_Change_spectrum_Series",AK_RCF_Change_spectrum_Series)
                    # return AK_RCF_Change_spectrum_Series
                else:
                    print("mid")
                    position = Sort_TK_SP_AK_list.index(RCF_TK)
                    AK_RCF_Change_spectrum_Series = K_list[position-1] * np.exp(-1 * A_list[position-1] * RCF_TK)
                    AK_RCF_Change_spectrum_Series_list.append(AK_RCF_Change_spectrum_Series)
                    # print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # return AK_RCF_Change_spectrum_Series
            print("AK_RCF_Change_spectrum_Series_list",AK_RCF_Change_spectrum_Series_list)
            for row,TK in enumerate(self.RCF_TK_list):
                # item = QTableWidgetItem(str(TK))
                self.TK_RCF_Table_Form.setItem(row + 1, 3, QTableWidgetItem(f"{TK:.3f}"))

            # 關閉連線
            connection_RCF_Change.close()

            return AK_RCF_Change_spectrum_Series_list
        else:
            connection_RCF_Differ = sqlite3.connect("RCF_Differ_spectrum.db")
            cursor_RCF_Differ = connection_RCF_Differ.cursor()
            # 取得Rdiffer資料
            column_name_RCF_Differ = self.TK_RCF_Combobox.currentText()
            # print("column_name_RCF_Differ",column_name_RCF_Differ)
            table_name_RCF_Differ = self.TK_RCF_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_RCF_Differ = f"SELECT * FROM '{table_name_RCF_Differ}';"
            cursor_RCF_Differ.execute(query_RCF_Differ)
            result_RCF_Differ = cursor_RCF_Differ.fetchall()

            # 找到指定標題的欄位索引
            header_RCF_Differ = [column[0] for column in cursor_RCF_Differ.description]

            # 移除所有標題中的編號
            header_RCF_Differ = [re.sub(r'_\d+$', '', col) for col in header_RCF_Differ]
            # print("header_RCF_Differ", header_RCF_Differ)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_RCF_Differ) if col == column_name_RCF_Differ]
            # print("indices_without_number", indices_without_number)

            # 使用 pandas Series 存儲結果
            series_list = [pd.Series([row[i] for row in result_RCF_Differ]) for i in indices_without_number]

            # print("series_list", series_list)
            TK_record = []
            series_list_length = len(series_list)
            for i in range(series_list_length):
                try:
                    TK_record.append(float(series_list[i].iloc[0]))
                except ValueError:
                    C_Series = 1
                    # 發生錯誤時，忽略並繼續執行
                    return C_Series
            self.TK_RCF_Range.setText(f"TKrange: {min(TK_record):.2f}~{max(TK_record):.2f}")
            self.RCF_TK_list = []
            if float(self.TK_RCF_Strat.text()) < float(self.TK_RCF_End.text()):
                i = 0
                while float(self.TK_RCF_Strat.text()) + float(self.TK_RCF_interval.text()) * i < float(self.TK_RCF_End.text()):
                    self.RCF_TK_list.append(float(self.TK_RCF_Strat.text()) + float(self.TK_RCF_interval.text()) * i)
                    i += 1
                self.RCF_TK_list.append(float(self.TK_RCF_End.text()))

            elif float(self.TK_RCF_End.text()) < float(self.TK_RCF_Strat.text()):
                i = 0
                while float(self.TK_RCF_End.text()) + float(self.TK_RCF_interval.text()) * i < float(self.TK_RCF_Strat.text()):
                    self.RCF_TK_list.append(float(self.TK_RCF_End.text()) + float(self.TK_RCF_interval.text()) * i)
                    i += 1
                self.RCF_TK_list.append(float(self.TK_RCF_Strat.text()))

            print("self.RCF_TK_list",self.RCF_TK_list)
            AK_RCF_Change_spectrum_Series_list = []
            for RCF_TK in self.RCF_TK_list:
                R_differ_TK = RCF_TK
                Thickness_list = [R_differ_TK]
                for i in range(series_list_length):
                    Thickness_list.append(float(series_list[i].iloc[0]))
                    TK_record.append(float(series_list[i].iloc[0]))
                Thickness_list_2 = sorted(Thickness_list)
                print("Thickness_list_2", Thickness_list_2)
                index_R_differ_TK = Thickness_list_2.index(R_differ_TK)
                print(Thickness_list_2.index(R_differ_TK))

                # 輸入厚度的三種情況
                if R_differ_TK == min(Thickness_list):
                    print("min")
                    A_index_value = Thickness_list_2[index_R_differ_TK + 1]
                    # print("A_index_value", A_index_value)
                    B_index_value = Thickness_list_2[index_R_differ_TK + 2]
                    # print("B_index_value", B_index_value)
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    # print("A_index", A_index)
                    B_index = Thickness_list.index(B_index_value)
                    # print("B_index", B_index)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()

                    # Final R Series
                    C_Series = A_Series + (R_differ_TK - A_index_value) * (B_Series - A_Series) / (
                            B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    AK_RCF_Change_spectrum_Series_list.append(C_Series)
                elif R_differ_TK == max(Thickness_list):
                    print("max")
                    A_index_value = Thickness_list_2[index_R_differ_TK - 2]
                    B_index_value = Thickness_list_2[index_R_differ_TK - 1]
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    B_index = Thickness_list.index(B_index_value)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()
                    # Final R Series
                    C_Series = A_Series + (R_differ_TK - A_index_value) * (B_Series - A_Series) / (
                            B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    AK_RCF_Change_spectrum_Series_list.append(C_Series)
                elif R_differ_TK in (Thickness_list[1:]):
                    print("equal")
                    # 取得對應的索引
                    Thickness_list.remove(Thickness_list[0])
                    equal_index = Thickness_list.index(R_differ_TK)
                    # print("Equal Index:", equal_index)
                    # print("Equal Value:", R_differ_TK)
                    C_Series = series_list[equal_index][1:]
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    AK_RCF_Change_spectrum_Series_list.append(C_Series)
                else:
                    print("mid")
                    # print("R_differ_TK",R_differ_TK)
                    # print("Thickness_list[1:]",Thickness_list[1:])
                    A_index_value = Thickness_list_2[index_R_differ_TK - 1]
                    # print("A_index_value",A_index_value)
                    B_index_value = Thickness_list_2[index_R_differ_TK + 1]
                    # print("B_index_value", B_index_value)
                    Thickness_list.remove(Thickness_list[0])
                    A_index = Thickness_list.index(A_index_value)
                    # print("A_index",A_index)
                    # print("Thickness_list",Thickness_list)
                    B_index = Thickness_list.index(B_index_value)
                    # print("B_index", B_index)
                    # A Series
                    A_Series = series_list[A_index][1:]
                    A_Series = pd.to_numeric(A_Series, errors='coerce')
                    A_Series = A_Series.fillna(0)
                    A_Series = A_Series.dropna()
                    # print("A_Series",A_Series)
                    # B Series
                    B_Series = series_list[B_index][1:]
                    B_Series = pd.to_numeric(B_Series, errors='coerce')
                    B_Series = B_Series.fillna(0)
                    B_Series = B_Series.dropna()
                    # print("B_Series", B_Series)

                    # Final R Series
                    C_Series = A_Series + (R_differ_TK - A_index_value) * (B_Series - A_Series) / (
                                B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    AK_RCF_Change_spectrum_Series_list.append(C_Series)
            # 關閉連線
            connection_RCF_Differ.close()
            for row,TK in enumerate(self.RCF_TK_list):
                # item = QTableWidgetItem(str(TK))
                self.TK_RCF_Table_Form.setItem(row + 1, 3, QTableWidgetItem(f"{TK:.3f}"))
            print("AK_RCF_Change_spectrum_Series_list",AK_RCF_Change_spectrum_Series_list)
            return AK_RCF_Change_spectrum_Series_list





    # updatealltable
    def updateall(self):
        self.update_light_source_datatable()
        self.updateLightSourceComboBox()
        self.update_light_source_led_datatable()
        self.updateLightSourceledComboBox()
        self.update_source_led_datatable()
        self.updateSourceledComboBox()
        self.update_layer1_datatable()
        self.updatelayer1ComboBox()
        self.update_layer2_datatable()
        self.updatelayer2ComboBox()
        self.update_layer3_datatable()
        self.updatelayer3ComboBox()
        self.update_layer4_datatable()
        self.updatelayer4ComboBox()
        self.update_layer5_datatable()
        self.updatelayer5ComboBox()
        self.update_layer6_datatable()
        self.updatelayer6ComboBox()
        self.update_RCF_Fix_datatable()
        self.updateRCF_Fix_ComboBox()
        self.update_GCF_Fix_datatable()
        self.updateGCF_Fix_ComboBox()
        self.update_BCF_Fix_datatable()
        self.updateBCF_Fix_ComboBox()
        self.update_RCF_Change_datatable()
        self.updateRCF_Change_ComboBox()
        self.update_GCF_Change_datatable()
        self.updateGCF_Change_ComboBox()
        self.update_BCF_Change_datatable()
        self.updateBCF_Change_ComboBox()
        self.update_RCF_Differ_datatable()
        self.updateRCF_Differ_ComboBox()
        self.update_GCF_Differ_datatable()
        self.updateGCF_Differ_ComboBox()
        self.update_BCF_Differ_datatable()
        self.updateBCF_Differ_ComboBox()


    def Color_Simulation(self):
        # 檢查是否已經存在對話框
        if hasattr(self, "colorsimdialog") and self.colorsimdialog.isVisible():
            # 如果存在且可見，則關閉舊的對話框
            self.colorsimdialog.close()
        # 創建新的對話框
        self.colorsimdialog = QDialog(None, Qt.Window)  # 這樣才有縮小鍵
        self.colorsimdialog.setWindowTitle("Color_Simulation")
        self.colorsimdialog.setStyleSheet("background-color: lightgrey;")
        self.colorsimdialog.resize(1000, 800)
        # Apply styles to the dialog
        self.colorsimdialog.setStyleSheet("""
                            QDialog {
                                background-color: lightyellow;
                            }

                            QTableWidget {
                                background-color: white;
                                alternate-background-color: lightgray;
                                selection-background-color: lightgreen;
                            }

                            QComboBox {
                                background-color: #6495ED;
                                selection-background-color: #4169E1;
                                color: black;  /* Set the text color */
                            }
                            QTableWidget::item:selected {
                        color: blcak; /* 設定文字顏色為黑色 */
                        background-color: #008080; /* 設定背景顏色為藍色，你可以根據需要調整 */
                            }
                        """)

        # dialog layout
        self.Color_Select_layout = QGridLayout(self.colorsimdialog)

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

        # Thickness select
        self.RGB_1_checkbox = QCheckBox()
        self.RGB_1_checkbox.setChecked(False)  # 預設非勾選狀態
        self.RGB_1_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.RGB_1_label = QLabel("R=G=B")

        self.RGB_2_checkbox = QCheckBox()
        self.RGB_2_checkbox.setChecked(False)  # 預設非勾選狀態
        self.RGB_2_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.RGB_2_label = QLabel("R=G>B")
        self.RGB_2_tolerance = QLabel("Thickness_Gap")
        self.RGB_2_tolerance_edit = QLineEdit()

        self.RGB_3_checkbox = QCheckBox()
        self.RGB_3_checkbox.setChecked(False)  # 預設非勾選狀態
        self.RGB_3_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.RGB_3_label = QLabel("R>G=B")
        self.RGB_3_tolerance = QLabel("Thickness_Gap")
        self.RGB_3_tolerance_edit = QLineEdit()

        self.RGB_4_checkbox = QCheckBox()
        self.RGB_4_checkbox.setChecked(False)  # 預設非勾選狀態
        self.RGB_4_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.RGB_4_label = QLabel("R>G>B")
        self.RGB_4_tolerance = QLabel("Thickness_Gap")
        self.RGB_4_tolerance_edit = QLineEdit()

        # table
        self.color_sim_table = QTableWidget()
        self.color_sim_table.setColumnCount(40)
        self.color_sim_table.setRowCount(200)
        self.color_sim_table.setFixedSize(1000, 600)
        # Apply styles to the dialog
        self.color_sim_table.setStyleSheet("""
                                    QTableWidget::item:selected {
                                color: blcak; /* 設定文字顏色為黑色 */
                                background-color: #008080; /* 設定背景顏色為藍色，你可以根據需要調整 */
                                    }
                                """)

        # 設定默認值
        column1_default_values = ["項目", "Wx", "Wy", "WY", "Rx", "Ry", "RY", "Rλ", "R_Purity", "Gx", "Gy", "GY",
                                  "Gλ", "G_Purity", "Bx", "By", "BY", "B_λ", "B_Purity", "NTSC%", "BLUx", "BLUy",
                                  "R色阻選擇", "R色阻厚度", "G色阻選擇", "G色阻厚度", "B色阻選擇", "B色阻厚度",
                                  "背光選擇", "原LED選擇", "替換LED", "Layer_1", "Layer_2", "Layer_3", "Layer_4",
                                  "Layer_5", "Layer_6"
                                  ]
        for row, values in enumerate([column1_default_values]):
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setBackground(QColor(173, 216, 230))  # 設置背景顏色為淺藍色
                item.setTextAlignment(Qt.AlignCenter)  # 設置文本居中對齊
                self.color_sim_table.setItem(row, column, item)

        # 設定特定單元格顏色
        color_ranges = [(1, 3), (4, 8), (9, 13), (14, 18)]
        colors = [QColor(RESULTWHITE), QColor(RESULTRED), QColor(RESULTGREEN), QColor(RESULTBLUE)]

        for row in range(self.color_sim_table.rowCount()):
            for (start_col, end_col), color in zip(color_ranges, colors):
                for col in range(start_col, end_col + 1):
                    item = self.color_sim_table.item(row, col)
                    if item:
                        item.setBackground(color)

        # Simulate_button
        self.simbutton = QPushButton("Simulation")
        self.simbutton.clicked.connect(self.calculate_color_W_Fix_customize_sim)

        # widget放置區
        self.Color_Select_layout.addWidget(Color_Select_label, 0, 0, 1, 14)
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

        # Thickness
        self.Color_Select_layout.addWidget(self.RGB_1_checkbox, 3, 0)
        self.Color_Select_layout.addWidget(self.RGB_1_label, 3, 1)

        self.Color_Select_layout.addWidget(self.RGB_2_checkbox, 4, 0)
        self.Color_Select_layout.addWidget(self.RGB_2_label, 4, 1)
        self.Color_Select_layout.addWidget(self.RGB_2_tolerance, 4, 2)
        self.Color_Select_layout.addWidget(self.RGB_2_tolerance_edit, 4, 3)
        self.Color_Select_layout.addWidget(self.RGB_3_checkbox, 5, 0)
        self.Color_Select_layout.addWidget(self.RGB_3_label, 5, 1)
        self.Color_Select_layout.addWidget(self.RGB_3_tolerance, 5, 2)
        self.Color_Select_layout.addWidget(self.RGB_3_tolerance_edit, 5, 3)
        self.Color_Select_layout.addWidget(self.RGB_4_checkbox, 6, 0)
        self.Color_Select_layout.addWidget(self.RGB_4_label, 6, 1)
        self.Color_Select_layout.addWidget(self.RGB_4_tolerance, 6, 2)
        self.Color_Select_layout.addWidget(self.RGB_4_tolerance_edit, 6, 3)

        # button
        self.Color_Select_layout.addWidget(self.simbutton, 7, 0)


        # table
        self.Color_Select_layout.addWidget(self.color_sim_table,8,0,5,14)

        # 使用 QSpacerItem 創建底部空間
        bottom_spacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.Color_Select_layout.addItem(bottom_spacer, 15, 0, 1, 14)  # 在第 3 行擴展到全部列

        self.Color_Select_layout.setHorizontalSpacing(5)  # 設定水平間距為5像素

        self.Color_Select_layout.setVerticalSpacing(1)  # 設定水平間距為5像素

        # # layout放置
        # self.setLayout(self.Color_Select_layout)
        #
        # # Set Background
        # self.setStyleSheet("background-color: #FFBD9D;")

        # 顯示新的對話框
        self.colorsimdialog.show()

    def calculate_RCF_Fix_sim(self):
        connection_RCF_Fix = sqlite3.connect("RCF_Fix_spectrum.db")
        cursor_RCF_Fix = connection_RCF_Fix.cursor()
        # 取得BLU資料
        table_name_RCF_Fix = self.R_fix_table.currentText()
        # 使用正確的引號包裹表名和列名
        query_RCF_Fix = f"SELECT * FROM '{table_name_RCF_Fix}';"
        cursor_RCF_Fix.execute(query_RCF_Fix)
        # 獲取查詢結果
        result_RCF_Fix = cursor_RCF_Fix.fetchall()

        # 獲取標頭（列名），排除第一個元素
        self.RCF_Fix_Sim_headers = [i[0] for i in cursor_RCF_Fix.description]
        # 為了去掉第一個元素
        self.RCF_Fix_Sim_headers_use = self.RCF_Fix_Sim_headers[1:]
        print("self.RCF_Fix_Sim_headers_use ", self.RCF_Fix_Sim_headers_use)

        # 使用 pandas DataFrame 儲存資料
        df = pd.DataFrame(result_RCF_Fix, columns=self.RCF_Fix_Sim_headers )

        # 初始化一個空列表來儲存每個欄位的第零列值
        self.R_FIX_TK_sim_list = []

        # 遍歷 DataFrame 的所有欄位
        for column in df.columns[1:]:
            # 將每個欄位的第零列值添加到列表中
            self.R_FIX_TK_sim_list.append(df[column].iloc[0])

        print("self.R_FIX_TK_sim_list",self.R_FIX_TK_sim_list)

        # 初始化一個空列表來儲存每個欄位的數據
        self.R_FIX_sim_Spectrum_list = []

        # 遍歷 DataFrame 的所有欄位
        for column in df.columns[1:]:
            # 將每個欄位從第一列開始的數據添加到列表中
            self.R_FIX_sim_Spectrum_list.append(df[column][1:].reset_index(drop=True))

        #print("self.R_FIX_sim_Spectrum_list",self.R_FIX_sim_Spectrum_list)
        # print("Type",type(R_FIX_Spectrum_list[0]))
        # print(R_FIX_Spectrum_list[0] * 5)

        # 關閉資料庫連接
        connection_RCF_Fix.close()
        return self.RCF_Fix_Sim_headers,self.R_FIX_TK_sim_list,self.R_FIX_sim_Spectrum_list

    def calculate_GCF_Fix_sim(self):
        connection_GCF_Fix = sqlite3.connect("GCF_Fix_spectrum.db")
        cursor_GCF_Fix = connection_GCF_Fix.cursor()
        # 取得BLU資料
        table_name_GCF_Fix = self.G_fix_table.currentText()
        # 使用正確的引號包裹表名和列名
        query_GCF_Fix = f"SELECT * FROM '{table_name_GCF_Fix}';"
        cursor_GCF_Fix.execute(query_GCF_Fix)
        # 獲取查詢結果
        result_GCF_Fix = cursor_GCF_Fix.fetchall()

        # 獲取標頭（列名）
        self.GCF_Fix_Sim_headers = [i[0] for i in cursor_GCF_Fix.description]
        # 為了去掉第一個元素
        self.GCF_Fix_Sim_headers_use = self.GCF_Fix_Sim_headers[1:]
        print("self.GCF_Fix_Sim_headers_use ", self.GCF_Fix_Sim_headers_use)

        # 使用 pandas DataFrame 儲存資料
        df = pd.DataFrame(result_GCF_Fix, columns=self.GCF_Fix_Sim_headers )

        # 初始化一個空列表來儲存每個欄位的第零列值
        self.G_FIX_TK_sim_list = []

        # 遍歷 DataFrame 的所有欄位
        for column in df.columns[1:]:
            # 將每個欄位的第零列值添加到列表中
            self.G_FIX_TK_sim_list.append(df[column].iloc[0])

        print("self.G_FIX_TK_sim_list",self.G_FIX_TK_sim_list)

        # 初始化一個空列表來儲存每個欄位的數據
        self.G_FIX_sim_Spectrum_list = []

        # 遍歷 DataFrame 的所有欄位
        for column in df.columns[1:]:
            # 將每個欄位從第一列開始的數據添加到列表中
            self.G_FIX_sim_Spectrum_list.append(df[column][1:].reset_index(drop=True))

        #print("self.G_FIX__sim_Spectrum_list",self.G_FIX_sim_Spectrum_list)
        # print("Type",type(G_FIX_Spectrum_list[0]))
        # print(G_FIX_Spectrum_list[0] * 5)

        # 關閉資料庫連接
        connection_GCF_Fix.close()
        return self.GCF_Fix_Sim_headers,self.G_FIX_TK_sim_list,self.G_FIX_sim_Spectrum_list

    def calculate_BCF_Fix_sim(self):
        connection_BCF_Fix = sqlite3.connect("BCF_Fix_spectrum.db")
        cursor_BCF_Fix = connection_BCF_Fix.cursor()
        # 取得BLU資料
        table_name_BCF_Fix = self.B_fix_table.currentText()
        # 使用正確的引號包裹表名和列名
        query_BCF_Fix = f"SELECT * FROM '{table_name_BCF_Fix}';"
        cursor_BCF_Fix.execute(query_BCF_Fix)
        # 獲取查詢結果
        result_BCF_Fix = cursor_BCF_Fix.fetchall()

        # 獲取標頭（列名）
        self.BCF_Fix_Sim_headers = [i[0] for i in cursor_BCF_Fix.description]
        # 為了去掉第一個元素
        self.BCF_Fix_Sim_headers_use = self.BCF_Fix_Sim_headers[1:]
        print("self.BCF_Fix_Sim_headers_use ",self.BCF_Fix_Sim_headers_use)

        # 使用 pandas DataFrame 儲存資料
        df = pd.DataFrame(result_BCF_Fix, columns=self.BCF_Fix_Sim_headers )

        # 初始化一個空列表來儲存每個欄位的第零列值
        self.B_FIX_TK_sim_list = []

        # 遍歷 DataFrame 的所有欄位
        for column in df.columns[1:]:
            # 將每個欄位的第零列值添加到列表中
            self.B_FIX_TK_sim_list.append(df[column].iloc[0])

        print("self.B_FIX_TK_sim_list",self.B_FIX_TK_sim_list)

        # 初始化一個空列表來儲存每個欄位的數據
        self.B_FIX_sim_Spectrum_list = []

        # 遍歷 DataFrame 的所有欄位
        for column in df.columns[1:]:
            # 將每個欄位從第一列開始的數據添加到列表中
            self.B_FIX_sim_Spectrum_list.append(df[column][1:].reset_index(drop=True))

        #print("self.B_FIX__sim_Spectrum_list",self.B_FIX_sim_Spectrum_list)
        # print("Type",type(B_FIX_Spectrum_list[0]))
        # print(B_FIX_Spectrum_list[0] * 5)

        # 關閉資料庫連接
        connection_BCF_Fix.close()
        return self.BCF_Fix_Sim_headers,self.B_FIX_TK_sim_list,self.B_FIX_sim_Spectrum_list

    def CIEparameter(self):
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
        self.CIE_spectrum_Series_X = CIE_spectrum_Series_X.astype(float)
        self.CIE_spectrum_Series_Y = CIE_spectrum_Series_Y.astype(float)
        self.CIE_spectrum_Series_Z = CIE_spectrum_Series_Z.astype(float)

    def calculate_color_RCF_Fix_customize_sim(self):
        self.CIEparameter()
        # BLU
        RxSxxl = self.CIE_spectrum_Series_X * self.calculate_BLU()
        RxSxyl = self.CIE_spectrum_Series_Y * self.calculate_BLU()
        RxSxzl = self.CIE_spectrum_Series_Z * self.calculate_BLU()

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

        # BLU +Cell part
        self.cell_blu_total_spectrum = self.calculate_BLU() * self.calculate_layer1() * self.calculate_layer2() \
                                  * self.calculate_layer3() * self.calculate_layer4() * self.calculate_layer5() \
                                  * self.calculate_layer6()
        # R_Fix
        if self.R_fix_mode.currentText() == "模擬":
            self.calculate_RCF_Fix_sim()
            #print("(self.cell_blu_total_spectrum",self.cell_blu_total_spectrum)
            # 轉換 self.R_FIX_sim_Spectrum_list 中的每個 Series 為 float64 類型
            self.R_FIX_sim_Spectrum_list_converted = [pd.to_numeric(series, errors='coerce').fillna(0) for series in
                                                 self.R_FIX_sim_Spectrum_list]

            # 使用轉換後的列表進行元素乘法
            R_fix_sim_spectrum_list = [series.multiply(self.cell_blu_total_spectrum) for series in
                                       self.R_FIX_sim_Spectrum_list_converted]
            #print("R_fix_sim_spectrum_list",R_fix_sim_spectrum_list)
            self.R_Fix_X_list = [series.multiply(self.CIE_spectrum_Series_X) for series in
                                       R_fix_sim_spectrum_list]
            #print("R_X_list",R_X_list)
            self.R_Fix_Y_list = [series.multiply(self.CIE_spectrum_Series_Y) for series in
                                       R_fix_sim_spectrum_list]
            self.R_Fix_Z_list = [series.multiply(self.CIE_spectrum_Series_Z) for series in
                                       R_fix_sim_spectrum_list]
            self.R_X_fix_sim_sum_list = []
            for i in range(len(self.R_Fix_X_list)):
                self.R_X_fix_sim_sum_list.append(self.R_Fix_X_list[i].sum())
            #print("self.R_X_fix_sim_sum_list",self.R_X_fix_sim_sum_list)
            self.R_Y_fix_sim_sum_list = []
            for i in range(len(self.R_Fix_X_list)):
                self.R_Y_fix_sim_sum_list.append(self.R_Fix_Y_list[i].sum())
            #print("self.R_Y_fix_sim_sum_list", self.R_Y_fix_sim_sum_list)
            self.R_Z_fix_sim_sum_list = []
            for i in range(len(self.R_Fix_X_list)):
                self.R_Z_fix_sim_sum_list.append(self.R_Fix_Z_list[i].sum())
            #print("self.R_Z_fix_sim_sum_list", self.R_Z_fix_sim_sum_list)

            self.R_Fix_x_sim = []
            self.R_Fix_y_sim = []
            self.R_Fix_T_sim = []
            for i in range(len(self.R_X_fix_sim_sum_list)):
                self.R_Fix_x_sim.append(self.R_X_fix_sim_sum_list[i] / (self.R_X_fix_sim_sum_list[i] + self.R_Y_fix_sim_sum_list[i] + self.R_Z_fix_sim_sum_list[i]))

            for i in range(len(self.R_Y_fix_sim_sum_list)):
                self.R_Fix_y_sim.append(self.R_Y_fix_sim_sum_list[i] / (
                            self.R_X_fix_sim_sum_list[i] + self.R_Y_fix_sim_sum_list[i] + self.R_Z_fix_sim_sum_list[i]))

            for i in range(len(self.R_Y_fix_sim_sum_list)):
                self.R_Fix_T_sim.append(self.R_Y_fix_sim_sum_list[i] * 3)

            print("self.R_Fix_x_sim", self.R_Fix_x_sim)
            print("self.R_Fix_y_sim", self.R_Fix_y_sim)
            print("self.R_Fix_T_sim", self.R_Fix_T_sim)
            self.R_Fix_wave_sim = []
            self.R_Fix_purity_sim = []

            for i in range(len(self.R_X_fix_sim_sum_list)):
                xy = [self.R_Fix_x_sim[i], self.R_Fix_y_sim[i]]
                xy_n = self.observer_D65
                # 計算主波長(
                dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
                dominant_wavelength_value = dominant_wavelength
                CIEcoordinate_value_1 = CIEcoordinate[0]
                CIEcoordinate_value_2 = CIEcoordinate[1]
                # 計算距離
                distance_from_white = math.sqrt((self.R_Fix_x_sim[i] - xy_n[0]) ** 2 + (self.R_Fix_y_sim[i] - xy_n[1]) ** 2)
                distance_from_black = math.sqrt(
                    (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
                # 計算色純度
                purity = distance_from_white / distance_from_black * 100
                wave = str(dominant_wavelength_value)
                self.R_Fix_wave_sim.append(wave)
                self.R_Fix_purity_sim.append(purity)
            print("self.R_Fix_wave_sim", self.R_Fix_wave_sim)
            print("self.R_Fix_purity_sim", self.R_Fix_purity_sim)

    def calculate_color_GCF_Fix_customize_sim(self):
        self.CIEparameter()
        # BLU
        GxSxxl = self.CIE_spectrum_Series_X * self.calculate_BLU()
        GxSxyl = self.CIE_spectrum_Series_Y * self.calculate_BLU()
        GxSxzl = self.CIE_spectrum_Series_Z * self.calculate_BLU()

        BLU_spectrum_Series_sum = self.calculate_BLU().sum()
        k = 100 / BLU_spectrum_Series_sum
        GxSxxl_sum = GxSxxl.sum()
        GxSxyl_sum = GxSxyl.sum()
        GxSxzl_sum = GxSxzl.sum()
        GxSxxl_sum_k = GxSxxl_sum * k
        GxSxyl_sum_k = GxSxyl_sum * k
        GxSxzl_sum_k = GxSxzl_sum * k
        BLU_x = GxSxxl_sum_k / (GxSxxl_sum_k + GxSxyl_sum_k + GxSxzl_sum_k)
        BLU_y = GxSxyl_sum_k / (GxSxxl_sum_k + GxSxyl_sum_k + GxSxzl_sum_k)

        # BLU +Cell part
        self.cell_blu_total_spectrum = self.calculate_BLU() * self.calculate_layer1() * self.calculate_layer2() \
                                  * self.calculate_layer3() * self.calculate_layer4() * self.calculate_layer5() \
                                  * self.calculate_layer6()
        # R_Fix
        if self.G_fix_mode.currentText() == "模擬":
            self.calculate_GCF_Fix_sim()
            #print("(self.cell_blu_total_spectrum",self.cell_blu_total_spectrum)
            # 轉換 self.G_FIX_sim_Spectrum_list 中的每個 Series 為 float64 類型
            self.G_FIX_sim_Spectrum_list_converted = [pd.to_numeric(series, errors='coerce').fillna(0) for series in
                                                 self.G_FIX_sim_Spectrum_list]

            # 使用轉換後的列表進行元素乘法
            G_fix_sim_spectrum_list = [series.multiply(self.cell_blu_total_spectrum) for series in
                                       self.G_FIX_sim_Spectrum_list_converted]
            #print("R_fix_sim_spectrum_list",R_fix_sim_spectrum_list)
            self.G_Fix_X_list = [series.multiply(self.CIE_spectrum_Series_X) for series in
                                       G_fix_sim_spectrum_list]
            #print("R_X_list",R_X_list)
            self.G_Fix_Y_list = [series.multiply(self.CIE_spectrum_Series_Y) for series in
                                       G_fix_sim_spectrum_list]
            self.G_Fix_Z_list = [series.multiply(self.CIE_spectrum_Series_Z) for series in
                                       G_fix_sim_spectrum_list]
            self.G_X_fix_sim_sum_list = []
            for i in range(len(self.G_Fix_X_list)):
                self.G_X_fix_sim_sum_list.append(self.G_Fix_X_list[i].sum())
            #print("self.G_X_fix_sim_sum_list",self.G_X_fix_sim_sum_list)
            self.G_Y_fix_sim_sum_list = []
            for i in range(len(self.G_Fix_X_list)):
                self.G_Y_fix_sim_sum_list.append(self.G_Fix_Y_list[i].sum())
            #print("self.G_Y_fix_sim_sum_list", self.G_Y_fix_sim_sum_list)
            self.G_Z_fix_sim_sum_list = []
            for i in range(len(self.G_Fix_X_list)):
                self.G_Z_fix_sim_sum_list.append(self.G_Fix_Z_list[i].sum())
            #print("self.G_Z_fix_sim_sum_list", self.G_Z_fix_sim_sum_list)

            self.G_Fix_x_sim = []
            self.G_Fix_y_sim = []
            self.G_Fix_T_sim = []
            for i in range(len(self.G_X_fix_sim_sum_list)):
                self.G_Fix_x_sim.append(self.G_X_fix_sim_sum_list[i] / (self.G_X_fix_sim_sum_list[i] + self.G_Y_fix_sim_sum_list[i] + self.G_Z_fix_sim_sum_list[i]))

            for i in range(len(self.G_Y_fix_sim_sum_list)):
                self.G_Fix_y_sim.append(self.G_Y_fix_sim_sum_list[i] / (
                            self.G_X_fix_sim_sum_list[i] + self.G_Y_fix_sim_sum_list[i] + self.G_Z_fix_sim_sum_list[i]))

            for i in range(len(self.G_Y_fix_sim_sum_list)):
                self.G_Fix_T_sim.append(self.G_Y_fix_sim_sum_list[i] * 3)

            print("self.G_Fix_x_sim", self.G_Fix_x_sim)
            print("self.G_Fix_y_sim", self.G_Fix_y_sim)
            print("self.G_Fix_T_sim", self.G_Fix_T_sim)
            self.G_Fix_wave_sim = []
            self.G_Fix_purity_sim = []

            for i in range(len(self.G_X_fix_sim_sum_list)):
                xy = [self.G_Fix_x_sim[i], self.G_Fix_y_sim[i]]
                xy_n = self.observer_D65
                # 計算主波長(
                dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
                dominant_wavelength_value = dominant_wavelength
                CIEcoordinate_value_1 = CIEcoordinate[0]
                CIEcoordinate_value_2 = CIEcoordinate[1]
                # 計算距離
                distance_from_white = math.sqrt((self.G_Fix_x_sim[i] - xy_n[0]) ** 2 + (self.G_Fix_y_sim[i] - xy_n[1]) ** 2)
                distance_from_black = math.sqrt(
                    (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
                # 計算色純度
                purity = distance_from_white / distance_from_black * 100
                wave = str(dominant_wavelength_value)
                self.G_Fix_wave_sim.append(wave)
                self.G_Fix_purity_sim.append(purity)
            print("self.G_Fix_wave_sim", self.G_Fix_wave_sim)
            print("self.G_Fix_purity_sim", self.G_Fix_purity_sim)

    def calculate_color_BCF_Fix_customize_sim(self):
        self.CIEparameter()
        # BLU
        BxSxxl = self.CIE_spectrum_Series_X * self.calculate_BLU()
        BxSxyl = self.CIE_spectrum_Series_Y * self.calculate_BLU()
        BxSxzl = self.CIE_spectrum_Series_Z * self.calculate_BLU()

        BLU_spectrum_Series_sum = self.calculate_BLU().sum()
        k = 100 / BLU_spectrum_Series_sum
        BxSxxl_sum = BxSxxl.sum()
        BxSxyl_sum = BxSxyl.sum()
        BxSxzl_sum = BxSxzl.sum()
        BxSxxl_sum_k = BxSxxl_sum * k
        BxSxyl_sum_k = BxSxyl_sum * k
        BxSxzl_sum_k = BxSxzl_sum * k
        BLU_x = BxSxxl_sum_k / (BxSxxl_sum_k + BxSxyl_sum_k + BxSxzl_sum_k)
        BLU_y = BxSxyl_sum_k / (BxSxxl_sum_k + BxSxyl_sum_k + BxSxzl_sum_k)

        # BLU +Cell part
        self.cell_blu_total_spectrum = self.calculate_BLU() * self.calculate_layer1() * self.calculate_layer2() \
                                  * self.calculate_layer3() * self.calculate_layer4() * self.calculate_layer5() \
                                  * self.calculate_layer6()
        # R_Fix
        if self.B_fix_mode.currentText() == "模擬":
            self.calculate_BCF_Fix_sim()
            #print("(self.cell_blu_total_spectrum",self.cell_blu_total_spectrum)
            # 轉換 self.B_FIX_sim_Spectrum_list 中的每個 Series 為 float64 類型
            self.B_FIX_sim_Spectrum_list_converted = [pd.to_numeric(series, errors='coerce').fillna(0) for series in
                                                 self.B_FIX_sim_Spectrum_list]

            # 使用轉換後的列表進行元素乘法
            B_fix_sim_spectrum_list = [series.multiply(self.cell_blu_total_spectrum) for series in
                                       self.B_FIX_sim_Spectrum_list_converted]
            #print("R_fix_sim_spectrum_list",R_fix_sim_spectrum_list)
            self.B_Fix_X_list = [series.multiply(self.CIE_spectrum_Series_X) for series in
                                       B_fix_sim_spectrum_list]
            #print("R_X_list",R_X_list)
            self.B_Fix_Y_list = [series.multiply(self.CIE_spectrum_Series_Y) for series in
                                       B_fix_sim_spectrum_list]
            self.B_Fix_Z_list = [series.multiply(self.CIE_spectrum_Series_Z) for series in
                                       B_fix_sim_spectrum_list]
            self.B_X_fix_sim_sum_list = []
            for i in range(len(self.B_Fix_X_list)):
                self.B_X_fix_sim_sum_list.append(self.B_Fix_X_list[i].sum())
            #print("self.B_X_fix_sim_sum_list",self.B_X_fix_sim_sum_list)
            self.B_Y_fix_sim_sum_list = []
            for i in range(len(self.B_Fix_X_list)):
                self.B_Y_fix_sim_sum_list.append(self.B_Fix_Y_list[i].sum())
            #print("self.B_Y_fix_sim_sum_list", self.B_Y_fix_sim_sum_list)
            self.B_Z_fix_sim_sum_list = []
            for i in range(len(self.B_Fix_X_list)):
                self.B_Z_fix_sim_sum_list.append(self.B_Fix_Z_list[i].sum())
            #print("self.B_Z_fix_sim_sum_list", self.B_Z_fix_sim_sum_list)

            self.B_Fix_x_sim = []
            self.B_Fix_y_sim = []
            self.B_Fix_T_sim = []
            for i in range(len(self.B_X_fix_sim_sum_list)):
                self.B_Fix_x_sim.append(self.B_X_fix_sim_sum_list[i] / (self.B_X_fix_sim_sum_list[i] + self.B_Y_fix_sim_sum_list[i] + self.B_Z_fix_sim_sum_list[i]))

            for i in range(len(self.B_Y_fix_sim_sum_list)):
                self.B_Fix_y_sim.append(self.B_Y_fix_sim_sum_list[i] / (
                            self.B_X_fix_sim_sum_list[i] + self.B_Y_fix_sim_sum_list[i] + self.B_Z_fix_sim_sum_list[i]))

            for i in range(len(self.B_Y_fix_sim_sum_list)):
                self.B_Fix_T_sim.append(self.B_Y_fix_sim_sum_list[i] * 3)

            print("self.B_Fix_x_sim", self.B_Fix_x_sim)
            print("self.B_Fix_y_sim", self.B_Fix_y_sim)
            print("self.B_Fix_T_sim", self.B_Fix_T_sim)
            self.B_Fix_wave_sim = []
            self.B_Fix_purity_sim = []

            for i in range(len(self.B_X_fix_sim_sum_list)):
                xy = [self.B_Fix_x_sim[i], self.B_Fix_y_sim[i]]
                xy_n = self.observer_D65
                # 計算主波長(
                dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
                dominant_wavelength_value = dominant_wavelength
                CIEcoordinate_value_1 = CIEcoordinate[0]
                CIEcoordinate_value_2 = CIEcoordinate[1]
                # 計算距離
                distance_from_white = math.sqrt((self.B_Fix_x_sim[i] - xy_n[0]) ** 2 + (self.B_Fix_y_sim[i] - xy_n[1]) ** 2)
                distance_from_black = math.sqrt(
                    (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
                # 計算色純度
                purity = distance_from_white / distance_from_black * 100
                wave = str(dominant_wavelength_value)
                self.B_Fix_wave_sim.append(wave)
                self.B_Fix_purity_sim.append(purity)
            print("self.B_Fix_wave_sim", self.B_Fix_wave_sim)
            print("self.B_Fix_purity_sim", self.B_Fix_purity_sim)

    def calculate_color_W_Fix_customize_sim(self):
        self.calculate_color_RCF_Fix_customize_sim()
        self.calculate_color_GCF_Fix_customize_sim()
        self.calculate_color_BCF_Fix_customize_sim()
        # W_Fix_Sim
        all_Fix_W_X_combinations = list(itertools.product(self.R_X_fix_sim_sum_list, self.G_X_fix_sim_sum_list, self.B_X_fix_sim_sum_list))
        all_Fix_W_Y_combinations = list(
            itertools.product(self.R_Y_fix_sim_sum_list, self.G_Y_fix_sim_sum_list, self.B_Y_fix_sim_sum_list))
        all_Fix_W_Z_combinations = list(
            itertools.product(self.R_Z_fix_sim_sum_list, self.G_Z_fix_sim_sum_list, self.B_Z_fix_sim_sum_list))
        self.W_X_Fix_list = []
        self.W_Y_Fix_list = []
        self.W_Z_Fix_list = []
        for i in range (len(all_Fix_W_X_combinations)):
            W_X_Fix_total = all_Fix_W_X_combinations[i][0] + all_Fix_W_X_combinations[i][1]+ all_Fix_W_X_combinations[i][2]
            self.W_X_Fix_list.append(W_X_Fix_total)
        for i in range (len(all_Fix_W_Y_combinations)):
            W_Y_Fix_total = all_Fix_W_Y_combinations[i][0] + all_Fix_W_Y_combinations[i][1]+ all_Fix_W_Y_combinations[i][2]
            self.W_Y_Fix_list.append(W_Y_Fix_total)
        for i in range (len(all_Fix_W_Z_combinations)):
            W_Z_Fix_total = all_Fix_W_Z_combinations[i][0] + all_Fix_W_Z_combinations[i][1]+ all_Fix_W_Z_combinations[i][2]
            self.W_Z_Fix_list.append(W_Z_Fix_total)
        self.W_X_Fix_sum_list = []
        self.W_Y_Fix_sum_list = []
        self.W_Z_Fix_sum_list = []
        for i in range(len(all_Fix_W_X_combinations)):
            self.W_X_Fix_sum_list.append(self.W_X_Fix_list[i].sum())
        for i in range(len(all_Fix_W_Y_combinations)):
            self.W_Y_Fix_sum_list.append(self.W_Y_Fix_list[i].sum())
        for i in range(len(all_Fix_W_Z_combinations)):
            self.W_Z_Fix_sum_list.append(self.W_Z_Fix_list[i].sum())

        self.W_x_Fix_sum_list = []
        self.W_y_Fix_sum_list = []
        for i in range(len(all_Fix_W_X_combinations)):
            W_x_total = self.W_X_Fix_sum_list[i]/(self.W_X_Fix_sum_list[i]+self.W_Y_Fix_sum_list[i]+self.W_Z_Fix_sum_list[i])
            self.W_x_Fix_sum_list.append(W_x_total)
        for i in range(len(all_Fix_W_Y_combinations)):
            W_y_total = self.W_Y_Fix_sum_list[i]/(self.W_X_Fix_sum_list[i]+self.W_Y_Fix_sum_list[i]+self.W_Z_Fix_sum_list[i])
            self.W_y_Fix_sum_list.append(W_y_total)
        print("self.W_x_Fix_sum_list",self.W_x_Fix_sum_list)
        print("self.W_y_Fix_sum_list",self.W_y_Fix_sum_list)
        print("self.W_Y_Fix_sum_list",self.W_Y_Fix_sum_list)
        for i in range(len(all_Fix_W_X_combinations)):
            self.color_sim_table.setItem(i + 1, 1, QTableWidgetItem(f"{self.W_x_Fix_sum_list[i]:.4f}"))
            self.color_sim_table.setItem(i + 1, 2, QTableWidgetItem(f"{self.W_y_Fix_sum_list[i]:.4f}"))
            self.color_sim_table.setItem(i + 1, 3, QTableWidgetItem(f"{self.W_Y_Fix_sum_list[i]:.4f}"))
        # RGB_x_Fix_sim
        # 生成所有可能組合的列表
        # 現在，all_combinations 包含所有可能的組合
        # 每個組合都是一個元組形式 (r, g, b)
        all_Fix_x_combinations = list(itertools.product(self.R_Fix_x_sim, self.G_Fix_x_sim, self.B_Fix_x_sim))
        print("all_Fix_x_combinations",all_Fix_x_combinations)
        # print(len(all_Fix_x_combinations))
        # print(all_Fix_x_combinations[0])
        # print(all_Fix_x_combinations[0][0])
        for i in range(len(all_Fix_x_combinations)):
            self.color_sim_table.setItem(i + 1, 4, QTableWidgetItem(f"{all_Fix_x_combinations[i][0]:.4f}"))
            self.color_sim_table.setItem(i + 1, 9, QTableWidgetItem(f"{all_Fix_x_combinations[i][1]:.4f}"))
            self.color_sim_table.setItem(i + 1, 14, QTableWidgetItem(f"{all_Fix_x_combinations[i][2]:.4f}"))

        # RGB_y_Fix_sim
        all_Fix_y_combinations = list(itertools.product(self.R_Fix_y_sim, self.G_Fix_y_sim, self.B_Fix_y_sim))
        for i in range(len(all_Fix_y_combinations)):
            self.color_sim_table.setItem(i + 1, 5, QTableWidgetItem(f"{all_Fix_y_combinations[i][0]:.4f}"))
            self.color_sim_table.setItem(i + 1, 10, QTableWidgetItem(f"{all_Fix_y_combinations[i][1]:.4f}"))
            self.color_sim_table.setItem(i + 1, 15, QTableWidgetItem(f"{all_Fix_y_combinations[i][2]:.4f}"))
        # RGB_y_Fix_sim
        all_Fix_Y_combinations = list(itertools.product(self.R_Y_fix_sim_sum_list, self.G_Y_fix_sim_sum_list, self.B_Y_fix_sim_sum_list))
        for i in range(len(all_Fix_Y_combinations)):
            self.color_sim_table.setItem(i + 1, 6, QTableWidgetItem(f"{all_Fix_Y_combinations[i][0]:.4f}"))
            self.color_sim_table.setItem(i + 1, 11, QTableWidgetItem(f"{all_Fix_Y_combinations[i][1]:.4f}"))
            self.color_sim_table.setItem(i + 1, 16, QTableWidgetItem(f"{all_Fix_Y_combinations[i][2]:.4f}"))

        # RGB_wave_Fix_sim
        all_Fix_wave_combinations = list(
            itertools.product(self.R_Fix_wave_sim, self.G_Fix_wave_sim, self.B_Fix_wave_sim))
        for i in range(len(all_Fix_wave_combinations)):
            self.color_sim_table.setItem(i + 1, 7, QTableWidgetItem(f"{all_Fix_wave_combinations[i][0]}"))
            self.color_sim_table.setItem(i + 1, 12, QTableWidgetItem(f"{all_Fix_wave_combinations[i][1]}"))
            self.color_sim_table.setItem(i + 1, 17, QTableWidgetItem(f"{all_Fix_wave_combinations[i][2]}"))

        # RGB_purity_Fix_sim
        all_Fix_purity_combinations = list(
            itertools.product(self.R_Fix_purity_sim, self.G_Fix_purity_sim, self.B_Fix_purity_sim))
        for i in range(len(all_Fix_purity_combinations)):
            self.color_sim_table.setItem(i + 1, 8, QTableWidgetItem(f"{all_Fix_purity_combinations[i][0]:.4f}"))
            self.color_sim_table.setItem(i + 1, 13, QTableWidgetItem(f"{all_Fix_purity_combinations[i][1]:.4f}"))
            self.color_sim_table.setItem(i + 1, 18, QTableWidgetItem(f"{all_Fix_purity_combinations[i][2]:.4f}"))

        # NTSC list
        self.RGB_Fix_NTSC_List = []
        for i in range(len(all_Fix_purity_combinations)):
            self.RGB_Fix_NTSC_List.append(100 * 0.5 * abs((all_Fix_x_combinations[i][0] * all_Fix_y_combinations[i][1] + all_Fix_x_combinations[i][1] * all_Fix_y_combinations[i][2] + all_Fix_x_combinations[i][2] * all_Fix_y_combinations[i][0] - (all_Fix_x_combinations[i][1] * all_Fix_y_combinations[i][0]) - (
                    all_Fix_x_combinations[i][2] * all_Fix_y_combinations[i][1]) - (all_Fix_x_combinations[i][0] * all_Fix_y_combinations[i][2]))) / 0.1582)

        for i in range(len(self.RGB_Fix_NTSC_List)):
            self.color_sim_table.setItem(i + 1, 19, QTableWidgetItem(f"{self.RGB_Fix_NTSC_List[i]:.4f}"))

        # R_Fix_CF_material
        all_Fix_material_combinations = list(
            itertools.product(self.RCF_Fix_Sim_headers_use, self.GCF_Fix_Sim_headers_use, self.BCF_Fix_Sim_headers_use))
        for i in range(len(all_Fix_purity_combinations)):
            self.color_sim_table.setItem(i + 1, 22, QTableWidgetItem(f"{all_Fix_material_combinations[i][0]}"))
            self.color_sim_table.setItem(i + 1, 24, QTableWidgetItem(f"{all_Fix_material_combinations[i][1]}"))
            self.color_sim_table.setItem(i + 1, 26, QTableWidgetItem(f"{all_Fix_material_combinations[i][2]}"))


