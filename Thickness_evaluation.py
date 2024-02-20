from PySide6.QtWidgets import QWidget, QVBoxLayout, QLabel, QHBoxLayout, QGridLayout, \
    QFormLayout, QLineEdit, QTabWidget, QTableWidgetItem, QTableWidget, QSizePolicy, QFrame, \
    QPushButton, QAbstractItemView,QComboBox,QPushButton,QCheckBox,QDialog,QFileDialog,QMessageBox,\
    QInputDialog,QHeaderView

from PySide6.QtGui import QKeyEvent,QColor,QPalette,QStandardItem,QKeySequence,QShortcut
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt,Signal,QObject
from PySide6.QtCharts import QChart
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import sqlite3
from Setting import *
# import chardet
import openpyxl
import pandas as pd
import numpy as np
import re
import colour
import math
from colour import xy_to_XYZ, XYZ_to_sRGB, CCS_ILLUMINANTS
from colour.plotting import plot_chromaticity_diagram_CIE1931, plot_single_colour_swatch
import colour
from colour.plotting import *
from colour import sd_single_led, plotting
from matplotlib.figure import Figure
import matplotlib.patches as patches
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backend_bases import MouseEvent
from signal_manager import global_signal_manager

class Thickness_evaluation(QWidget):
    def __init__(self):
        super().__init__()
        # dialog layout
        self.TK_layout = QGridLayout()


        # 標準觀察者
        self.xy_n = [0.313, 0.3290]  # D65 光源的 CIE 1931 2度標準觀察者

        # TK_BLU
        self.TK_light_source_label = QLabel("Light_Source")
        self.TK_light_source_datatable = QComboBox()
        self.TK_light_source_datatable.setStyleSheet(QCOMBOXSETTING)
        self.TK_light_source_datatable.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.TK_light_source = QComboBox()
        self.TK_light_source.setStyleSheet(QCOMBOXSETTING)
        # 更新light_source_table
        self.update_TK_light_source_datatable()
        # 觸發table更新,lightsource表單
        self.updateTK_light_sourceComboBox()  # 初始化
        self.TK_light_source_datatable.currentIndexChanged.connect(self.updateTK_light_sourceComboBox)
        # # 連動觸發計算
        # self.TK_light_source_datatable.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_light_source.currentIndexChanged.connect(self.calculate_TK_Color)
        # layer1
        # self.TK_layer1_label = QLabel("Layer_1")
        self.TK_layer1_mode = QComboBox()
        self.TK_layer1_mode.addItem("Layer1_未選")
        self.TK_layer1_mode.addItem("layer1_自訂")
        self.TK_layer1_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.TK_layer1_box = QComboBox()
        self.TK_layer1_box.setStyleSheet(QCOMBOXSETTING)
        self.TK_layer1_table = QComboBox()
        self.TK_layer1_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_TK_layer1_datatable()
        self.updateTKlayer1ComboBox()  # 初始化
        self.TK_layer1_table.currentIndexChanged.connect(self.updateTKlayer1ComboBox)
        # # 連動觸發計算
        # self.TK_layer1_mode.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer1_box.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer1_table.currentIndexChanged.connect(self.calculate_TK_Color)
        # layer2
        # self.TK_layer2_label = QLabel("Layer_2")
        self.TK_layer2_mode = QComboBox()
        self.TK_layer2_mode.addItem("Layer2_未選")
        self.TK_layer2_mode.addItem("layer2_自訂")
        self.TK_layer2_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.TK_layer2_box = QComboBox()
        self.TK_layer2_box.setStyleSheet(QCOMBOXSETTING)
        self.TK_layer2_table = QComboBox()
        self.TK_layer2_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_TK_layer2_datatable()
        self.updateTKlayer2ComboBox()  # 初始化
        self.TK_layer2_table.currentIndexChanged.connect(self.updateTKlayer2ComboBox)
        # # 連動觸發計算
        # self.TK_layer2_mode.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer2_box.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer2_table.currentIndexChanged.connect(self.calculate_TK_Color)
        # layer3
        # self.TK_layer3_label = QLabel("Layer_3")
        self.TK_layer3_mode = QComboBox()
        self.TK_layer3_mode.addItem("Layer3_未選")
        self.TK_layer3_mode.addItem("layer3_自訂")
        self.TK_layer3_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.TK_layer3_box = QComboBox()
        self.TK_layer3_box.setStyleSheet(QCOMBOXSETTING)
        self.TK_layer3_table = QComboBox()
        self.TK_layer3_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_TK_layer3_datatable()
        self.updateTKlayer3ComboBox()  # 初始化
        self.TK_layer3_table.currentIndexChanged.connect(self.updateTKlayer3ComboBox)
        # # cal
        # self.TK_layer3_mode.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer3_box.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer3_table.currentIndexChanged.connect(self.calculate_TK_Color)
        # layer4
        # self.TK_layer4_label = QLabel("Layer_4")
        self.TK_layer4_mode = QComboBox()
        self.TK_layer4_mode.addItem("Layer4_未選")
        self.TK_layer4_mode.addItem("layer4_自訂")
        self.TK_layer4_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.TK_layer4_box = QComboBox()
        self.TK_layer4_box.setStyleSheet(QCOMBOXSETTING)
        self.TK_layer4_table = QComboBox()
        self.TK_layer4_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_TK_layer4_datatable()
        self.updateTKlayer4ComboBox()  # 初始化
        self.TK_layer4_table.currentIndexChanged.connect(self.updateTKlayer4ComboBox)
        # # cal
        # self.TK_layer4_mode.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer4_box.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer4_table.currentIndexChanged.connect(self.calculate_TK_Color)

        # layer5
        # self.TK_layer5_label = QLabel("Layer_5")
        self.TK_layer5_mode = QComboBox()
        self.TK_layer5_mode.addItem("Layer5_未選")
        self.TK_layer5_mode.addItem("layer5_自訂")
        self.TK_layer5_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.TK_layer5_box = QComboBox()
        self.TK_layer5_box.setStyleSheet(QCOMBOXSETTING)
        self.TK_layer5_table = QComboBox()
        self.TK_layer5_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_TK_layer5_datatable()
        self.updateTKlayer5ComboBox()  # 初始化
        self.TK_layer5_table.currentIndexChanged.connect(self.updateTKlayer5ComboBox)
        # # cal
        # self.TK_layer5_mode.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer5_box.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer5_table.currentIndexChanged.connect(self.calculate_TK_Color)
        # layer6
        # self.TK_layer6_label = QLabel("Layer_6")
        self.TK_layer6_mode = QComboBox()
        self.TK_layer6_mode.addItem("Layer6_未選")
        self.TK_layer6_mode.addItem("layer6_自訂")
        self.TK_layer6_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.TK_layer6_box = QComboBox()
        self.TK_layer6_box.setStyleSheet(QCOMBOXSETTING)
        self.TK_layer6_table = QComboBox()
        self.TK_layer6_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_TK_layer6_datatable()
        self.updateTKlayer6ComboBox()  # 初始化
        self.TK_layer6_table.currentIndexChanged.connect(self.updateTKlayer6ComboBox)
        # # cal
        # self.TK_layer6_mode.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer6_box.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_layer6_table.currentIndexChanged.connect(self.calculate_TK_Color)

        # RCF
        self.TK_RCF_Change_label = QLabel("RCF_Thickness")
        self.TK_RCF_mode_option = QComboBox()
        self.TK_RCF_mode_option.setStyleSheet(QCOMBOXMODESETTING)
        self.TK_RCF_mode_option.addItem("RCF_Change")
        self.TK_RCF_mode_option.addItem("RCF_Differ")
        # initial
        self.TK_RCF_table = QComboBox()
        self.TK_RCF_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.TK_RCF_Range = QLabel("TK_Range")
        self.TK_RCF_Combobox = QComboBox()
        self.TK_RCF_Combobox.setStyleSheet(QCOMBOXSETTING)
        self.TK_RCF_Strat = QLineEdit()
        self.TK_RCF_Strat.setPlaceholderText("TK_Start")
        self.TK_RCF_End = QLineEdit()
        self.TK_RCF_End.setPlaceholderText("TK_End")
        self.TK_RCF_interval = QLineEdit()
        self.TK_RCF_interval.setPlaceholderText("Interval")
        # 初始化
        self.update_TK_RCF_datatable()
        self.update_TK_RCF_ComboBox()
        # 連動更新
        self.TK_RCF_mode_option.currentIndexChanged.connect(self.update_TK_RCF_datatable)
        self.TK_RCF_mode_option.currentIndexChanged.connect(self.update_TK_RCF_ComboBox)
        self.TK_RCF_table.currentIndexChanged.connect(self.update_TK_RCF_ComboBox)
        # # cal

        self.TK_RCF_mode_option.currentIndexChanged.connect(self.show_TK_RCF)
        self.TK_RCF_table.currentIndexChanged.connect(self.show_TK_RCF)
        self.TK_RCF_Combobox.currentIndexChanged.connect(self.show_TK_RCF)
        # self.TK_RCF_Strat.textChanged.connect(self.calculate_TK_Color)
        # self.TK_RCF_End.textChanged.connect(self.calculate_TK_Color)
        # self.TK_RCF_interval.textChanged.connect(self.calculate_TK_Color)

        self.TK_RCF_Table_Form = QTableWidget()
        self.TK_RCF_Table_Form.setColumnCount(6)
        self.TK_RCF_Table_Form.setRowCount(20)
        self.TK_RCF_Table_Form.setFixedSize(400, 200)
        TK_items = ["Item", "Rx", "Ry", "RTK","R_λ","Purity"]
        TK_colors = [QColor(RESULTRED), QColor(RESULTRED), QColor(RESULTRED), QColor(RESULTRED),QColor(RESULTRED),QColor(RESULTRED)]
        for column, (TK_item, TK_color) in enumerate(zip(TK_items, TK_colors)):
            item = QTableWidgetItem(TK_item)
            item.setBackground(TK_color)
            item.setTextAlignment(Qt.AlignCenter)
            self.TK_RCF_Table_Form.setItem(0, column, item)

        # Table style
        self.TK_RCF_Table_Form.setStyleSheet("""
                                                QTableWidget::item:selected {
                                            color: blcak; /* 設定文字顏色為黑色 */
                                            background-color: #008080; /* 設定背景顏色為藍色，你可以根據需要調整 */
                                                }
                                            """)

        # 創建一個通用的 Ctrl+C 快捷鍵
        copy_shortcut = QShortcut(QKeySequence.Copy, self)
        copy_shortcut.activated.connect(self.copy_table_content)


        # RCF Picture
        # 圖窗創建
        self.CIE_R_Figure = Figure(figsize=(9, 6))
        # 創立CIE畫布
        self.CIE_R_canvas = FigureCanvas(self.CIE_R_Figure)

        # 畫布放置
        self.TK_layout.addWidget(self.CIE_R_canvas,4,6,2,3)

        # 創立建圖物件
        self.CIE_R_chart = self.CIE_R_canvas.figure.add_subplot(111)

        # GCF
        self.TK_GCF_Change_label = QLabel("GCF_Thickness")
        self.TK_GCF_mode_option = QComboBox()
        self.TK_GCF_mode_option.setStyleSheet(QCOMBOXMODESETTING)
        self.TK_GCF_mode_option.addItem("GCF_Change")
        self.TK_GCF_mode_option.addItem("GCF_Differ")
        # initial
        self.TK_GCF_table = QComboBox()
        self.TK_GCF_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.TK_GCF_Range = QLabel("TK_Range")
        self.TK_GCF_Combobox = QComboBox()
        self.TK_GCF_Combobox.setStyleSheet(QCOMBOXSETTING)
        self.TK_GCF_Strat = QLineEdit()
        self.TK_GCF_Strat.setPlaceholderText("TK_Start")
        self.TK_GCF_End = QLineEdit()
        self.TK_GCF_End.setPlaceholderText("TK_End")
        self.TK_GCF_interval = QLineEdit()
        self.TK_GCF_interval.setPlaceholderText("Interval")
        # 初始化
        self.update_TK_GCF_datatable()
        self.update_TK_GCF_ComboBox()
        # 連動更新
        self.TK_GCF_mode_option.currentIndexChanged.connect(self.update_TK_GCF_datatable)
        self.TK_GCF_mode_option.currentIndexChanged.connect(self.update_TK_GCF_ComboBox)
        self.TK_GCF_table.currentIndexChanged.connect(self.update_TK_GCF_ComboBox)
        # # cal
        self.TK_GCF_mode_option.currentIndexChanged.connect(self.show_TK_GCF)
        self.TK_GCF_table.currentIndexChanged.connect(self.show_TK_GCF)
        self.TK_GCF_Combobox.currentIndexChanged.connect(self.show_TK_GCF)
        # self.TK_GCF_table.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_GCF_Strat.textChanged.connect(self.calculate_TK_Color)
        # self.TK_GCF_End.textChanged.connect(self.calculate_TK_Color)
        # self.TK_GCF_interval.textChanged.connect(self.calculate_TK_Color)

        self.TK_GCF_Table_Form = QTableWidget()
        self.TK_GCF_Table_Form.setColumnCount(6)
        self.TK_GCF_Table_Form.setRowCount(20)
        self.TK_GCF_Table_Form.setFixedSize(400, 200)
        TK_items = ["Item", "Gx", "Gy", "GTK", "G_λ", "Purity"]
        TK_colors = [QColor(RESULTGREEN), QColor(RESULTGREEN), QColor(RESULTGREEN), QColor(RESULTGREEN), QColor(RESULTGREEN),
                     QColor(RESULTGREEN)]
        for column, (TK_item, TK_color) in enumerate(zip(TK_items, TK_colors)):
            item = QTableWidgetItem(TK_item)
            item.setBackground(TK_color)
            item.setTextAlignment(Qt.AlignCenter)
            self.TK_GCF_Table_Form.setItem(0, column, item)

        # Table style
        self.TK_GCF_Table_Form.setStyleSheet("""
                                                        QTableWidget::item:selected {
                                                    color: blcak; /* 設定文字顏色為黑色 */
                                                    background-color: #008080; /* 設定背景顏色為藍色，你可以根據需要調整 */
                                                        }
                                                    """)



        # RCF Picture
        # 圖窗創建
        self.CIE_G_Figure = Figure(figsize=(9, 6))
        # 創立CIE畫布
        self.CIE_G_canvas = FigureCanvas(self.CIE_G_Figure)

        # 畫布放置
        self.TK_layout.addWidget(self.CIE_G_canvas, 6, 6, 2, 3)

        # 創立建圖物件
        self.CIE_G_chart = self.CIE_G_canvas.figure.add_subplot(111)

        # BCF
        self.TK_BCF_Change_label = QLabel("BCF_Thickness")
        self.TK_BCF_mode_option = QComboBox()
        self.TK_BCF_mode_option.setStyleSheet(QCOMBOXMODESETTING)
        self.TK_BCF_mode_option.addItem("BCF_Change")
        self.TK_BCF_mode_option.addItem("BCF_Differ")
        # initial
        self.TK_BCF_table = QComboBox()
        self.TK_BCF_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.TK_BCF_Range = QLabel("TK_Range")
        self.TK_BCF_Combobox = QComboBox()
        self.TK_BCF_Combobox.setStyleSheet(QCOMBOXSETTING)
        self.TK_BCF_Strat = QLineEdit()
        self.TK_BCF_Strat.setPlaceholderText("TK_Start")
        self.TK_BCF_End = QLineEdit()
        self.TK_BCF_End.setPlaceholderText("TK_End")
        self.TK_BCF_interval = QLineEdit()
        self.TK_BCF_interval.setPlaceholderText("Interval")
        # 初始化
        self.update_TK_BCF_datatable()
        self.update_TK_BCF_ComboBox()
        # 連動更新
        self.TK_BCF_mode_option.currentIndexChanged.connect(self.update_TK_BCF_datatable)
        self.TK_BCF_mode_option.currentIndexChanged.connect(self.update_TK_BCF_ComboBox)
        self.TK_BCF_table.currentIndexChanged.connect(self.update_TK_BCF_ComboBox)
        # # cal
        self.TK_BCF_mode_option.currentIndexChanged.connect(self.show_TK_BCF)
        self.TK_BCF_table.currentIndexChanged.connect(self.show_TK_BCF)
        self.TK_BCF_Combobox.currentIndexChanged.connect(self.show_TK_BCF)
        # self.TK_BCF_table.currentIndexChanged.connect(self.calculate_TK_Color)
        # self.TK_BCF_Strat.textChanged.connect(self.calculate_TK_Color)
        # self.TK_BCF_End.textChanged.connect(self.calculate_TK_Color)
        # self.TK_BCF_interval.textChanged.connect(self.calculate_TK_Color)

        self.TK_BCF_Table_Form = QTableWidget()
        self.TK_BCF_Table_Form.setColumnCount(6)
        self.TK_BCF_Table_Form.setRowCount(20)
        self.TK_BCF_Table_Form.setFixedSize(400, 200)
        TK_items = ["Item", "Bx", "By", "BTK", "B_λ", "Purity"]
        TK_colors = [QColor(RESULTBLUE), QColor(RESULTBLUE), QColor(RESULTBLUE), QColor(RESULTBLUE), QColor(RESULTBLUE),
                     QColor(RESULTBLUE)]
        for column, (TK_item, TK_color) in enumerate(zip(TK_items, TK_colors)):
            item = QTableWidgetItem(TK_item)
            item.setBackground(TK_color)
            item.setTextAlignment(Qt.AlignCenter)
            self.TK_BCF_Table_Form.setItem(0, column, item)

        # Table style
        self.TK_BCF_Table_Form.setStyleSheet("""
                                                        QTableWidget::item:selected {
                                                    color: blcak; /* 設定文字顏色為黑色 */
                                                    background-color: #008080; /* 設定背景顏色為藍色，你可以根據需要調整 */
                                                        }
                                                    """)



        # RCF Picture
        # 圖窗創建
        self.CIE_B_Figure = Figure(figsize=(9, 6))
        # 創立CIE畫布
        self.CIE_B_canvas = FigureCanvas(self.CIE_B_Figure)

        # 畫布放置
        self.TK_layout.addWidget(self.CIE_B_canvas, 8, 6, 2, 3)

        # 創立建圖物件
        self.CIE_B_chart = self.CIE_B_canvas.figure.add_subplot(111)

        # cal button
        self.TK_test_button = QPushButton("Calculate_all")
        self.TK_test_button.clicked.connect(self.calculae_TK_Color_Twice)
        # 設置樣式表
        self.TK_test_button.setStyleSheet("QPushButton {"
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
        self.TK_test_button.setCursor(Qt.PointingHandCursor)  # 手指形狀

        # 放置
        self.TK_layout.addWidget(self.TK_light_source_label, 0, 0)
        self.TK_layout.addWidget(self.TK_light_source_datatable, 0, 1)
        self.TK_layout.addWidget(self.TK_light_source, 1, 0, 1, 2)
        # layer
        self.TK_layout.addWidget(self.TK_layer1_mode, 0, 2)
        self.TK_layout.addWidget(self.TK_layer1_table, 0, 3)
        self.TK_layout.addWidget(self.TK_layer1_box, 1, 2, 1, 2)

        self.TK_layout.addWidget(self.TK_layer2_mode, 0, 4)
        self.TK_layout.addWidget(self.TK_layer2_table, 0, 5)
        self.TK_layout.addWidget(self.TK_layer2_box, 1, 4, 1, 2)

        self.TK_layout.addWidget(self.TK_layer3_mode, 0, 6)
        self.TK_layout.addWidget(self.TK_layer3_table, 0, 7)
        self.TK_layout.addWidget(self.TK_layer3_box, 1, 6, 1, 2)

        self.TK_layout.addWidget(self.TK_layer4_mode, 2, 0)
        self.TK_layout.addWidget(self.TK_layer4_table, 2, 1)
        self.TK_layout.addWidget(self.TK_layer4_box, 3, 0, 1, 2)

        self.TK_layout.addWidget(self.TK_layer5_mode, 2, 2)
        self.TK_layout.addWidget(self.TK_layer5_table, 2, 3)
        self.TK_layout.addWidget(self.TK_layer5_box, 3, 2, 1, 2)

        self.TK_layout.addWidget(self.TK_layer6_mode, 2, 4)
        self.TK_layout.addWidget(self.TK_layer6_table, 2, 5)
        self.TK_layout.addWidget(self.TK_layer6_box, 3, 4, 1, 2)
        # RCF
        self.TK_layout.addWidget(self.TK_RCF_Change_label, 4, 0)
        self.TK_layout.addWidget(self.TK_RCF_mode_option, 4, 1)
        self.TK_layout.addWidget(self.TK_RCF_table, 4, 2)
        self.TK_layout.addWidget(self.TK_RCF_Range, 4, 3)
        self.TK_layout.addWidget(self.TK_RCF_Combobox, 5, 0)
        self.TK_layout.addWidget(self.TK_RCF_Strat, 5, 1)
        self.TK_layout.addWidget(self.TK_RCF_End, 5, 2)
        self.TK_layout.addWidget(self.TK_RCF_interval, 5, 3)
        self.TK_layout.addWidget(self.TK_RCF_Table_Form, 4, 4, 2, 2)

        # GCF
        self.TK_layout.addWidget(self.TK_GCF_Change_label, 6, 0)
        self.TK_layout.addWidget(self.TK_GCF_mode_option, 6, 1)
        self.TK_layout.addWidget(self.TK_GCF_table, 6, 2)
        self.TK_layout.addWidget(self.TK_GCF_Range, 6, 3)
        self.TK_layout.addWidget(self.TK_GCF_Combobox, 7, 0)
        self.TK_layout.addWidget(self.TK_GCF_Strat, 7, 1)
        self.TK_layout.addWidget(self.TK_GCF_End, 7, 2)
        self.TK_layout.addWidget(self.TK_GCF_interval, 7, 3)
        self.TK_layout.addWidget(self.TK_GCF_Table_Form, 6, 4, 2, 2)

        # BCF
        self.TK_layout.addWidget(self.TK_BCF_Change_label, 8, 0)
        self.TK_layout.addWidget(self.TK_BCF_mode_option, 8, 1)
        self.TK_layout.addWidget(self.TK_BCF_table, 8, 2)
        self.TK_layout.addWidget(self.TK_BCF_Range, 8, 3)
        self.TK_layout.addWidget(self.TK_BCF_Combobox, 9, 0)
        self.TK_layout.addWidget(self.TK_BCF_Strat, 9, 1)
        self.TK_layout.addWidget(self.TK_BCF_End, 9, 2)
        self.TK_layout.addWidget(self.TK_BCF_interval, 9, 3)
        self.TK_layout.addWidget(self.TK_BCF_Table_Form, 8, 4, 2, 2)

        # test
        self.TK_layout.addWidget(self.TK_test_button, 2, 6,2,2)
        # copy button---------------------------------------------------
        self.copy_R_form = QPushButton("R_Form_Copy")
        self.copy_G_form = QPushButton("G_Form_Copy")
        self.copy_B_form = QPushButton("B_Form_Copy")
        self.TK_layout.addWidget(self.copy_R_form, 0, 8)
        self.TK_layout.addWidget(self.copy_G_form, 1, 8)
        self.TK_layout.addWidget(self.copy_B_form, 2, 8)

        self.copy_R_form.clicked.connect(self.copy_R_table_content)
        self.copy_G_form.clicked.connect(self.copy_G_table_content)
        self.copy_B_form.clicked.connect(self.copy_B_table_content)

        self.TK_layout.setSpacing(1)

        self.setLayout(self.TK_layout)
        # # 連接全局信號到相應的方法
        global_signal_manager.databaseUpdated.connect(self.update_TK_light_source_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateTK_light_sourceComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_TK_layer1_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateTKlayer1ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_TK_layer2_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateTKlayer2ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_TK_layer3_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateTKlayer3ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_TK_layer4_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateTKlayer4ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_TK_layer5_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateTKlayer5ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_TK_layer6_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateTKlayer6ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_TK_RCF_datatable)
        global_signal_manager.databaseUpdated.connect(self.update_TK_RCF_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_TK_GCF_datatable)
        global_signal_manager.databaseUpdated.connect(self.update_TK_GCF_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_TK_BCF_datatable)
        global_signal_manager.databaseUpdated.connect(self.update_TK_BCF_ComboBox)

#-------------TK區------------------------------------------------------------------------------------------

    # 複製整個表格的內容
    def copy_R_table_content(self):
        row_count = self.TK_RCF_Table_Form.rowCount()
        column_count = self.TK_RCF_Table_Form.columnCount()

        copied_data = []
        for row in range(row_count):
            row_data = []
            for col in range(column_count):
                item = self.TK_RCF_Table_Form.item(row, col)
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

    def copy_G_table_content(self):
        row_count = self.TK_GCF_Table_Form.rowCount()
        column_count = self.TK_GCF_Table_Form.columnCount()

        copied_data = []
        for row in range(row_count):
            row_data = []
            for col in range(column_count):
                item = self.TK_GCF_Table_Form.item(row, col)
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

    def copy_B_table_content(self):
        row_count = self.TK_BCF_Table_Form.rowCount()
        column_count = self.TK_BCF_Table_Form.columnCount()

        copied_data = []
        for row in range(row_count):
            row_data = []
            for col in range(column_count):
                item = self.TK_BCF_Table_Form.item(row, col)
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
                print("header_name", header_name)

                # 尋找最後一個 "_" 的位置
                underscore_index = header_name.rfind('_')
                # 如果存在 "_" 且不在開頭，則截取到這個位置
                if underscore_index > 0:
                    header_name = header_name[:underscore_index]

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
                print("header_name", header_name)

                # 尋找最後一個 "_" 的位置
                underscore_index = header_name.rfind('_')
                # 如果存在 "_" 且不在開頭，則截取到這個位置
                if underscore_index > 0:
                    header_name = header_name[:underscore_index]

                if header_name not in added_items:
                    self.TK_RCF_Combobox.addItem(str(header_name))
                    added_items.add(header_name)
                    print("item", str(header_name))
            # 關閉連線
            connection.close()

    # TK GCF
    def update_TK_GCF_datatable(self):
        if self.TK_GCF_mode_option.currentText() == "GCF_Change":
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
            self.TK_GCF_table.clear()
            for table in tables:
                self.TK_GCF_table.addItem(table[0])

            # 關閉連線
            conn.close()
        else:
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
            self.TK_GCF_table.clear()
            for table in tables:
                self.TK_GCF_table.addItem(table[0])

            # 關閉連線
            conn.close()

    def update_TK_GCF_ComboBox(self):
        if self.TK_GCF_mode_option.currentText() == "GCF_Change":
            connection = sqlite3.connect("GCF_Change_spectrum.db")
            cursor = connection.cursor()

            # 获取表格的標題
            cursor.execute(f"PRAGMA table_info('{self.TK_GCF_table.currentText()}');")
            header_data = cursor.fetchall()
            header_labels = [column[1] for column in header_data]
            # print("headerlabels-from source", header_labels)
            self.TK_GCF_Combobox.clear()
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
                    self.TK_GCF_Combobox.addItem(str(header_name))
                    added_items.add(header_name)
                    print("item", str(header_name))
            # 關閉連線
            connection.close()
        else:
            connection = sqlite3.connect("GCF_Differ_spectrum.db")
            cursor = connection.cursor()

            # 获取表格的標題
            cursor.execute(f"PRAGMA table_info('{self.TK_GCF_table.currentText()}');")
            header_data = cursor.fetchall()
            header_labels = [column[1] for column in header_data]
            # print("headerlabels-from source", header_labels)
            self.TK_GCF_Combobox.clear()
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
                    self.TK_GCF_Combobox.addItem(str(header_name))
                    added_items.add(header_name)
                    print("item", str(header_name))
            # 關閉連線
            connection.close()

    # TK BCF
    def update_TK_BCF_datatable(self):
        if self.TK_BCF_mode_option.currentText() == "BCF_Change":
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
            self.TK_BCF_table.clear()
            for table in tables:
                self.TK_BCF_table.addItem(table[0])

            # 關閉連線
            conn.close()
        else:
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
            self.TK_BCF_table.clear()
            for table in tables:
                self.TK_BCF_table.addItem(table[0])

            # 關閉連線
            conn.close()

    def update_TK_BCF_ComboBox(self):
        if self.TK_BCF_mode_option.currentText() == "BCF_Change":
            connection = sqlite3.connect("BCF_Change_spectrum.db")
            cursor = connection.cursor()

            # 获取表格的標題
            cursor.execute(f"PRAGMA table_info('{self.TK_BCF_table.currentText()}');")
            header_data = cursor.fetchall()
            header_labels = [column[1] for column in header_data]
            # print("headerlabels-from source", header_labels)
            self.TK_BCF_Combobox.clear()
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
                    self.TK_BCF_Combobox.addItem(str(header_name))
                    added_items.add(header_name)
                    print("item", str(header_name))
            # 關閉連線
            connection.close()
        else:
            connection = sqlite3.connect("BCF_Differ_spectrum.db")
            cursor = connection.cursor()

            # 获取表格的標題
            cursor.execute(f"PRAGMA table_info('{self.TK_BCF_table.currentText()}');")
            header_data = cursor.fetchall()
            header_labels = [column[1] for column in header_data]
            # print("headerlabels-from source", header_labels)
            self.TK_BCF_Combobox.clear()
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
                    self.TK_BCF_Combobox.addItem(str(header_name))
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
            self.Sort_TK_SP_list = sorted(TK_SP_list,reverse=True)
            print("self.Sort_TK_SP_list",self.Sort_TK_SP_list)
            self.TK_RCF_Range.setText(f"TKrange: {min(TK_SP_list):.2f}~{max(TK_SP_list):.2f}")

            # 反轉找回原本的順序list
            Re_list = []
            for i in range(len(indices_without_number)):
                Re_list.append(TK_SP_list.index(self.Sort_TK_SP_list[i]))
            print("Re_list",Re_list)

            A_list = []
            for i in range(len(indices_without_number)-1):
                A_list.append((-1/(float(self.Sort_TK_SP_list[i+1]) - float(self.Sort_TK_SP_list[i])) * np.log(SP_series_list[Re_list[i+1]][1:]/SP_series_list[Re_list[i]][1:])).reset_index(drop=True))
            #print("A_list",A_list)
            K_list = []
            for i in range(len(indices_without_number)-1):
                K_list.append(SP_series_list[Re_list[i]][1:].reset_index(drop=True)/np.exp((-1 * A_list[i])* self.Sort_TK_SP_list[i]))
            print("K_list",K_list)
            self.RCF_TK_list = []
            try:
                TK_RCF_Start = float(self.TK_RCF_Strat.text())
                TK_RCF_End = float(self.TK_RCF_End.text())
                TK_RCF_Interval = float(self.TK_RCF_interval.text())
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
                    print("self.RCF_TK_list", self.RCF_TK_list)
            except ValueError:
                print("R輸入的數值格式為空，請檢查後重新輸入。")

            AK_RCF_Change_spectrum_Series_list = []
            for RCF_TK in self.RCF_TK_list:
                # 先将 R_aK_TK 添加到 self.Sort_TK_SP_list
                self.Sort_TK_SP_list.append(RCF_TK)
                # 然后对更新后的列表进行排序
                Sort_TK_SP_AK_list = sorted(self.Sort_TK_SP_list, reverse=True)
                print("Judge")
                print("Sort_TK_SP_AK_list", Sort_TK_SP_AK_list)
                print("RCF_TK",RCF_TK)
                # 4種情況
                if RCF_TK == max(Sort_TK_SP_AK_list):
                    print("max")
                    AK_RCF_Change_spectrum_Series = K_list[0] * np.exp(-1 * A_list[0] * RCF_TK)
                    AK_RCF_Change_spectrum_Series_list.append(AK_RCF_Change_spectrum_Series)
                    # print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # return AK_RCF_Change_spectrum_Series
                elif RCF_TK == min(Sort_TK_SP_AK_list):
                    print("min")
                    print("K_list_min", K_list)
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
                    print("RCF_TK",RCF_TK)
                    print("position_mid",position)
                    print("Sort_TK_SP_AK_list",Sort_TK_SP_AK_list)
                    print("K_list_mid",K_list)
                    AK_RCF_Change_spectrum_Series = K_list[position-1] * np.exp(-1 * A_list[position-1] * RCF_TK)
                    AK_RCF_Change_spectrum_Series_list.append(AK_RCF_Change_spectrum_Series)
                    # print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # return AK_RCF_Change_spectrum_Series
                self.Sort_TK_SP_list.pop()
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
            try:
                TK_RCF_Start = float(self.TK_RCF_Strat.text())
                TK_RCF_End = float(self.TK_RCF_End.text())
                TK_RCF_Interval = float(self.TK_RCF_interval.text())

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
            except ValueError:
                # 如果轉換出錯，打印錯誤消息
                print("R輸入的數值格式不正確，請檢查後重新輸入。")

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

    def calculate_TK_GCF(self):
        if self.TK_GCF_mode_option.currentText() == "GCF_Change":
            print("in_GCF_Change自訂")
            connection_GCF_Change = sqlite3.connect("GCF_Change_spectrum.db")
            cursor_GCF_Change = connection_GCF_Change.cursor()
            # 取得Rdiffer資料
            column_name_GCF_Change = self.TK_GCF_Combobox.currentText()
            print("column_name_GCF_Change",column_name_GCF_Change)
            table_name_GCF_Change = self.TK_GCF_table.currentText()
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
            print("SP_series_list",SP_series_list)

            # TK_SP準備分厚度順序
            TK_SP_list = []
            for i in range(len(indices_without_number)):
                TK_SP_list.append(SP_series_list[i].iloc[0])
            print("TK_SP_list",TK_SP_list)
            self.Sort_TK_SP_list = sorted(TK_SP_list,reverse=True)
            print("self.Sort_TK_SP_list",self.Sort_TK_SP_list)
            self.TK_GCF_Range.setText(f"TKrange: {min(TK_SP_list):.2f}~{max(TK_SP_list):.2f}")

            # 反轉找回原本的順序list
            Re_list = []
            for i in range(len(indices_without_number)):
                Re_list.append(TK_SP_list.index(self.Sort_TK_SP_list[i]))
            #print("Re_list",Re_list)

            A_list = []
            for i in range(len(indices_without_number)-1):
                A_list.append((-1/(float(self.Sort_TK_SP_list[i+1]) - float(self.Sort_TK_SP_list[i])) * np.log(SP_series_list[Re_list[i+1]][1:]/SP_series_list[Re_list[i]][1:])).reset_index(drop=True))
            #print("A_list",A_list)
            K_list = []
            for i in range(len(indices_without_number)-1):
                K_list.append(SP_series_list[Re_list[i]][1:].reset_index(drop=True)/np.exp((-1 * A_list[i])* self.Sort_TK_SP_list[i]))
            #print("K_list",K_list)
            self.GCF_TK_list = []
            try:
                TK_GCF_Start = float(self.TK_GCF_Strat.text())
                TK_GCF_End = float(self.TK_GCF_End.text())
                TK_GCF_Interval = float(self.TK_GCF_interval.text())
                if float(self.TK_GCF_Strat.text()) < float(self.TK_GCF_End.text()):
                    i = 0
                    while float(self.TK_GCF_Strat.text()) + float(self.TK_GCF_interval.text()) * i < float(self.TK_GCF_End.text()):
                        self.GCF_TK_list.append(float(self.TK_GCF_Strat.text()) + float(self.TK_GCF_interval.text()) * i)
                        i += 1
                    self.GCF_TK_list.append(float(self.TK_GCF_End.text()))

                elif float(self.TK_GCF_End.text()) < float(self.TK_GCF_Strat.text()):
                    i = 0
                    while float(self.TK_GCF_End.text()) + float(self.TK_GCF_interval.text()) * i < float(self.TK_GCF_Strat.text()):
                        self.GCF_TK_list.append(float(self.TK_GCF_End.text()) + float(self.TK_GCF_interval.text()) * i)
                        i += 1
                    self.GCF_TK_list.append(float(self.TK_GCF_Strat.text()))
            except ValueError:
                # 如果轉換出錯，打印錯誤消息
                print("G輸入的數值格式不正確，請檢查後重新輸入。")
            print("self.GCF_TK_list",self.GCF_TK_list)

            AK_GCF_Change_spectrum_Series_list = []
            for GCF_TK in self.GCF_TK_list:
                # 先将 R_aK_TK 添加到 self.Sort_TK_SP_list
                self.Sort_TK_SP_list.append(GCF_TK)
                # 然后对更新后的列表进行排序
                Sort_TK_SP_AK_list = sorted(self.Sort_TK_SP_list, reverse=True)
                print("Sort_TK_SP_AK_list", Sort_TK_SP_AK_list)
                # 4種情況
                if GCF_TK == max(Sort_TK_SP_AK_list):
                    print("max")
                    AK_GCF_Change_spectrum_Series = K_list[0] * np.exp(-1 * A_list[0] * GCF_TK)
                    AK_GCF_Change_spectrum_Series_list.append(AK_GCF_Change_spectrum_Series)
                    # print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                    # return AK_GCF_Change_spectrum_Series
                elif GCF_TK == min(Sort_TK_SP_AK_list):
                    print("min")
                    AK_GCF_Change_spectrum_Series = K_list[len(K_list)-1] * np.exp(-1 * A_list[len(A_list)-1] * GCF_TK)
                    AK_GCF_Change_spectrum_Series_list.append(AK_GCF_Change_spectrum_Series)
                    # print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                    # return AK_GCF_Change_spectrum_Series
                elif GCF_TK in TK_SP_list:
                    print("equal")
                    AK_GCF_Change_spectrum_Series = SP_series_list[TK_SP_list.index(GCF_TK)][1:].reset_index(drop=True)
                    AK_GCF_Change_spectrum_Series_list.append(AK_GCF_Change_spectrum_Series)
                    # print("AK_GCF_Change_spectrum_Series",AK_GCF_Change_spectrum_Series)
                    # return AK_GCF_Change_spectrum_Series
                else:
                    print("mid")
                    position = Sort_TK_SP_AK_list.index(GCF_TK)
                    AK_GCF_Change_spectrum_Series = K_list[position-1] * np.exp(-1 * A_list[position-1] * GCF_TK)
                    AK_GCF_Change_spectrum_Series_list.append(AK_GCF_Change_spectrum_Series)
                    # print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                    # return AK_GCF_Change_spectrum_Series
                # 避免重複加入值判斷
                self.Sort_TK_SP_list.pop()
            print("AK_GCF_Change_spectrum_Series_list",AK_GCF_Change_spectrum_Series_list)
            for row,TK in enumerate(self.GCF_TK_list):
                # item = QTableWidgetItem(str(TK))
                self.TK_GCF_Table_Form.setItem(row + 1, 3, QTableWidgetItem(f"{TK:.3f}"))

            # 關閉連線
            connection_GCF_Change.close()

            return AK_GCF_Change_spectrum_Series_list
        else:
            connection_GCF_Differ = sqlite3.connect("GCF_Differ_spectrum.db")
            cursor_GCF_Differ = connection_GCF_Differ.cursor()
            # 取得Rdiffer資料
            column_name_GCF_Differ = self.TK_GCF_Combobox.currentText()
            # print("column_name_GCF_Differ",column_name_GCF_Differ)
            table_name_GCF_Differ = self.TK_GCF_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_GCF_Differ = f"SELECT * FROM '{table_name_GCF_Differ}';"
            cursor_GCF_Differ.execute(query_GCF_Differ)
            result_GCF_Differ = cursor_GCF_Differ.fetchall()

            # 找到指定標題的欄位索引
            header_GCF_Differ = [column[0] for column in cursor_GCF_Differ.description]

            # 移除所有標題中的編號
            header_GCF_Differ = [re.sub(r'_\d+$', '', col) for col in header_GCF_Differ]
            # print("header_GCF_Differ", header_GCF_Differ)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_GCF_Differ) if col == column_name_GCF_Differ]
            # print("indices_without_number", indices_without_number)

            # 使用 pandas Series 存儲結果
            series_list = [pd.Series([row[i] for row in result_GCF_Differ]) for i in indices_without_number]

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
            self.TK_GCF_Range.setText(f"TKrange: {min(TK_record):.2f}~{max(TK_record):.2f}")
            self.GCF_TK_list = []
            try:
                TK_GCF_Start = float(self.TK_GCF_Strat.text())
                TK_GCF_End = float(self.TK_GCF_End.text())
                TK_GCF_Interval = float(self.TK_GCF_interval.text())
                if float(self.TK_GCF_Strat.text()) < float(self.TK_GCF_End.text()):
                    i = 0
                    while float(self.TK_GCF_Strat.text()) + float(self.TK_GCF_interval.text()) * i < float(self.TK_GCF_End.text()):
                        self.GCF_TK_list.append(float(self.TK_GCF_Strat.text()) + float(self.TK_GCF_interval.text()) * i)
                        i += 1
                    self.GCF_TK_list.append(float(self.TK_GCF_End.text()))

                elif float(self.TK_GCF_End.text()) < float(self.TK_GCF_Strat.text()):
                    i = 0
                    while float(self.TK_GCF_End.text()) + float(self.TK_GCF_interval.text()) * i < float(self.TK_GCF_Strat.text()):
                        self.GCF_TK_list.append(float(self.TK_GCF_End.text()) + float(self.TK_GCF_interval.text()) * i)
                        i += 1
                    self.GCF_TK_list.append(float(self.TK_GCF_Strat.text()))
            except ValueError:
                # 如果轉換出錯，打印錯誤消息
                print("G輸入的數值格式不正確，請檢查後重新輸入。")
            print("self.GCF_TK_list",self.GCF_TK_list)
            AK_GCF_Change_spectrum_Series_list = []
            for GCF_TK in self.GCF_TK_list:
                G_differ_TK = GCF_TK
                Thickness_list = [G_differ_TK]
                for i in range(series_list_length):
                    Thickness_list.append(float(series_list[i].iloc[0]))
                    TK_record.append(float(series_list[i].iloc[0]))
                Thickness_list_2 = sorted(Thickness_list)
                print("Thickness_list_2", Thickness_list_2)
                index_G_differ_TK = Thickness_list_2.index(G_differ_TK)
                print(Thickness_list_2.index(G_differ_TK))

                # 輸入厚度的三種情況
                if G_differ_TK == min(Thickness_list):
                    print("min")
                    A_index_value = Thickness_list_2[index_G_differ_TK + 1]
                    # print("A_index_value", A_index_value)
                    B_index_value = Thickness_list_2[index_G_differ_TK + 2]
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
                    C_Series = A_Series + (G_differ_TK - A_index_value) * (B_Series - A_Series) / (
                            B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    AK_GCF_Change_spectrum_Series_list.append(C_Series)
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
                    AK_GCF_Change_spectrum_Series_list.append(C_Series)
                elif G_differ_TK in (Thickness_list[1:]):
                    print("equal")
                    # 取得對應的索引
                    Thickness_list.remove(Thickness_list[0])
                    equal_index = Thickness_list.index(G_differ_TK)
                    # print("Equal Index:", equal_index)
                    # print("Equal Value:", G_differ_TK)
                    C_Series = series_list[equal_index][1:]
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    AK_GCF_Change_spectrum_Series_list.append(C_Series)
                else:
                    print("mid")
                    # print("G_differ_TK",G_differ_TK)
                    # print("Thickness_list[1:]",Thickness_list[1:])
                    A_index_value = Thickness_list_2[index_G_differ_TK - 1]
                    # print("A_index_value",A_index_value)
                    B_index_value = Thickness_list_2[index_G_differ_TK + 1]
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
                    C_Series = A_Series + (G_differ_TK - A_index_value) * (B_Series - A_Series) / (
                                B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    AK_GCF_Change_spectrum_Series_list.append(C_Series)
            # 關閉連線
            connection_GCF_Differ.close()
            for row,TK in enumerate(self.GCF_TK_list):
                # item = QTableWidgetItem(str(TK))
                self.TK_GCF_Table_Form.setItem(row + 1, 3, QTableWidgetItem(f"{TK:.3f}"))
            print("AK_GCF_Change_spectrum_Series_list",AK_GCF_Change_spectrum_Series_list)
            return AK_GCF_Change_spectrum_Series_list

    def calculate_TK_BCF(self):
        if self.TK_BCF_mode_option.currentText() == "BCF_Change":
            print("in_BCF_Change自訂")
            connection_BCF_Change = sqlite3.connect("BCF_Change_spectrum.db")
            cursor_BCF_Change = connection_BCF_Change.cursor()
            # 取得Rdiffer資料
            column_name_BCF_Change = self.TK_BCF_Combobox.currentText()
            print("column_name_BCF_Change",column_name_BCF_Change)
            table_name_BCF_Change = self.TK_BCF_table.currentText()
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
            print("SP_series_list",SP_series_list)

            # TK_SP準備分厚度順序
            TK_SP_list = []
            for i in range(len(indices_without_number)):
                TK_SP_list.append(SP_series_list[i].iloc[0])
            print("TK_SP_list",TK_SP_list)
            self.Sort_TK_SP_list = sorted(TK_SP_list,reverse=True)
            print("self.Sort_TK_SP_list",self.Sort_TK_SP_list)
            self.TK_BCF_Range.setText(f"TKrange: {min(TK_SP_list):.2f}~{max(TK_SP_list):.2f}")

            # 反轉找回原本的順序list
            Re_list = []
            for i in range(len(indices_without_number)):
                Re_list.append(TK_SP_list.index(self.Sort_TK_SP_list[i]))
            #print("Re_list",Re_list)

            A_list = []
            for i in range(len(indices_without_number)-1):
                A_list.append((-1/(float(self.Sort_TK_SP_list[i+1]) - float(self.Sort_TK_SP_list[i])) * np.log(SP_series_list[Re_list[i+1]][1:]/SP_series_list[Re_list[i]][1:])).reset_index(drop=True))
            #print("A_list",A_list)
            K_list = []
            for i in range(len(indices_without_number)-1):
                K_list.append(SP_series_list[Re_list[i]][1:].reset_index(drop=True)/np.exp((-1 * A_list[i])* self.Sort_TK_SP_list[i]))
            #print("K_list",K_list)
            self.BCF_TK_list = []
            try:
                TK_BCF_Start = float(self.TK_BCF_Strat.text())
                TK_BCF_End = float(self.TK_BCF_End.text())
                TK_BCF_Interval = float(self.TK_BCF_interval.text())
                if float(self.TK_BCF_Strat.text()) < float(self.TK_BCF_End.text()):
                    i = 0
                    while float(self.TK_BCF_Strat.text()) + float(self.TK_BCF_interval.text()) * i < float(self.TK_BCF_End.text()):
                        self.BCF_TK_list.append(float(self.TK_BCF_Strat.text()) + float(self.TK_BCF_interval.text()) * i)
                        i += 1
                    self.BCF_TK_list.append(float(self.TK_BCF_End.text()))

                elif float(self.TK_BCF_End.text()) < float(self.TK_BCF_Strat.text()):
                    i = 0
                    while float(self.TK_BCF_End.text()) + float(self.TK_BCF_interval.text()) * i < float(self.TK_BCF_Strat.text()):
                        self.BCF_TK_list.append(float(self.TK_BCF_End.text()) + float(self.TK_BCF_interval.text()) * i)
                        i += 1
                    self.BCF_TK_list.append(float(self.TK_BCF_Strat.text()))

            except ValueError:
                # 如果轉換出錯，打印錯誤消息
                print("B輸入的數值格式不正確，請檢查後重新輸入。")

            print("self.BCF_TK_list",self.BCF_TK_list)

            AK_BCF_Change_spectrum_Series_list = []
            for BCF_TK in self.BCF_TK_list:
                # 先将 R_aK_TK 添加到 self.Sort_TK_SP_list
                self.Sort_TK_SP_list.append(BCF_TK)
                # 然后对更新后的列表进行排序
                Sort_TK_SP_AK_list = sorted(self.Sort_TK_SP_list, reverse=True)
                print("Sort_TK_SP_AK_list", Sort_TK_SP_AK_list)
                # 4種情況
                if BCF_TK == max(Sort_TK_SP_AK_list):
                    print("max")
                    AK_BCF_Change_spectrum_Series = K_list[0] * np.exp(-1 * A_list[0] * BCF_TK)
                    AK_BCF_Change_spectrum_Series_list.append(AK_BCF_Change_spectrum_Series)
                    # print("AK_BCF_Change_spectrum_Series", AK_BCF_Change_spectrum_Series)
                    # return AK_BCF_Change_spectrum_Series
                elif BCF_TK == min(Sort_TK_SP_AK_list):
                    print("min")
                    AK_BCF_Change_spectrum_Series = K_list[len(K_list)-1] * np.exp(-1 * A_list[len(A_list)-1] * BCF_TK)
                    AK_BCF_Change_spectrum_Series_list.append(AK_BCF_Change_spectrum_Series)
                    # print("AK_BCF_Change_spectrum_Series", AK_BCF_Change_spectrum_Series)
                    # return AK_BCF_Change_spectrum_Series
                elif BCF_TK in TK_SP_list:
                    print("equal")
                    AK_BCF_Change_spectrum_Series = SP_series_list[TK_SP_list.index(BCF_TK)][1:].reset_index(drop=True)
                    AK_BCF_Change_spectrum_Series_list.append(AK_BCF_Change_spectrum_Series)
                    # print("AK_BCF_Change_spectrum_Series",AK_BCF_Change_spectrum_Series)
                    # return AK_BCF_Change_spectrum_Series
                else:
                    print("mid")
                    position = Sort_TK_SP_AK_list.index(BCF_TK)
                    AK_BCF_Change_spectrum_Series = K_list[position-1] * np.exp(-1 * A_list[position-1] * BCF_TK)
                    AK_BCF_Change_spectrum_Series_list.append(AK_BCF_Change_spectrum_Series)
                    # print("AK_BCF_Change_spectrum_Series", AK_BCF_Change_spectrum_Series)
                    # return AK_BCF_Change_spectrum_Series
                # 避免重複加入值判斷
                self.Sort_TK_SP_list.pop()
            print("AK_BCF_Change_spectrum_Series_list",AK_BCF_Change_spectrum_Series_list)
            for row,TK in enumerate(self.BCF_TK_list):
                # item = QTableWidgetItem(str(TK))
                self.TK_BCF_Table_Form.setItem(row + 1, 3, QTableWidgetItem(f"{TK:.3f}"))

            # 關閉連線
            connection_BCF_Change.close()

            return AK_BCF_Change_spectrum_Series_list
        else:
            connection_BCF_Differ = sqlite3.connect("BCF_Differ_spectrum.db")
            cursor_BCF_Differ = connection_BCF_Differ.cursor()
            # 取得Rdiffer資料
            column_name_BCF_Differ = self.TK_BCF_Combobox.currentText()
            # print("column_name_BCF_Differ",column_name_BCF_Differ)
            table_name_BCF_Differ = self.TK_BCF_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_BCF_Differ = f"SELECT * FROM '{table_name_BCF_Differ}';"
            cursor_BCF_Differ.execute(query_BCF_Differ)
            result_BCF_Differ = cursor_BCF_Differ.fetchall()

            # 找到指定標題的欄位索引
            header_BCF_Differ = [column[0] for column in cursor_BCF_Differ.description]

            # 移除所有標題中的編號
            header_BCF_Differ = [re.sub(r'_\d+$', '', col) for col in header_BCF_Differ]
            # print("header_BCF_Differ", header_BCF_Differ)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_BCF_Differ) if col == column_name_BCF_Differ]
            # print("indices_without_number", indices_without_number)

            # 使用 pandas Series 存儲結果
            series_list = [pd.Series([row[i] for row in result_BCF_Differ]) for i in indices_without_number]

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
            self.TK_BCF_Range.setText(f"TKrange: {min(TK_record):.2f}~{max(TK_record):.2f}")
            self.BCF_TK_list = []
            try:
                TK_BCF_Start = float(self.TK_BCF_Strat.text())
                TK_BCF_End = float(self.TK_BCF_End.text())
                TK_BCF_Interval = float(self.TK_BCF_interval.text())
                if float(self.TK_BCF_Strat.text()) < float(self.TK_BCF_End.text()):
                    i = 0
                    while float(self.TK_BCF_Strat.text()) + float(self.TK_BCF_interval.text()) * i < float(self.TK_BCF_End.text()):
                        self.BCF_TK_list.append(float(self.TK_BCF_Strat.text()) + float(self.TK_BCF_interval.text()) * i)
                        i += 1
                    self.BCF_TK_list.append(float(self.TK_BCF_End.text()))

                elif float(self.TK_BCF_End.text()) < float(self.TK_BCF_Strat.text()):
                    i = 0
                    while float(self.TK_BCF_End.text()) + float(self.TK_BCF_interval.text()) * i < float(self.TK_BCF_Strat.text()):
                        self.BCF_TK_list.append(float(self.TK_BCF_End.text()) + float(self.TK_BCF_interval.text()) * i)
                        i += 1
                    self.BCF_TK_list.append(float(self.TK_BCF_Strat.text()))
            except ValueError:
                # 如果轉換出錯，打印錯誤消息
                print("B輸入的數值格式不正確，請檢查後重新輸入。")
            print("self.BCF_TK_list",self.BCF_TK_list)
            AK_BCF_Change_spectrum_Series_list = []
            for BCF_TK in self.BCF_TK_list:
                B_differ_TK = BCF_TK
                Thickness_list = [B_differ_TK]
                for i in range(series_list_length):
                    Thickness_list.append(float(series_list[i].iloc[0]))
                    TK_record.append(float(series_list[i].iloc[0]))
                Thickness_list_2 = sorted(Thickness_list)
                print("Thickness_list_2", Thickness_list_2)
                index_B_differ_TK = Thickness_list_2.index(B_differ_TK)
                print(Thickness_list_2.index(B_differ_TK))

                # 輸入厚度的三種情況
                if B_differ_TK == min(Thickness_list):
                    print("min")
                    A_index_value = Thickness_list_2[index_B_differ_TK + 1]
                    # print("A_index_value", A_index_value)
                    B_index_value = Thickness_list_2[index_B_differ_TK + 2]
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
                    C_Series = A_Series + (B_differ_TK - A_index_value) * (B_Series - A_Series) / (
                            B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    AK_BCF_Change_spectrum_Series_list.append(C_Series)
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
                    AK_BCF_Change_spectrum_Series_list.append(C_Series)
                elif B_differ_TK in (Thickness_list[1:]):
                    print("equal")
                    # 取得對應的索引
                    Thickness_list.remove(Thickness_list[0])
                    equal_index = Thickness_list.index(B_differ_TK)
                    # print("Equal Index:", equal_index)
                    # print("Equal Value:", G_differ_TK)
                    C_Series = series_list[equal_index][1:]
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    AK_BCF_Change_spectrum_Series_list.append(C_Series)
                else:
                    print("mid")
                    # print("G_differ_TK",G_differ_TK)
                    # print("Thickness_list[1:]",Thickness_list[1:])
                    A_index_value = Thickness_list_2[index_B_differ_TK - 1]
                    # print("A_index_value",A_index_value)
                    B_index_value = Thickness_list_2[index_B_differ_TK + 1]
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
                    C_Series = A_Series + (B_differ_TK - A_index_value) * (B_Series - A_Series) / (
                                B_index_value - A_index_value)
                    # 在計算完 C_Series 後，加上以下代碼,將index重置
                    C_Series = C_Series.reset_index(drop=True)
                    AK_BCF_Change_spectrum_Series_list.append(C_Series)
            # 關閉連線
            connection_BCF_Differ.close()
            for row,TK in enumerate(self.BCF_TK_list):
                # item = QTableWidgetItem(str(TK))
                self.TK_BCF_Table_Form.setItem(row + 1, 3, QTableWidgetItem(f"{TK:.3f}"))
            print("AK_BCF_Change_spectrum_Series_list",AK_BCF_Change_spectrum_Series_list)
            return AK_BCF_Change_spectrum_Series_list

    def show_TK_RCF(self):
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
            self.Sort_TK_SP_list = sorted(TK_SP_list,reverse=True)
            print("self.Sort_TK_SP_list",self.Sort_TK_SP_list)
            self.TK_RCF_Range.setText(f"TKrange: {min(TK_SP_list):.2f}~{max(TK_SP_list):.2f}")

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

    def show_TK_GCF(self):
        if self.TK_GCF_mode_option.currentText() == "GCF_Change":
            print("in_GCF_Change自訂")
            connection_GCF_Change = sqlite3.connect("GCF_Change_spectrum.db")
            cursor_GCF_Change = connection_GCF_Change.cursor()
            # 取得Rdiffer資料
            column_name_GCF_Change = self.TK_GCF_Combobox.currentText()
            print("column_name_GCF_Change",column_name_GCF_Change)
            table_name_GCF_Change = self.TK_GCF_table.currentText()
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
            print("SP_series_list",SP_series_list)

            # TK_SP準備分厚度順序
            TK_SP_list = []
            for i in range(len(indices_without_number)):
                TK_SP_list.append(SP_series_list[i].iloc[0])
            print("TK_SP_list",TK_SP_list)
            self.Sort_TK_SP_list = sorted(TK_SP_list,reverse=True)
            print("self.Sort_TK_SP_list",self.Sort_TK_SP_list)
            self.TK_GCF_Range.setText(f"TKrange: {min(TK_SP_list):.2f}~{max(TK_SP_list):.2f}")
        else:
            connection_GCF_Differ = sqlite3.connect("GCF_Differ_spectrum.db")
            cursor_GCF_Differ = connection_GCF_Differ.cursor()
            # 取得Rdiffer資料
            column_name_GCF_Differ = self.TK_GCF_Combobox.currentText()
            # print("column_name_GCF_Differ",column_name_GCF_Differ)
            table_name_GCF_Differ = self.TK_GCF_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_GCF_Differ = f"SELECT * FROM '{table_name_GCF_Differ}';"
            cursor_GCF_Differ.execute(query_GCF_Differ)
            result_GCF_Differ = cursor_GCF_Differ.fetchall()

            # 找到指定標題的欄位索引
            header_GCF_Differ = [column[0] for column in cursor_GCF_Differ.description]

            # 移除所有標題中的編號
            header_GCF_Differ = [re.sub(r'_\d+$', '', col) for col in header_GCF_Differ]
            # print("header_GCF_Differ", header_GCF_Differ)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_GCF_Differ) if col == column_name_GCF_Differ]
            # print("indices_without_number", indices_without_number)

            # 使用 pandas Series 存儲結果
            series_list = [pd.Series([row[i] for row in result_GCF_Differ]) for i in indices_without_number]

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
            self.TK_GCF_Range.setText(f"TKrange: {min(TK_record):.2f}~{max(TK_record):.2f}")

    def show_TK_BCF(self):
        if self.TK_BCF_mode_option.currentText() == "BCF_Change":
            print("in_BCF_Change自訂")
            connection_BCF_Change = sqlite3.connect("BCF_Change_spectrum.db")
            cursor_BCF_Change = connection_BCF_Change.cursor()
            # 取得Rdiffer資料
            column_name_BCF_Change = self.TK_BCF_Combobox.currentText()
            print("column_name_BCF_Change",column_name_BCF_Change)
            table_name_BCF_Change = self.TK_BCF_table.currentText()
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
            print("SP_series_list",SP_series_list)

            # TK_SP準備分厚度順序
            TK_SP_list = []
            for i in range(len(indices_without_number)):
                TK_SP_list.append(SP_series_list[i].iloc[0])
            print("TK_SP_list",TK_SP_list)
            self.Sort_TK_SP_list = sorted(TK_SP_list,reverse=True)
            print("self.Sort_TK_SP_list",self.Sort_TK_SP_list)
            self.TK_BCF_Range.setText(f"TKrange: {min(TK_SP_list):.2f}~{max(TK_SP_list):.2f}")
        else:
            connection_BCF_Differ = sqlite3.connect("BCF_Differ_spectrum.db")
            cursor_BCF_Differ = connection_BCF_Differ.cursor()
            # 取得Rdiffer資料
            column_name_BCF_Differ = self.TK_BCF_Combobox.currentText()
            # print("column_name_BCF_Differ",column_name_BCF_Differ)
            table_name_BCF_Differ = self.TK_BCF_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_BCF_Differ = f"SELECT * FROM '{table_name_BCF_Differ}';"
            cursor_BCF_Differ.execute(query_BCF_Differ)
            result_BCF_Differ = cursor_BCF_Differ.fetchall()

            # 找到指定標題的欄位索引
            header_BCF_Differ = [column[0] for column in cursor_BCF_Differ.description]

            # 移除所有標題中的編號
            header_BCF_Differ = [re.sub(r'_\d+$', '', col) for col in header_BCF_Differ]
            # print("header_BCF_Differ", header_BCF_Differ)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(header_BCF_Differ) if col == column_name_BCF_Differ]
            # print("indices_without_number", indices_without_number)

            # 使用 pandas Series 存儲結果
            series_list = [pd.Series([row[i] for row in result_BCF_Differ]) for i in indices_without_number]

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
            self.TK_BCF_Range.setText(f"TKrange: {min(TK_record):.2f}~{max(TK_record):.2f}")

    # DRAW PICTURE
    def draw_R(self):
        # 清空畫布
        self.CIE_R_chart.clear()
        RGB_values = []
        for x, y in zip(self.TK_R_x_list, self.TK_R_y_list):
            XYZ = xy_to_XYZ(np.array([x, y]))
            RGB = XYZ_to_sRGB(XYZ)
            RGB_values.append(RGB)

            # 转换为 numpy 数组
        RGB_array = np.array(RGB_values)


        # 显示对应的颜色在 CIE 1931 色度图上
        # 注意：这里可能需要调整以适应您的具体需求
        plotting.plot_chromaticity_diagram_CIE1931(axes =self.CIE_R_chart, show=False)

        # 使用 plt.scatter 添加自定义的标记
        # markers = ['o', 's', '^', '*']  # 可以根据需要更改标记形状
        labels = []
        for i in range (len(self.TK_R_x_list)):
            labels.append(i)
        edgecolors = 'black'
        for x, y, label, RGB in zip(self.TK_R_x_list, self.TK_R_y_list, labels, RGB_array):
            # 确保颜色值在 0 到 1 的范围内
            RGB_clipped = np.clip(RGB, 0, 1)
            self.CIE_R_chart.scatter(x, y, label=f'{label}:({x:.3f}, {y:.3f})', c=[RGB_clipped],
                        edgecolors=edgecolors)

        plt.tight_layout()

        # 显示图例
        self.CIE_R_chart.legend(loc='lower center', bbox_to_anchor=(1.2, 0.2), fontsize='small')

        self.CIE_R_canvas.draw()

        # 连接滚轮事件
        self.CIE_R_canvas.mpl_connect('scroll_event', self.zoom_R_on_scroll)
        self.CIE_R_canvas.mpl_connect('button_press_event', self.on_R_press)
        self.CIE_R_canvas.mpl_connect('motion_notify_event', self.on_R_drag)
        self.CIE_R_canvas.mpl_connect('button_release_event', self.on_R_release)
        self.CIE_R_canvas.mpl_connect('motion_notify_event', self.hover_R)

        self.dragging = False

        # 创建注释，用于显示信息
        self.annot_R = self.CIE_R_chart.annotate("", xy=(0, 0), xytext=(-20, 20),
                                               textcoords="offset points",
                                               bbox=dict(boxstyle="round", fc="w"),
                                               arrowprops=dict(arrowstyle="->"))
        self.annot_R.set_visible(False)

    def draw_G(self):
        # 清空畫布
        self.CIE_G_chart.clear()
        RGB_values = []
        for x, y in zip(self.TK_G_x_list, self.TK_G_y_list):
            XYZ = xy_to_XYZ(np.array([x, y]))
            RGB = XYZ_to_sRGB(XYZ)
            RGB_values.append(RGB)

            # 转换为 numpy 数组
        RGB_array = np.array(RGB_values)


        # 显示对应的颜色在 CIE 1931 色度图上
        # 注意：这里可能需要调整以适应您的具体需求
        plotting.plot_chromaticity_diagram_CIE1931(axes =self.CIE_G_chart, show=False)

        # 使用 plt.scatter 添加自定义的标记
        # markers = ['o', 's', '^', '*']  # 可以根据需要更改标记形状
        labels = []
        for i in range (len(self.TK_G_x_list)):
            labels.append(i)
        edgecolors = 'black'
        for x, y, label, RGB in zip(self.TK_G_x_list, self.TK_G_y_list, labels, RGB_array):
            # 确保颜色值在 0 到 1 的范围内
            RGB_clipped = np.clip(RGB, 0, 1)
            self.CIE_G_chart.scatter(x, y, label=f'{label}:({x:.3f}, {y:.3f})', c=[RGB_clipped],
                        edgecolors=edgecolors)

        plt.tight_layout()

        # 显示图例
        self.CIE_G_chart.legend(loc='lower center', bbox_to_anchor=(1.2, 0.2), fontsize='small')

        self.CIE_G_canvas.draw()

        # 连接滚轮事件
        self.CIE_G_canvas.mpl_connect('scroll_event', self.zoom_G_on_scroll)
        self.CIE_G_canvas.mpl_connect('button_press_event', self.on_G_press)
        self.CIE_G_canvas.mpl_connect('motion_notify_event', self.on_G_drag)
        self.CIE_G_canvas.mpl_connect('button_release_event', self.on_G_release)
        self.CIE_G_canvas.mpl_connect('motion_notify_event', self.hover_G)

        self.dragging = False

        # 创建注释，用于显示信息
        self.annot_G = self.CIE_G_chart.annotate("", xy=(0, 0), xytext=(-20, 20),
                                               textcoords="offset points",
                                               bbox=dict(boxstyle="round", fc="w"),
                                               arrowprops=dict(arrowstyle="->"))
        self.annot_G.set_visible(False)

    def draw_B(self):
        # 清空畫布
        self.CIE_B_chart.clear()
        RGB_values = []
        for x, y in zip(self.TK_B_x_list, self.TK_B_y_list):
            XYZ = xy_to_XYZ(np.array([x, y]))
            RGB = XYZ_to_sRGB(XYZ)
            RGB_values.append(RGB)

            # 转换为 numpy 数组
        RGB_array = np.array(RGB_values)


        # 显示对应的颜色在 CIE 1931 色度图上
        # 注意：这里可能需要调整以适应您的具体需求
        plotting.plot_chromaticity_diagram_CIE1931(axes =self.CIE_B_chart, show=False)

        # 使用 plt.scatter 添加自定义的标记
        # markers = ['o', 's', '^', '*']  # 可以根据需要更改标记形状
        labels = []
        for i in range (len(self.TK_B_x_list)):
            labels.append(i)
        edgecolors = 'black'
        for x, y, label, RGB in zip(self.TK_B_x_list, self.TK_B_y_list, labels, RGB_array):
            # 确保颜色值在 0 到 1 的范围内
            RGB_clipped = np.clip(RGB, 0, 1)
            self.CIE_B_chart.scatter(x, y, label=f'{label}:({x:.3f}, {y:.3f})', c=[RGB_clipped],
                        edgecolors=edgecolors)

        plt.tight_layout()

        # 显示图例
        self.CIE_B_chart.legend(loc='lower center', bbox_to_anchor=(1.2, 0.2), fontsize='small')

        self.CIE_B_canvas.draw()

        # 连接滚轮事件
        self.CIE_B_canvas.mpl_connect('scroll_event', self.zoom_B_on_scroll)
        self.CIE_B_canvas.mpl_connect('button_press_event', self.on_B_press)
        self.CIE_B_canvas.mpl_connect('motion_notify_event', self.on_B_drag)
        self.CIE_B_canvas.mpl_connect('button_release_event', self.on_B_release)
        self.CIE_B_canvas.mpl_connect('motion_notify_event', self.hover_B)

        self.dragging = False

        # 创建注释，用于显示信息
        self.annot_B = self.CIE_B_chart.annotate("", xy=(0, 0), xytext=(-20, 20),
                                               textcoords="offset points",
                                               bbox=dict(boxstyle="round", fc="w"),
                                               arrowprops=dict(arrowstyle="->"))
        self.annot_B.set_visible(False)


    # FUNCTION
    def zoom_R_on_scroll(self, event: MouseEvent):
        # 缩放函数
        base_scale = 1.1  # 缩放基数
        cur_xlim = self.CIE_R_chart.get_xlim()
        cur_ylim = self.CIE_R_chart.get_ylim()
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
        self.CIE_R_chart.set_xlim([xdata - cur_xrange * scale_factor,
                                   xdata + cur_xrange * scale_factor])
        self.CIE_R_chart.set_ylim([ydata - cur_yrange * scale_factor,
                                   ydata + cur_yrange * scale_factor])

        self.CIE_R_canvas.draw()

    def on_R_press(self, event: MouseEvent):
        if event.button == 1:  # 检查是否为鼠标左键
            self.dragging = True
            self.drag_start_x = event.xdata
            self.drag_start_y = event.ydata

    def on_R_drag(self, event: MouseEvent):
        if self.dragging:
            dx = event.xdata - self.drag_start_x
            dy = event.ydata - self.drag_start_y
            self.drag_start_x = event.xdata
            self.drag_start_y = event.ydata
            self.pan_R(dx, dy)

    def on_R_release(self, event: MouseEvent):
        self.dragging = False

    def pan_R(self, dx, dy):
        cur_xlim = self.CIE_R_chart.get_xlim()
        cur_ylim = self.CIE_R_chart.get_ylim()

        self.CIE_R_chart.set_xlim(cur_xlim[0] - dx, cur_xlim[1] - dx)
        self.CIE_R_chart.set_ylim(cur_ylim[0] - dy, cur_ylim[1] - dy)

        self.CIE_R_canvas.draw()

    def hover_R(self, event):
        # 鼠标悬停事件处理函数
        vis = self.annot_R.get_visible()
        if event.inaxes == self.CIE_R_chart:
            for scatter in self.CIE_R_chart.collections:
                cont, ind = scatter.contains(event)
                if cont:
                    pos = scatter.get_offsets()[ind["ind"][0]]
                    self.annot_R.xy = pos
                    text = f"{pos[0]:.3f}, {pos[1]:.3f}"
                    self.annot_R.set_text(text)
                    self.annot_R.set_visible(True)
                    self.CIE_R_canvas.draw_idle()
                    return
        if vis:
            self.annot_R.set_visible(False)
            self.CIE_R_canvas.draw_idle()

    def zoom_G_on_scroll(self, event: MouseEvent):
        # 缩放函数
        base_scale = 1.1  # 缩放基数
        cur_xlim = self.CIE_G_chart.get_xlim()
        cur_ylim = self.CIE_G_chart.get_ylim()
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
        self.CIE_G_chart.set_xlim([xdata - cur_xrange * scale_factor,
                                   xdata + cur_xrange * scale_factor])
        self.CIE_G_chart.set_ylim([ydata - cur_yrange * scale_factor,
                                   ydata + cur_yrange * scale_factor])

        self.CIE_G_canvas.draw()

    def on_G_press(self, event: MouseEvent):
        if event.button == 1:  # 检查是否为鼠标左键
            self.dragging = True
            self.drag_start_x = event.xdata
            self.drag_start_y = event.ydata

    def on_G_drag(self, event: MouseEvent):
        if self.dragging:
            dx = event.xdata - self.drag_start_x
            dy = event.ydata - self.drag_start_y
            self.drag_start_x = event.xdata
            self.drag_start_y = event.ydata
            self.pan_G(dx, dy)

    def on_G_release(self, event: MouseEvent):
        self.dragging = False

    def pan_G(self, dx, dy):
        cur_xlim = self.CIE_G_chart.get_xlim()
        cur_ylim = self.CIE_G_chart.get_ylim()

        self.CIE_G_chart.set_xlim(cur_xlim[0] - dx, cur_xlim[1] - dx)
        self.CIE_G_chart.set_ylim(cur_ylim[0] - dy, cur_ylim[1] - dy)

        self.CIE_G_canvas.draw()

    def hover_G(self, event):
        # 鼠标悬停事件处理函数
        vis = self.annot_G.get_visible()
        if event.inaxes == self.CIE_G_chart:
            for scatter in self.CIE_G_chart.collections:
                cont, ind = scatter.contains(event)
                if cont:
                    pos = scatter.get_offsets()[ind["ind"][0]]
                    self.annot_G.xy = pos
                    text = f"{pos[0]:.3f}, {pos[1]:.3f}"
                    self.annot_G.set_text(text)
                    self.annot_G.set_visible(True)
                    self.CIE_G_canvas.draw_idle()
                    return
        if vis:
            self.annot_G.set_visible(False)
            self.CIE_G_canvas.draw_idle()

    def zoom_B_on_scroll(self, event: MouseEvent):
        # 缩放函数
        base_scale = 1.1  # 缩放基数
        cur_xlim = self.CIE_G_chart.get_xlim()
        cur_ylim = self.CIE_G_chart.get_ylim()
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
        self.CIE_B_chart.set_xlim([xdata - cur_xrange * scale_factor,
                                   xdata + cur_xrange * scale_factor])
        self.CIE_B_chart.set_ylim([ydata - cur_yrange * scale_factor,
                                   ydata + cur_yrange * scale_factor])

        self.CIE_B_canvas.draw()

    def on_B_press(self, event: MouseEvent):
        if event.button == 1:  # 检查是否为鼠标左键
            self.dragging = True
            self.drag_start_x = event.xdata
            self.drag_start_y = event.ydata

    def on_B_drag(self, event: MouseEvent):
        if self.dragging:
            dx = event.xdata - self.drag_start_x
            dy = event.ydata - self.drag_start_y
            self.drag_start_x = event.xdata
            self.drag_start_y = event.ydata
            self.pan_B(dx, dy)

    def on_B_release(self, event: MouseEvent):
        self.dragging = False

    def pan_B(self, dx, dy):
        cur_xlim = self.CIE_B_chart.get_xlim()
        cur_ylim = self.CIE_B_chart.get_ylim()

        self.CIE_B_chart.set_xlim(cur_xlim[0] - dx, cur_xlim[1] - dx)
        self.CIE_B_chart.set_ylim(cur_ylim[0] - dy, cur_ylim[1] - dy)

        self.CIE_B_canvas.draw()

    def hover_B(self, event):
        # 鼠标悬停事件处理函数
        vis = self.annot_B.get_visible()
        if event.inaxes == self.CIE_B_chart:
            for scatter in self.CIE_B_chart.collections:
                cont, ind = scatter.contains(event)
                if cont:
                    pos = scatter.get_offsets()[ind["ind"][0]]
                    self.annot_B.xy = pos
                    text = f"{pos[0]:.3f}, {pos[1]:.3f}"
                    self.annot_B.set_text(text)
                    self.annot_B.set_visible(True)
                    self.CIE_B_canvas.draw_idle()
                    return
        if vis:
            self.annot_B.set_visible(False)
            self.CIE_B_canvas.draw_idle()



    def calculate_TK_Color(self):
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

        # BLU +Cell part
        cell_blu_total_spectrum = self.calculate_TK_BLU() * self.calculate_TK_layer1() * self.calculate_TK_layer2() \
                                  * self.calculate_TK_layer3() * self.calculate_TK_layer4() * self.calculate_TK_layer5() \
                                  * self.calculate_TK_layer6()
        # RCF--------------------------------------------------------------------
        self.TK_R_x_list = []
        self.TK_R_y_list = []
        self.TK_R_wave = []
        self.TK_R_Purity = []
        for R_spectrum in self.calculate_TK_RCF():
            R = cell_blu_total_spectrum * R_spectrum
            R_X = R * CIE_spectrum_Series_X
            # print("R_X", R_X)
            R_Y = R * CIE_spectrum_Series_Y
            # print("R_Y", R_Y)
            R_Z = R * CIE_spectrum_Series_Z
            # print("R_Z", R_Z)
            R_X_sum = R_X.sum()
            # print("R_X_sum", R_X_sum)
            R_Y_sum = R_Y.sum()
            # print("R_Y_sum", R_Y_sum)
            R_Z_sum = R_Z.sum()
            self.TK_R_x = R_X_sum / (R_X_sum + R_Y_sum + R_Z_sum)
            self.TK_R_y = R_Y_sum / (R_X_sum + R_Y_sum + R_Z_sum)
            Temp_XY = [self.TK_R_x,self.TK_R_y]
            # 主波長計算
            dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(Temp_XY, self.xy_n)
            CIEcoordinate_value_1 = CIEcoordinate[0]
            CIEcoordinate_value_2 = CIEcoordinate[1]
            # 計算距離
            distance_from_white = math.sqrt((self.TK_R_x - self.xy_n[0]) ** 2 + (self.TK_R_y - self.xy_n[1]) ** 2)
            distance_from_black = math.sqrt(
                (CIEcoordinate_value_1 - self.xy_n[0]) ** 2 + (CIEcoordinate_value_2 - self.xy_n[1]) ** 2)

            # 計算色純度
            purity = distance_from_white / distance_from_black
            self.TK_R_x_list.append(self.TK_R_x)
            self.TK_R_y_list.append(self.TK_R_y)
            self.TK_R_wave.append(dominant_wavelength)
            self.TK_R_Purity.append(purity)
        # 更新Table_Form的列數
        self.TK_RCF_Table_Form.setRowCount(len(self.RCF_TK_list)+1)

        for row in range(len(self.RCF_TK_list)):
            self.TK_RCF_Table_Form.setItem(row+1,0,QTableWidgetItem(f"{row}"))

        for row, Rx in enumerate(self.TK_R_x_list):
            # item = QTableWidgetItem(str(Rx))
            self.TK_RCF_Table_Form.setItem(row + 1, 1, QTableWidgetItem(f"{Rx:.3f}"))

        for row, Ry in enumerate(self.TK_R_y_list):
            # item = QTableWidgetItem(str(Ry))
            self.TK_RCF_Table_Form.setItem(row + 1, 2, QTableWidgetItem(f"{Ry:.3f}"))

        for row, R_wave in enumerate(self.TK_R_wave):
            # item = QTableWidgetItem(str(Ry))
            self.TK_RCF_Table_Form.setItem(row + 1, 4, QTableWidgetItem(f"{R_wave:.3f}"))

        for row, R_purity in enumerate(self.TK_R_Purity):
            # item = QTableWidgetItem(str(Ry))
            self.TK_RCF_Table_Form.setItem(row + 1, 5, QTableWidgetItem(f"{R_purity:.3f}"))

        self.draw_R()

        # GCF--------------------------------------------------------------------
        self.TK_G_x_list = []
        self.TK_G_y_list = []
        self.TK_G_wave = []
        self.TK_G_Purity = []
        for G_spectrum in self.calculate_TK_GCF():
            G = cell_blu_total_spectrum * G_spectrum
            G_X = G * CIE_spectrum_Series_X
            # print("R_X", R_X)
            G_Y = G * CIE_spectrum_Series_Y
            # print("R_Y", R_Y)
            G_Z = G * CIE_spectrum_Series_Z
            # print("R_Z", R_Z)
            G_X_sum = G_X.sum()
            # print("R_X_sum", R_X_sum)
            G_Y_sum = G_Y.sum()
            # print("R_Y_sum", R_Y_sum)
            G_Z_sum = G_Z.sum()
            self.TK_G_x = G_X_sum / (G_X_sum + G_Y_sum + G_Z_sum)
            self.TK_G_y = G_Y_sum / (G_X_sum + G_Y_sum + G_Z_sum)
            Temp_XY = [self.TK_G_x, self.TK_G_y]
            # 主波長計算
            dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(Temp_XY, self.xy_n)
            CIEcoordinate_value_1 = CIEcoordinate[0]
            CIEcoordinate_value_2 = CIEcoordinate[1]
            # 計算距離
            distance_from_white = math.sqrt((self.TK_G_x - self.xy_n[0]) ** 2 + (self.TK_G_y - self.xy_n[1]) ** 2)
            distance_from_black = math.sqrt(
                (CIEcoordinate_value_1 - self.xy_n[0]) ** 2 + (CIEcoordinate_value_2 - self.xy_n[1]) ** 2)

            # 計算色純度
            purity = distance_from_white / distance_from_black
            self.TK_G_x_list.append(self.TK_G_x)
            self.TK_G_y_list.append(self.TK_G_y)
            self.TK_G_wave.append(dominant_wavelength)
            self.TK_G_Purity.append(purity)
        self.TK_GCF_Table_Form.setRowCount(len(self.GCF_TK_list)+1)

        for row in range(len(self.GCF_TK_list)):
            self.TK_GCF_Table_Form.setItem(row + 1, 0, QTableWidgetItem(f"{row}"))

        for row, Gx in enumerate(self.TK_G_x_list):
            # item = QTableWidgetItem(str(Gx))
            self.TK_GCF_Table_Form.setItem(row + 1, 1, QTableWidgetItem(f"{Gx:.3f}"))

        for row, Gy in enumerate(self.TK_G_y_list):
            # item = QTableWidgetItem(str(Ry))
            self.TK_GCF_Table_Form.setItem(row + 1, 2, QTableWidgetItem(f"{Gy:.3f}"))

        for row, G_wave in enumerate(self.TK_G_wave):
            # item = QTableWidgetItem(str(Ry))
            self.TK_GCF_Table_Form.setItem(row + 1, 4, QTableWidgetItem(f"{G_wave:.3f}"))

        for row, G_purity in enumerate(self.TK_G_Purity):
            # item = QTableWidgetItem(str(Ry))
            self.TK_GCF_Table_Form.setItem(row + 1, 5, QTableWidgetItem(f"{G_purity:.3f}"))

        self.draw_G()

        # BCF--------------------------------------------------------------------
        self.TK_B_x_list = []
        self.TK_B_y_list = []
        self.TK_B_wave = []
        self.TK_B_Purity = []
        for B_spectrum in self.calculate_TK_BCF():
            B = cell_blu_total_spectrum * B_spectrum
            B_X = B * CIE_spectrum_Series_X
            # print("B_X", B_X)
            B_Y = B * CIE_spectrum_Series_Y
            # print("R_Y", R_Y)
            B_Z = B * CIE_spectrum_Series_Z
            # print("R_Z", R_Z)
            B_X_sum = B_X.sum()
            # print("B_X_sum", B_X_sum)
            B_Y_sum = B_Y.sum()
            # print("B_Y_sum", B_Y_sum)
            B_Z_sum = B_Z.sum()
            self.TK_B_x = B_X_sum / (B_X_sum + B_Y_sum + B_Z_sum)
            self.TK_B_y = B_Y_sum / (B_X_sum + B_Y_sum + B_Z_sum)
            Temp_XY = [self.TK_B_x, self.TK_B_y]
            # 主波長計算
            dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(Temp_XY, self.xy_n)
            CIEcoordinate_value_1 = CIEcoordinate[0]
            CIEcoordinate_value_2 = CIEcoordinate[1]
            # 計算距離
            distance_from_white = math.sqrt((self.TK_B_x - self.xy_n[0]) ** 2 + (self.TK_B_y - self.xy_n[1]) ** 2)
            distance_from_black = math.sqrt(
                (CIEcoordinate_value_1 - self.xy_n[0]) ** 2 + (CIEcoordinate_value_2 - self.xy_n[1]) ** 2)

            # 計算色純度
            purity = distance_from_white / distance_from_black
            self.TK_B_x_list.append(self.TK_B_x)
            self.TK_B_y_list.append(self.TK_B_y)
            self.TK_B_wave.append(dominant_wavelength)
            self.TK_B_Purity.append(purity)
        self.TK_BCF_Table_Form.setRowCount(len(self.BCF_TK_list)+1)

        for row in range(len(self.BCF_TK_list)):
            self.TK_BCF_Table_Form.setItem(row + 1, 0, QTableWidgetItem(f"{row}"))

        for row, Bx in enumerate(self.TK_B_x_list):
            # item = QTableWidgetItem(str(Gx))
            self.TK_BCF_Table_Form.setItem(row + 1, 1, QTableWidgetItem(f"{Bx:.3f}"))

        for row, By in enumerate(self.TK_B_y_list):
            # item = QTableWidgetItem(str(By))
            self.TK_BCF_Table_Form.setItem(row + 1, 2, QTableWidgetItem(f"{By:.3f}"))

        for row, B_wave in enumerate(self.TK_B_wave):
            # item = QTableWidgetItem(str(Ry))
            self.TK_BCF_Table_Form.setItem(row + 1, 4, QTableWidgetItem(f"{B_wave:.3f}"))

        for row, B_purity in enumerate(self.TK_B_Purity):
            # item = QTableWidgetItem(str(Ry))
            self.TK_BCF_Table_Form.setItem(row + 1, 5, QTableWidgetItem(f"{B_purity:.3f}"))

        self.draw_B()

    def calculae_TK_Color_Twice(self):
        self.calculate_TK_Color()
        self.calculate_TK_Color()


    # COPY FUNCTION
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