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
from collections import Counter
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
import itertools

class CS_evaluation(QWidget):
    def __init__(self):
        super().__init__()
        # dialog layout
        self.Color_Select_layout = QGridLayout()

        # 觀察者光源
        self.observer_D65 = [0.3127, 0.329]

        # 創建一個通用的 Ctrl+C 快捷鍵
        copy_shortcut = QShortcut(QKeySequence.Copy, self)
        copy_shortcut.activated.connect(self.copy_table_content)

        # TK_BLU
        self.CS_light_source_label = QLabel("Light_Source")
        self.CS_light_source_datatable = QComboBox()
        self.CS_light_source_datatable.setStyleSheet(QCOMBOXSETTING)
        self.CS_light_source_datatable.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.CS_light_source = QComboBox()
        self.CS_light_source.setStyleSheet(QCOMBOXSETTING)
        # 更新light_source_table
        self.update_CS_light_source_datatable()
        # 觸發table更新,lightsource表單
        self.update_CS_light_sourceComboBox()  # 初始化
        self.CS_light_source_datatable.currentIndexChanged.connect(self.update_CS_light_sourceComboBox)

        # layer1
        # self.TK_layer1_label = QLabel("Layer_1")
        self.CS_layer1_mode = QComboBox()
        self.CS_layer1_mode.addItem("Layer1_未選")
        self.CS_layer1_mode.addItem("layer1_自訂")
        self.CS_layer1_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.CS_layer1_box = QComboBox()
        self.CS_layer1_box.setStyleSheet(QCOMBOXSETTING)
        self.CS_layer1_table = QComboBox()
        self.CS_layer1_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_CS_layer1_datatable()
        self.update_CS_layer1ComboBox()  # 初始化
        self.CS_layer1_table.currentIndexChanged.connect(self.update_CS_layer1ComboBox)

        # layer2
        # self.TK_layer2_label = QLabel("Layer_2")
        self.CS_layer2_mode = QComboBox()
        self.CS_layer2_mode.addItem("Layer2_未選")
        self.CS_layer2_mode.addItem("layer2_自訂")
        self.CS_layer2_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.CS_layer2_box = QComboBox()
        self.CS_layer2_box.setStyleSheet(QCOMBOXSETTING)
        self.CS_layer2_table = QComboBox()
        self.CS_layer2_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_CS_layer2_datatable()
        self.update_CS_layer2ComboBox()  # 初始化
        self.CS_layer2_table.currentIndexChanged.connect(self.update_CS_layer2ComboBox)

        # layer3
        # self.TK_layer3_label = QLabel("Layer_3")
        self.CS_layer3_mode = QComboBox()
        self.CS_layer3_mode.addItem("Layer3_未選")
        self.CS_layer3_mode.addItem("layer3_自訂")
        self.CS_layer3_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.CS_layer3_box = QComboBox()
        self.CS_layer3_box.setStyleSheet(QCOMBOXSETTING)
        self.CS_layer3_table = QComboBox()
        self.CS_layer3_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_CS_layer3_datatable()
        self.update_CS_layer3ComboBox()  # 初始化
        self.CS_layer3_table.currentIndexChanged.connect(self.update_CS_layer3ComboBox)

        # layer4
        # self.TK_layer4_label = QLabel("Layer_4")
        self.CS_layer4_mode = QComboBox()
        self.CS_layer4_mode.addItem("Layer4_未選")
        self.CS_layer4_mode.addItem("layer4_自訂")
        self.CS_layer4_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.CS_layer4_box = QComboBox()
        self.CS_layer4_box.setStyleSheet(QCOMBOXSETTING)
        self.CS_layer4_table = QComboBox()
        self.CS_layer4_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_CS_layer4_datatable()
        self.update_CS_layer4ComboBox()  # 初始化
        self.CS_layer4_table.currentIndexChanged.connect(self.update_CS_layer4ComboBox)


        # layer5
        # self.TK_layer5_label = QLabel("Layer_5")
        self.CS_layer5_mode = QComboBox()
        self.CS_layer5_mode.addItem("Layer5_未選")
        self.CS_layer5_mode.addItem("layer5_自訂")
        self.CS_layer5_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.CS_layer5_box = QComboBox()
        self.CS_layer5_box.setStyleSheet(QCOMBOXSETTING)
        self.CS_layer5_table = QComboBox()
        self.CS_layer5_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_CS_layer5_datatable()
        self.update_CS_layer5ComboBox()  # 初始化
        self.CS_layer5_table.currentIndexChanged.connect(self.update_CS_layer5ComboBox)

        # layer6
        # self.TK_layer6_label = QLabel("Layer_6")
        self.CS_layer6_mode = QComboBox()
        self.CS_layer6_mode.addItem("Layer6_未選")
        self.CS_layer6_mode.addItem("layer6_自訂")
        self.CS_layer6_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.CS_layer6_box = QComboBox()
        self.CS_layer6_box.setStyleSheet(QCOMBOXSETTING)
        self.CS_layer6_table = QComboBox()
        self.CS_layer6_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        self.update_CS_layer6_datatable()
        self.update_CS_layer6ComboBox()  # 初始化
        self.CS_layer6_table.currentIndexChanged.connect(self.update_CS_layer6ComboBox)

        # RCF
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
        self.R_fix_table = QComboBox()
        # 防呆反灰設定
        # self.R_fix_box.setEnabled(False)
        # self.R_fix_table.setEnabled(False)
        self.R_fix_box.setStyleSheet(QCOMBOXSETTING)
        self.R_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 觸發table連動改變
        self.update_RCF_Fix_datatable()
        self.updateRCF_Fix_ComboBox()
        self.R_fix_table.currentIndexChanged.connect(self.updateRCF_Fix_ComboBox)
        #self.R_fix_mode.currentIndexChanged.connect(self.update_RCF_Fix_modeclose)

        self.R_TK_edit_label = QLabel("R-Fix-TK")
        # self.R_TK_edit = QLineEdit()
        # self.R_TK_edit.setFixedSize(100, 25)
        self.R_TK_label = QLabel()


        # G-Fix
        self.G_fix_mode = QComboBox()
        G_fix_mode_items = ["未選", "自訂", "模擬"]
        for item in G_fix_mode_items:
            self.G_fix_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.G_fix_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.G_fix_label = QLabel("G-CF-Fix")
        self.G_fix_box = QComboBox()
        self.G_fix_table = QComboBox()
        # 防呆反灰設定
        # self.G_fix_box.setEnabled(False)
        # self.G_fix_table.setEnabled(False)
        self.G_fix_box.setStyleSheet(QCOMBOXSETTING)
        self.G_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 觸發table連動改變
        self.update_GCF_Fix_datatable()
        self.updateGCF_Fix_ComboBox()
        self.G_fix_table.currentIndexChanged.connect(self.updateGCF_Fix_ComboBox)
        #self.G_fix_mode.currentIndexChanged.connect(self.update_GCF_Fix_modeclose)

        self.G_TK_edit_label = QLabel("G-Fix-TK")
        # self.G_TK_edit = QLineEdit()
        # self.G_TK_edit.setFixedSize(100, 25)
        self.G_TK_label = QLabel()


        # B-Fix
        self.B_fix_mode = QComboBox()
        B_fix_mode_items = ["未選", "自訂", "模擬"]
        for item in B_fix_mode_items:
            self.B_fix_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.B_fix_mode.setStyleSheet(QCOMBOXMODESETTING)
        self.B_fix_label = QLabel("B-CF-Fix")
        self.B_fix_box = QComboBox()
        self.B_fix_table = QComboBox()
        # 防呆反灰設定
        # self.B_fix_box.setEnabled(False)
        # self.B_fix_table.setEnabled(False)
        self.B_fix_box.setStyleSheet(QCOMBOXSETTING)
        self.B_fix_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 觸發table連動改變
        self.update_BCF_Fix_datatable()
        self.updateBCF_Fix_ComboBox()
        self.B_fix_table.currentIndexChanged.connect(self.updateBCF_Fix_ComboBox)
        #self.B_fix_mode.currentIndexChanged.connect(self.update_BCF_Fix_modeclose)

        self.B_TK_edit_label = QLabel("B-Fix-TK")
        # self.B_TK_edit = QLineEdit()
        # self.B_TK_edit.setFixedSize(100, 25)
        self.B_TK_label = QLabel()


        # RGB-α,K 區域---------------------------------------------------
        self.RGB_aK_label = QLabel("RGB-α,K")
        self.RGB_aK_label.setStyleSheet("color: #5151A2; font-weight: bold; border: 2px solid black;")
        self.R_aK_mode = QComboBox()
        R_aK_mode_items = ["未選", "自訂", "模擬"]
        for item in R_aK_mode_items:
            self.R_aK_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.R_aK_mode.setStyleSheet(QCOMBOXMODESETTING)
        # self.R_aK_label = QLabel("R-CF-α,K")
        self.R_aK_box = QComboBox()
        self.R_aK_table = QComboBox()
        # 防呆反灰設定
        # self.R_aK_box.setEnabled(False)
        # self.R_aK_table.setEnabled(False)
        self.R_aK_box.setStyleSheet(QCOMBOXSETTING)
        self.R_aK_table.setStyleSheet(QCOMBOBOXTABLESELECT)

        # 觸發table連動改變
        self.update_RCF_Change_datatable()
        self.updateRCF_Change_ComboBox()
        self.R_aK_table.currentIndexChanged.connect(self.updateRCF_Change_ComboBox)
        #self.R_aK_mode.currentIndexChanged.connect(self.update_RCF_Change_modeclose)

        self.R_aK_TK_edit_label = QLabel("R-α,K-TK")
        self.R_aK_TK_edit = QLineEdit()
        self.R_aK_TK_edit.setFixedSize(100, 25)

        self.R_aK_TK_Start = QLineEdit()
        self.R_aK_TK_Start.setPlaceholderText("TK_Start")
        self.R_aK_TK_End = QLineEdit()
        self.R_aK_TK_End.setPlaceholderText("TK_End")
        self.R_aK_TK_interval = QLineEdit()
        self.R_aK_TK_interval.setPlaceholderText("Interval")

        # G-α,K
        self.G_aK_mode = QComboBox()
        G_aK_mode_items = ["未選", "自訂", "模擬"]
        for item in G_aK_mode_items:
            self.G_aK_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.G_aK_mode.setStyleSheet(QCOMBOXMODESETTING)
        # self.G_aK_label = QLabel("G-F-α,K")
        self.G_aK_box = QComboBox()
        self.G_aK_table = QComboBox()
        # 防呆反灰設定
        # self.G_aK_box.setEnabled(False)
        # self.G_aK_table.setEnabled(False)
        self.G_aK_box.setStyleSheet(QCOMBOXSETTING)
        self.G_aK_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 觸發table連動改變
        self.update_GCF_Change_datatable()
        self.updateGCF_Change_ComboBox()
        self.G_aK_table.currentIndexChanged.connect(self.updateGCF_Change_ComboBox)
        #self.G_aK_mode.currentIndexChanged.connect(self.update_GCF_Change_modeclose)

        self.G_aK_TK_edit_label = QLabel("G-α,K-TK")
        self.G_aK_TK_edit = QLineEdit()
        self.G_aK_TK_edit.setFixedSize(100, 25)

        self.G_aK_TK_Start = QLineEdit()
        self.G_aK_TK_Start.setPlaceholderText("TK_Start")
        self.G_aK_TK_End = QLineEdit()
        self.G_aK_TK_End.setPlaceholderText("TK_End")
        self.G_aK_TK_interval = QLineEdit()
        self.G_aK_TK_interval.setPlaceholderText("Interval")

        # B-α,K
        self.B_aK_mode = QComboBox()
        B_aK_mode_items = ["未選", "自訂", "模擬"]
        for item in B_aK_mode_items:
            self.B_aK_mode.addItem(str(item))
        # 設定當前選中項目的文字顏色
        self.B_aK_mode.setStyleSheet(QCOMBOXMODESETTING)
        # self.B_aK_label = QLabel("B-CF-α,K")
        self.B_aK_box = QComboBox()
        self.B_aK_table = QComboBox()
        # 防呆反灰設定
        # self.B_aK_box.setEnabled(False)
        # self.B_aK_table.setEnabled(False)
        self.B_aK_box.setStyleSheet(QCOMBOXSETTING)
        self.B_aK_table.setStyleSheet(QCOMBOBOXTABLESELECT)
        # 觸發table連動改變
        self.update_BCF_Change_datatable()
        self.updateBCF_Change_ComboBox()
        self.B_aK_table.currentIndexChanged.connect(self.updateBCF_Change_ComboBox)
        #self.B_aK_mode.currentIndexChanged.connect(self.update_BCF_Change_modeclose)

        self.B_aK_TK_edit_label = QLabel("B-α,K-TK")
        self.B_aK_TK_edit = QLineEdit()
        self.B_aK_TK_edit.setFixedSize(100, 25)

        self.B_aK_TK_Start = QLineEdit()
        self.B_aK_TK_Start.setPlaceholderText("TK_Start")
        self.B_aK_TK_End = QLineEdit()
        self.B_aK_TK_End.setPlaceholderText("TK_End")
        self.B_aK_TK_interval = QLineEdit()
        self.B_aK_TK_interval.setPlaceholderText("Interval")

        # Check
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

        # RGB wave and purity
        self.R_wave_checkbox = QCheckBox()
        self.R_wave_checkbox.setChecked(False)  # 預設非勾選狀態
        self.R_wave_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.R_wave_start_label = QLabel("R_wave_Start")
        self.R_wave_start_edit = QLineEdit()
        self.R_wave_end_label = QLabel("R_wave_End")
        self.R_wave_end_edit = QLineEdit()
        self.R_wave_tolerance = QLabel("Tolerance")
        self.R_wave_tolerance_edit = QLineEdit()
        self.R_purity_checkbox = QCheckBox()
        self.R_purity_checkbox.setChecked(False)  # 預設非勾選狀態
        self.R_purity_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.R_purity_limit_label = QLabel("R_purity_limit")
        self.R_purity_limit_edit = QLineEdit()

        self.G_wave_checkbox = QCheckBox()
        self.G_wave_checkbox.setChecked(False)  # 預設非勾選狀態
        self.G_wave_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.G_wave_start_label = QLabel("G_wave_Start")
        self.G_wave_start_edit = QLineEdit()
        self.G_wave_end_label = QLabel("G_wave_End")
        self.G_wave_end_edit = QLineEdit()
        self.G_wave_tolerance = QLabel("Tolerance")
        self.G_wave_tolerance_edit = QLineEdit()
        self.G_purity_checkbox = QCheckBox()
        self.G_purity_checkbox.setChecked(False)  # 預設非勾選狀態
        self.G_purity_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.G_purity_limit_label = QLabel("G_purity_limit")
        self.G_purity_limit_edit = QLineEdit()

        self.B_wave_checkbox = QCheckBox()
        self.B_wave_checkbox.setChecked(False)  # 預設非勾選狀態
        self.B_wave_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.B_wave_start_label = QLabel("B_wave_Start")
        self.B_wave_start_edit = QLineEdit()
        self.B_wave_end_label = QLabel("B_wave_End")
        self.B_wave_end_edit = QLineEdit()
        self.B_wave_tolerance = QLabel("Tolerance")
        self.B_wave_tolerance_edit = QLineEdit()
        self.B_purity_checkbox = QCheckBox()
        self.B_purity_checkbox.setChecked(False)  # 預設非勾選狀態
        self.B_purity_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.B_purity_limit_label = QLabel("B_purity_limit")
        self.B_purity_limit_edit = QLineEdit()

        # Thickness select
        self.RGB_1_checkbox = QCheckBox()
        self.RGB_1_checkbox.setChecked(False)  # 預設非勾選狀態
        self.RGB_1_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.RGB_1_label = QLabel("R=G=B")
        # self.RGB_1_tolerance = QLabel("Thickness_Gap")
        # self.RGB_1_tolerance_edit = QLineEdit()

        self.RGB_2_checkbox = QCheckBox()
        self.RGB_2_checkbox.setChecked(False)  # 預設非勾選狀態
        self.RGB_2_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.RGB_2_label = QLabel("R=G>B")
        # self.RGB_2_tolerance = QLabel("Thickness_Gap")
        # self.RGB_2_tolerance_edit = QLineEdit()

        self.RGB_3_checkbox = QCheckBox()
        self.RGB_3_checkbox.setChecked(False)  # 預設非勾選狀態
        self.RGB_3_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.RGB_3_label = QLabel("R>G=B")
        # self.RGB_3_tolerance = QLabel("Thickness_Gap")
        # self.RGB_3_tolerance_edit = QLineEdit()

        self.RGB_4_checkbox = QCheckBox()
        self.RGB_4_checkbox.setChecked(False)  # 預設非勾選狀態
        self.RGB_4_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.RGB_4_label = QLabel("R>G>B")
        # self.RGB_4_tolerance = QLabel("Thickness_Gap")
        # self.RGB_4_tolerance_edit = QLineEdit()

        self.RGB_5_checkbox = QCheckBox()
        self.RGB_5_checkbox.setChecked(False)  # 預設非勾選狀態
        self.RGB_5_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.RGB_5_label = QLabel("R<G=B")
        # self.RGB_5_tolerance = QLabel("Thickness_Gap")
        # self.RGB_5_tolerance_edit = QLineEdit()

        self.RGB_6_checkbox = QCheckBox()
        self.RGB_6_checkbox.setChecked(False)  # 預設非勾選狀態
        self.RGB_6_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.RGB_6_label = QLabel("R<G<B")
        # self.RGB_6_tolerance = QLabel("Thickness_Gap")
        # self.RGB_6_tolerance_edit = QLineEdit()

        self.RGB_7_checkbox = QCheckBox()
        self.RGB_7_checkbox.setChecked(False)  # 預設非勾選狀態
        self.RGB_7_checkbox.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.RGB_7_label = QLabel("R=G<B")
        # self.RGB_7_tolerance = QLabel("Thickness_Gap")
        # self.RGB_7_tolerance_edit = QLineEdit()

        # Free select range
        self.R_TK_check = QCheckBox()
        self.R_TK_check.setChecked(False)
        self.R_TK_check.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.R_TK_check_start = QLabel("R_TK_Start")
        self.R_TK_start_edit = QLineEdit()
        self.R_TK_check_end = QLabel("R_TK_End")
        self.R_TK_end_edit = QLineEdit()

        self.G_TK_check = QCheckBox()
        self.G_TK_check.setChecked(False)
        self.G_TK_check.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.G_TK_check_start = QLabel("G_TK_Start")
        self.G_TK_start_edit = QLineEdit()
        self.G_TK_check_end = QLabel("G_TK_End")
        self.G_TK_end_edit = QLineEdit()

        self.B_TK_check = QCheckBox()
        self.B_TK_check.setChecked(False)
        self.B_TK_check.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.B_TK_check_start = QLabel("B_TK_Start")
        self.B_TK_start_edit = QLineEdit()
        self.B_TK_check_end = QLabel("B_TK_End")
        self.B_TK_end_edit = QLineEdit()

        # NTSC
        self.NTSC_check = QCheckBox()
        self.NTSC_check.setChecked(False)
        self.NTSC_check.setStyleSheet("QCheckBox { font-size: 32px; color: black; }")
        self.NTSC_check_label = QLabel("NTSC_Limit")
        self.NTSC_check_edit = QLineEdit()

        # table
        self.color_sim_table = QTableWidget()
        self.color_sim_table.setColumnCount(40)
        self.color_sim_table.setRowCount(800)
        # self.color_sim_table.setFixedSize(1000, 600)
        # Apply styles to the dialog
        self.color_sim_table.setStyleSheet("""
                                            QTableWidget::item:selected {
                                        color: blcak; /* 設定文字顏色為黑色 */
                                        background-color: #008080; /* 設定背景顏色為藍色，你可以根據需要調整 */
                                            }
                                        """)
        # Header
        self.color_sim_table.setHorizontalHeaderLabels(["項目", "Wx", "Wy", "WY", "Rx", "Ry", "RY", "Rλ", "R_Purity", "Gx", "Gy", "GY",
                                  "Gλ", "G_Purity", "Bx", "By", "BY", "B_λ", "B_Purity", "NTSC%", "BLUx", "BLUy",
                                  "R色阻選擇", "R色阻厚度", "G色阻選擇", "G色阻厚度", "B色阻選擇", "B色阻厚度",
                                  "背光選擇", "Layer_1", "Layer_2", "Layer_3", "Layer_4",
                                  "Layer_5", "Layer_6"
                                  ])

        # 設定默認值
        column1_default_values = ["項目", "Wx", "Wy", "WY", "Rx", "Ry", "RY", "Rλ", "R_Purity", "Gx", "Gy", "GY",
                                  "Gλ", "G_Purity", "Bx", "By", "BY", "B_λ", "B_Purity", "NTSC%", "BLUx", "BLUy",
                                  "R色阻選擇", "R色阻厚度", "G色阻選擇", "G色阻厚度", "B色阻選擇", "B色阻厚度",
                                  "背光選擇", "Layer_1", "Layer_2", "Layer_3", "Layer_4",
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

        # Simulate_button---------------------------------------------
        self.simbutton = QPushButton("Simulation_Fix")
        self.simbutton.clicked.connect(self.calculate_color_W_Fix_customize_sim)
        self.copy_table = QPushButton("Copy_Entire_Form")
        self.copy_table.clicked.connect(self.copy_entire_table_content)
        self.Simulation_Change = QPushButton("Simulation_Change")
        self.Simulation_Change.clicked.connect(self.calculate_color_W_Change_customize_sim)

        # 放置
        self.Color_Select_layout.addWidget(self.CS_light_source_label, 0, 0)
        self.Color_Select_layout.addWidget(self.CS_light_source_datatable, 0, 1)
        self.Color_Select_layout.addWidget(self.CS_light_source, 1, 0, 1, 2)
        # layer
        self.Color_Select_layout.addWidget(self.CS_layer1_mode, 0, 2)
        self.Color_Select_layout.addWidget(self.CS_layer1_table, 0, 3)
        self.Color_Select_layout.addWidget(self.CS_layer1_box, 1, 2, 1, 2)

        self.Color_Select_layout.addWidget(self.CS_layer2_mode, 0, 4)
        self.Color_Select_layout.addWidget(self.CS_layer2_table, 0, 5)
        self.Color_Select_layout.addWidget(self.CS_layer2_box, 1, 4, 1, 2)

        self.Color_Select_layout.addWidget(self.CS_layer3_mode, 0, 6)
        self.Color_Select_layout.addWidget(self.CS_layer3_table, 0, 7)
        self.Color_Select_layout.addWidget(self.CS_layer3_box, 1, 6, 1, 2)

        self.Color_Select_layout.addWidget(self.CS_layer4_mode, 2, 0)
        self.Color_Select_layout.addWidget(self.CS_layer4_table, 2, 1)
        self.Color_Select_layout.addWidget(self.CS_layer4_box, 3, 0, 1, 2)

        self.Color_Select_layout.addWidget(self.CS_layer5_mode, 2, 2)
        self.Color_Select_layout.addWidget(self.CS_layer5_table, 2, 3)
        self.Color_Select_layout.addWidget(self.CS_layer5_box, 3, 2, 1, 2)

        self.Color_Select_layout.addWidget(self.CS_layer6_mode, 2, 4)
        self.Color_Select_layout.addWidget(self.CS_layer6_table, 2, 5)
        self.Color_Select_layout.addWidget(self.CS_layer6_box, 3, 4, 1, 2)



        # RGB-Fix區域---------------------------------------------------
        self.Color_RGB_widget = QWidget()
        self.Color_RGB_layout = QGridLayout()
        self.Color_RGB_layout.addWidget(self.RGB_fix_label, 0, 0, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_fix_mode, 1, 0, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_fix_box, 1, 1, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_fix_table, 0, 1, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_TK_edit_label, 0, 2, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_TK_label, 1, 2, alignment=Qt.AlignTop)

        self.Color_RGB_layout.addWidget(self.G_fix_mode, 1, 3, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.G_fix_box, 1, 4, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.G_fix_table, 0, 4, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.G_TK_edit_label, 0, 5, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.G_TK_label, 1, 5, alignment=Qt.AlignTop)

        self.Color_RGB_layout.addWidget(self.B_fix_mode, 1, 6, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.B_fix_box, 1, 7, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.B_fix_table, 0, 7, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.B_TK_edit_label, 0, 8, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.B_TK_label, 1, 8, alignment=Qt.AlignTop)

        # RGB-α,K區域---------------------------------------------------
        self.Color_RGB_layout.addWidget(self.RGB_aK_label, 2, 0, alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_aK_mode, 3, 0,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_aK_box, 3, 1,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_aK_table, 2, 1,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_aK_TK_edit_label, 2, 2,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_aK_TK_edit, 3, 2,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_aK_TK_Start, 4, 0,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_aK_TK_End, 4, 1,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.R_aK_TK_interval, 4, 2,alignment=Qt.AlignTop)

        self.Color_RGB_layout.addWidget(self.G_aK_mode, 3, 3,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.G_aK_box, 3, 4,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.G_aK_table, 2, 4,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.G_aK_TK_edit_label, 2, 5,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.G_aK_TK_edit, 3, 5,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.G_aK_TK_Start, 4, 3,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.G_aK_TK_End, 4, 4,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.G_aK_TK_interval, 4, 5,alignment=Qt.AlignTop)

        self.Color_RGB_layout.addWidget(self.B_aK_mode, 3, 6,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.B_aK_box, 3, 7,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.B_aK_table, 2, 7,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.B_aK_TK_edit_label, 2, 8,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.B_aK_TK_edit, 3, 8,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.B_aK_TK_Start, 4, 6,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.B_aK_TK_End, 4, 7,alignment=Qt.AlignTop)
        self.Color_RGB_layout.addWidget(self.B_aK_TK_interval, 4, 8,alignment=Qt.AlignTop)

        self.Color_Check_widget = QWidget()
        self.Color_Check_layout = QGridLayout()
        # check
        self.Color_Check_layout.addWidget(self.W_checkbox, 0, 0,alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Wx_label, 0, 1,alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Wx_edit,0,2,alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Wy_label, 0, 3,alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Wy_edit, 0, 4,alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.W_tolerance, 0, 5,alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.W_tolerance_edit, 0, 6,alignment=Qt.AlignTop | Qt.AlignLeft)

        self.Color_Check_layout.addWidget(self.R_checkbox, 0, 7, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Rx_label, 0, 8, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Rx_edit, 0, 9, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Ry_label, 0, 10, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Ry_edit, 0, 11, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.R_tolerance, 0, 12, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.R_tolerance_edit, 0, 13, alignment=Qt.AlignTop | Qt.AlignLeft)

        self.Color_Check_layout.addWidget(self.G_checkbox, 0, 14, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Gx_label, 0, 15, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Gx_edit, 0, 16, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Gy_label, 0, 17, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Gy_edit, 0, 18, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.G_tolerance, 0, 19, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.G_tolerance_edit, 0, 20, alignment=Qt.AlignTop | Qt.AlignLeft)

        self.Color_Check_layout.addWidget(self.B_checkbox, 0, 21, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Bx_label, 0, 22, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.Bx_edit, 0, 23, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.By_label, 0, 24, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.By_edit, 0, 25, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.B_tolerance, 0, 26, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Color_Check_layout.addWidget(self.B_tolerance_edit, 0, 27, alignment=Qt.AlignTop | Qt.AlignLeft)

        # RGB wave and purity
        self.Wave_Check_widget = QWidget()
        self.Wave_Check_layout = QGridLayout()

        self.Wave_Check_layout.addWidget(self.R_wave_checkbox, 0, 0, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.R_wave_start_label, 0, 1, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.R_wave_start_edit, 0, 2, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.R_wave_end_label, 0, 3, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.R_wave_end_edit, 0, 4, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.R_purity_checkbox, 0, 5, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.R_purity_limit_label, 0, 6, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.R_purity_limit_edit, 0, 7, alignment=Qt.AlignTop | Qt.AlignLeft)

        self.Wave_Check_layout.addWidget(self.G_wave_checkbox, 0, 8, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.G_wave_start_label, 0, 9, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.G_wave_start_edit, 0, 10, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.G_wave_end_label, 0, 11, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.G_wave_end_edit, 0, 12, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.G_purity_checkbox, 0, 13, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.G_purity_limit_label, 0, 14, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.G_purity_limit_edit, 0, 15, alignment=Qt.AlignTop | Qt.AlignLeft)

        self.Wave_Check_layout.addWidget(self.B_wave_checkbox, 0, 16, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.B_wave_start_label, 0, 17, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.B_wave_start_edit, 0, 18, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.B_wave_end_label, 0, 19, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.B_wave_end_edit, 0, 20, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.B_purity_checkbox, 0, 21, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.B_purity_limit_label, 0, 22, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.B_purity_limit_edit, 0, 23, alignment=Qt.AlignTop | Qt.AlignLeft)
        # NTSC
        self.Wave_Check_layout.addWidget(self.NTSC_check, 0, 24, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.NTSC_check_label, 0, 25, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.Wave_Check_layout.addWidget(self.NTSC_check_edit, 0, 26, alignment=Qt.AlignTop | Qt.AlignLeft)

        # TK
        self.TK_Check_widget = QWidget()
        self.TK_Check_layout = QGridLayout()
        self.TK_Check_layout.addWidget(self.RGB_1_checkbox, 0, 0, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_1_label, 0, 1, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_2_checkbox, 0, 2, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_2_label, 0, 3, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_3_checkbox, 0, 4, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_3_label, 0, 5, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_4_checkbox, 0, 6, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_4_label, 0, 7, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_5_checkbox, 0, 8, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_5_label, 0, 9, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_6_checkbox, 0, 10, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_6_label, 0, 11, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_7_checkbox, 0, 12, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.RGB_7_label, 0, 13, alignment=Qt.AlignTop | Qt.AlignLeft)

        # Free TK
        self.TK_Check_layout.addWidget(self.R_TK_check, 0, 14, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.R_TK_check_start, 0, 15, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.R_TK_start_edit, 0, 16, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.R_TK_check_end, 0, 17, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.R_TK_end_edit, 0, 18, alignment=Qt.AlignTop | Qt.AlignLeft)

        self.TK_Check_layout.addWidget(self.G_TK_check, 0, 19, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.G_TK_check_start, 0, 20, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.G_TK_start_edit, 0, 21, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.G_TK_check_end, 0, 22, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.G_TK_end_edit, 0, 23, alignment=Qt.AlignTop | Qt.AlignLeft)

        self.TK_Check_layout.addWidget(self.B_TK_check, 0, 24, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.B_TK_check_start, 0, 25, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.B_TK_start_edit, 0, 26, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.B_TK_check_end, 0, 27, alignment=Qt.AlignTop | Qt.AlignLeft)
        self.TK_Check_layout.addWidget(self.B_TK_end_edit, 0, 28, alignment=Qt.AlignTop | Qt.AlignLeft)



        # Sim_table
        self.Color_Sim_Table_widget = QWidget()
        self.Color_Sim_Table_layout = QGridLayout()
        self.Color_Sim_Table_layout.addWidget(self.simbutton,0,0,1,3)
        self.Color_Sim_Table_layout.addWidget(self.copy_table,0,6,1,3)
        self.Color_Sim_Table_layout.addWidget(self.Simulation_Change,0,3,1,3)
        self.Color_Sim_Table_layout.addWidget(self.color_sim_table,1,0,20,12)

        self.Color_Select_layout.setSpacing(1)

        # 設置新 widget 的布局
        self.Color_Check_widget.setLayout(self.Color_Check_layout)
        self.Color_RGB_widget.setLayout(self.Color_RGB_layout)
        self.Color_Sim_Table_widget.setLayout(self.Color_Sim_Table_layout)
        self.Wave_Check_widget.setLayout(self.Wave_Check_layout)
        self.TK_Check_widget.setLayout(self.TK_Check_layout)

        # 在主布局中添加新 widget
        self.Color_Select_layout.addWidget(self.Color_Check_widget, 14, 0,4,8)
        self.Color_Select_layout.addWidget(self.Wave_Check_widget, 16, 0,4,8)
        self.Color_Select_layout.addWidget(self.TK_Check_widget, 18, 0,4,8)
        self.Color_Select_layout.addWidget(self.Color_RGB_widget, 4, 0,1,8)
        self.Color_Select_layout.addWidget(self.Color_Sim_Table_widget,20,0,20,8)

        # 設定widget樣式
        # # 為 widget 設置背景顏色和邊框
        # self.Color_Check_widget.setStyleSheet("""
        #             QWidget {
        #                 background-color: #CCCCFF;  # 淺灰色背景
        #                 border: 1px solid black;  # 黑色邊框
        #             }
        #         """)
        #
        # self.Color_RGB_widget.setStyleSheet("""
        #                     QWidget {
        #                         background-color: #FF8033;  # 淺灰色背景
        #                         border: 1px solid black;  # 黑色邊框
        #                     }
        #                 """)

        # 設置主 widget 的布局
        self.setLayout(self.Color_Select_layout)
        # # 連接全局信號到相應的方法
        global_signal_manager.databaseUpdated.connect(self.update_CS_light_source_datatable)
        global_signal_manager.databaseUpdated.connect(self.update_CS_light_sourceComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer1_datatable)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer1ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer2_datatable)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer2ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer3_datatable)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer3ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer4_datatable)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer4ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer5_datatable)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer5ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer6_datatable)
        global_signal_manager.databaseUpdated.connect(self.update_CS_layer6ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_RCF_Fix_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateRCF_Fix_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_GCF_Fix_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateGCF_Fix_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_BCF_Fix_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateBCF_Fix_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_RCF_Change_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateRCF_Change_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_GCF_Change_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateGCF_Change_ComboBox)
        global_signal_manager.databaseUpdated.connect(self.update_BCF_Change_datatable)
        global_signal_manager.databaseUpdated.connect(self.updateBCF_Change_ComboBox)


#--------------------Function------------------------------------------------------
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

    # 複製整個表格的內容
    def copy_entire_table_content(self):
        row_count = self.color_sim_table.rowCount()
        column_count = self.color_sim_table.columnCount()

        copied_data = []
        for row in range(row_count):
            row_data = []
            for col in range(column_count):
                item = self.color_sim_table.item(row, col)
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
    def update_CS_light_source_datatable(self):

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
        self.CS_light_source_datatable.clear()
        for table in tables:
            self.CS_light_source_datatable.addItem(table[0])

        # 關閉連線
        conn.close()


    def update_CS_light_sourceComboBox(self):

        connection = sqlite3.connect("blu_database.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.CS_light_source_datatable.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.CS_light_source.clear()
        for item in header_labels:
            self.CS_light_source.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_CS_layer1_datatable(self):
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
        self.CS_layer1_table.clear()
        for table in tables:
            self.CS_layer1_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def update_CS_layer1ComboBox(self):

        connection = sqlite3.connect("cell_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.CS_layer1_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.CS_layer1_box.clear()
        for item in header_labels:
            self.CS_layer1_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_CS_layer2_datatable(self):

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
        self.CS_layer2_table.clear()
        for table in tables:
            self.CS_layer2_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def update_CS_layer2ComboBox(self):

        connection = sqlite3.connect("BSITO_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.CS_layer2_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.CS_layer2_box.clear()
        for item in header_labels:
            self.CS_layer2_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_CS_layer3_datatable(self):

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
        self.CS_layer3_table.clear()
        for table in tables:
            self.CS_layer3_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def update_CS_layer3ComboBox(self):

        connection = sqlite3.connect("Layer3_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.CS_layer3_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.CS_layer3_box.clear()
        for item in header_labels:
            self.CS_layer3_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_CS_layer4_datatable(self):

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
        self.CS_layer4_table.clear()
        for table in tables:
            self.CS_layer4_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def update_CS_layer4ComboBox(self):

        connection = sqlite3.connect("Layer4_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.CS_layer4_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.CS_layer4_box.clear()
        for item in header_labels:
            self.CS_layer4_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_CS_layer5_datatable(self):

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
        self.CS_layer5_table.clear()
        for table in tables:
            self.CS_layer5_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def update_CS_layer5ComboBox(self):

        connection = sqlite3.connect("Layer5_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.CS_layer5_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.CS_layer5_box.clear()
        for item in header_labels:
            self.CS_layer5_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    def update_CS_layer6_datatable(self):

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
        self.CS_layer6_table.clear()
        for table in tables:
            self.CS_layer6_table.addItem(table[0])

        # 關閉連線
        conn.close()

    def update_CS_layer6ComboBox(self):

        connection = sqlite3.connect("Layer5_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.CS_layer6_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
        self.CS_layer6_box.clear()
        for item in header_labels:
            self.CS_layer6_box.addItem(str(item))
            # print("item", str(item))
        # 關閉連線
        connection.close()

    # RCF_Fix
    def update_RCF_Fix_datatable(self):
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


    def updateRCF_Fix_ComboBox(self):
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

    # def update_RCF_Fix_modeclose(self):
    #     if self.R_fix_mode.currentText() == "自訂" or self.R_fix_mode.currentText() == "模擬":
    #         self.R_aK_mode.setCurrentText("未選")

    # GCF_Fix
    def update_GCF_Fix_datatable(self):
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

    def updateGCF_Fix_ComboBox(self):
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

    # def update_GCF_Fix_modeclose(self):
    #     if self.G_fix_mode.currentText() == "自訂" or self.G_fix_mode.currentText() == "模擬":
    #         self.G_aK_mode.setCurrentText("未選")


    # BCF_Fix
    def update_BCF_Fix_datatable(self):
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

    def updateBCF_Fix_ComboBox(self):
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

    # def update_BCF_Fix_modeclose(self):
    #     if self.B_fix_mode.currentText() == "自訂" or self.B_fix_mode.currentText() == "模擬":
    #         self.B_aK_mode.setCurrentText("未選")


    # RCF_Change
    def update_RCF_Change_datatable(self):
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

    def updateRCF_Change_ComboBox(self):
        # # 防止更新不停Reload
        # self.light_source.setEnabled(False)
        # self.light_source_datatable.setEnabled(False)
        connection = sqlite3.connect("RCF_Change_spectrum.db")
        cursor = connection.cursor()

        # 获取表格的標題
        cursor.execute(f"PRAGMA table_info('{self.R_aK_table.currentText()}');")
        header_data = cursor.fetchall()
        header_labels = [column[1] for column in header_data]
        # print("headerlabels-from source", header_labels)
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

    # def update_RCF_Change_modeclose(self):
    #     if self.R_aK_mode.currentText() == "自訂" or self.R_aK_mode.currentText() == "模擬":
    #         self.R_fix_mode.setCurrentText("未選")


    # GCF_Change
    def update_GCF_Change_datatable(self):
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

    def updateGCF_Change_ComboBox(self):
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

    # def update_GCF_Change_modeclose(self):
    #     if self.G_aK_mode.currentText() == "自訂" or self.G_aK_mode.currentText() == "模擬":
    #         self.G_fix_mode.setCurrentText("未選")


    # BCF_Change
    def update_BCF_Change_datatable(self):
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

    def updateBCF_Change_ComboBox(self):
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

    # def update_BCF_Change_modeclose(self):
    #     if self.B_aK_mode.currentText() == "自訂" or self.B_aK_mode.currentText() == "模擬":
    #         self.B_fix_mode.setCurrentText("未選")

    # original_cal
    def calculate_BLU(self):
        # BLU
        connection_BLU = sqlite3.connect("blu_database.db")
        cursor_BLU = connection_BLU.cursor()
        # 取得BLU資料
        column_name_BLU = self.CS_light_source.currentText()
        table_name_BLU = self.CS_light_source_datatable.currentText()
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
        # RxSxxl = self.CIE_spectrum_Series_X * self.calculate_BLU()
        # RxSxyl = self.CIE_spectrum_Series_Y * self.calculate_BLU()
        # RxSxzl = self.CIE_spectrum_Series_Z * self.calculate_BLU()
        # k = 100 / BLU_spectrum_Series.sum()
        # RxSxxl_sum = RxSxxl.sum()
        # RxSxyl_sum = RxSxyl.sum()
        # RxSxzl_sum = RxSxzl.sum()
        # RxSxxl_sum_k = RxSxxl_sum * k
        # RxSxyl_sum_k = RxSxyl_sum * k
        # RxSxzl_sum_k = RxSxzl_sum * k
        # self.BLU_x = RxSxxl_sum_k / (RxSxxl_sum_k + RxSxyl_sum_k + RxSxzl_sum_k)
        # self.BLU_y = RxSxyl_sum_k / (RxSxxl_sum_k + RxSxyl_sum_k + RxSxzl_sum_k)
        # 自訂BLU_Spectrum回傳
        return BLU_spectrum_Series


    # Cell
    def calculate_layer1(self):
        if self.CS_layer1_mode.currentText() == "layer1_自訂":

            print("in_layer1_自訂")
            connection_layer1 = sqlite3.connect("cell_spectrum.db")
            cursor_layer1 = connection_layer1.cursor()
            # 取得BLU資料
            column_name_layer1 = self.CS_layer1_box.currentText()
            table_name_layer1 = self.CS_layer1_table.currentText()
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
            # 自訂BLU_Spectrum回傳
            return layer1_spectrum_Series
        else:

            layer1_spectrum_Series = 1
            print("layer1:未選")
            return layer1_spectrum_Series

    # BSITO
    def calculate_layer2(self):
        if self.CS_layer2_mode.currentText() == "layer2_自訂":

            print("in_layer2_自訂")
            connection_layer2 = sqlite3.connect("BSITO_spectrum.db")
            cursor_layer2 = connection_layer2.cursor()
            # 取得BLU資料
            column_name_layer2 = self.CS_layer2_box.currentText()
            table_name_layer2 = self.CS_layer2_table.currentText()
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

    def calculate_layer3(self):
        if self.CS_layer3_mode.currentText() == "layer3_自訂":

            print("in_layer3_自訂")
            connection_layer3 = sqlite3.connect("Layer3_spectrum.db")
            cursor_layer3 = connection_layer3.cursor()
            # 取得BLU資料
            column_name_layer3 = self.CS_layer3_box.currentText()
            table_name_layer3 = self.CS_layer3_table.currentText()
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

    def calculate_layer4(self):
        if self.CS_layer4_mode.currentText() == "layer4_自訂":

            print("in_layer4_自訂")
            connection_layer4 = sqlite3.connect("Layer4_spectrum.db")
            cursor_layer4 = connection_layer4.cursor()
            # 取得BLU資料
            column_name_layer4 = self.CS_layer4_box.currentText()
            table_name_layer4 = self.CS_layer4_table.currentText()
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

    def calculate_layer5(self):
        if self.CS_layer5_mode.currentText() == "layer5_自訂":

            print("in_layer5_自訂")
            connection_layer5 = sqlite3.connect("Layer5_spectrum.db")
            cursor_layer5 = connection_layer5.cursor()
            # 取得BLU資料
            column_name_layer5 = self.CS_layer5_box.currentText()
            table_name_layer5 = self.CS_layer5_table.currentText()
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

    def calculate_layer6(self):
        if self.CS_layer6_mode.currentText() == "layer6_自訂":

            print("in_layer6_自訂")
            connection_layer6 = sqlite3.connect("Layer6_spectrum.db")
            cursor_layer6 = connection_layer6.cursor()
            # 取得BLU資料
            column_name_layer6 = self.CS_layer6_box.currentText()
            table_name_layer6 = self.CS_layer6_table.currentText()
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


    # sim
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
        self.BLU_x = RxSxxl_sum_k / (RxSxxl_sum_k + RxSxyl_sum_k + RxSxzl_sum_k)
        self.BLU_y = RxSxyl_sum_k / (RxSxxl_sum_k + RxSxyl_sum_k + RxSxzl_sum_k)

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
            print("R_fix_sim_spectrum_list-模擬",R_fix_sim_spectrum_list)
            self.R_Fix_X_list = [series.multiply(self.CIE_spectrum_Series_X) for series in
                                       R_fix_sim_spectrum_list]
            print("self.R_Fix_X_list-模擬",self.R_Fix_X_list)
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
                self.R_Fix_T_sim.append((((R_fix_sim_spectrum_list[i] / self.calculate_BLU()) * self.CIE_spectrum_Series_Y).sum())/((self.calculate_BLU() * self.CIE_spectrum_Series_Y).sum())*100)

            print("self.R_Fix_x_sim", self.R_Fix_x_sim)
            print("self.R_Fix_y_sim", self.R_Fix_y_sim)
            print("self.R_Fix_T_sim-模擬", self.R_Fix_T_sim)
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
        elif self.R_fix_mode.currentText() == "自訂":
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
            self.R_FIX_TK_sim_list = [self.RTK]
            self.RCF_Fix_Sim_headers_use = [self.R_fix_box.currentText()]
            self.R_TK_label.setText(f"{self.RTK:.3f} um")

            # 在計算完 C_Series 後，加上以下代碼
            new_RCF_Fix_spectrum_Series = new_RCF_Fix_spectrum_Series.reset_index(drop=True)
            print("new_RCF_Fix_spectrum_Series",new_RCF_Fix_spectrum_Series)
            R_fix_sim_spectrum_list = new_RCF_Fix_spectrum_Series * self.cell_blu_total_spectrum

            # 開始計算新的R座標
            self.R_Fix_X_list = [R_fix_sim_spectrum_list * self.CIE_spectrum_Series_X]
            print("self.R_Fix_X_list-自訂",self.R_Fix_X_list)
            self.R_Fix_Y_list = [R_fix_sim_spectrum_list * self.CIE_spectrum_Series_Y]
            print("self.R_Fix_Y_list-自訂", self.R_Fix_Y_list)
            self.R_Fix_Z_list = [R_fix_sim_spectrum_list * self.CIE_spectrum_Series_Z]
            self.R_X_fix_sim_sum_list = [self.R_Fix_X_list[0].sum()]
            self.R_Y_fix_sim_sum_list = [self.R_Fix_Y_list[0].sum()]
            self.R_Z_fix_sim_sum_list = [self.R_Fix_Z_list[0].sum()]
            self.R_Fix_x_sim = [self.R_X_fix_sim_sum_list[0]/(self.R_X_fix_sim_sum_list[0] + self.R_Y_fix_sim_sum_list[0] + self.R_Z_fix_sim_sum_list[0])]
            print("self.R_Fix_x_sim",self.R_Fix_x_sim)
            self.R_Fix_y_sim = [self.R_Y_fix_sim_sum_list[0]/(self.R_X_fix_sim_sum_list[0] + self.R_Y_fix_sim_sum_list[0] + self.R_Z_fix_sim_sum_list[0])]
            print("self.R_Fix_y_sim", self.R_Fix_y_sim)
            self.R_Fix_T_sim = [((((R_fix_sim_spectrum_list / self.calculate_BLU())* self.CIE_spectrum_Series_Y).sum())/((self.calculate_BLU() * self.CIE_spectrum_Series_Y).sum()))*100]
            print("self.R_Fix_T_sim-自訂",self.R_Fix_T_sim)
            xy = [self.R_Fix_x_sim[0], self.R_Fix_y_sim[0]]
            xy_n = self.observer_D65
            # 計算主波長(
            dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
            dominant_wavelength_value = dominant_wavelength
            CIEcoordinate_value_1 = CIEcoordinate[0]
            CIEcoordinate_value_2 = CIEcoordinate[1]
            # 計算距離
            distance_from_white = math.sqrt((self.R_Fix_x_sim[0] - xy_n[0]) ** 2 + (self.R_Fix_y_sim[0] - xy_n[1]) ** 2)
            distance_from_black = math.sqrt(
                (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
            # 計算色純度
            purity = distance_from_white / distance_from_black * 100
            wave = str(dominant_wavelength_value)
            self.R_Fix_wave_sim = [wave]
            self.R_Fix_purity_sim = [purity]
            print("self.R_Fix_wave_sim", self.R_Fix_wave_sim)
            print("self.R_Fix_purity_sim", self.R_Fix_purity_sim)
            # 關閉連線
            connection_RCF_Fix.close()


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
                self.G_Fix_T_sim.append((((G_fix_sim_spectrum_list[i]/self.calculate_BLU()) * self.CIE_spectrum_Series_Y).sum())/((self.calculate_BLU() * self.CIE_spectrum_Series_Y).sum())*100)

            print("self.G_Fix_x_sim", self.G_Fix_x_sim)
            print("self.G_Fix_y_sim", self.G_Fix_y_sim)
            print("self.G_Fix_T_sim-模擬", self.G_Fix_T_sim)
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
        elif self.G_fix_mode.currentText() == "自訂":
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
            self.G_FIX_TK_sim_list = [self.GTK]
            self.GCF_Fix_Sim_headers_use = [self.G_fix_box.currentText()]
            self.G_TK_label.setText(f"{self.GTK:.3f} um")

            # 在計算完 C_Series 後，加上以下代碼
            new_GCF_Fix_spectrum_Series = new_GCF_Fix_spectrum_Series.reset_index(drop=True)
            print("new_GCF_Fix_spectrum_Series", new_GCF_Fix_spectrum_Series)
            G_fix_sim_spectrum_list = new_GCF_Fix_spectrum_Series * self.cell_blu_total_spectrum

            # 開始計算新的R座標
            self.G_Fix_X_list = [G_fix_sim_spectrum_list * self.CIE_spectrum_Series_X]
            print("self.G_Fix_X_list-自訂", self.G_Fix_X_list)
            self.G_Fix_Y_list = [G_fix_sim_spectrum_list * self.CIE_spectrum_Series_Y]
            print("self.G_Fix_Y_list-自訂", self.G_Fix_Y_list)
            self.G_Fix_Z_list = [G_fix_sim_spectrum_list * self.CIE_spectrum_Series_Z]
            self.G_X_fix_sim_sum_list = [self.G_Fix_X_list[0].sum()]
            self.G_Y_fix_sim_sum_list = [self.G_Fix_Y_list[0].sum()]
            self.G_Z_fix_sim_sum_list = [self.G_Fix_Z_list[0].sum()]
            self.G_Fix_x_sim = [self.G_X_fix_sim_sum_list[0] / (
                        self.G_X_fix_sim_sum_list[0] + self.G_Y_fix_sim_sum_list[0] + self.G_Z_fix_sim_sum_list[0])]
            print("self.G_Fix_x_sim", self.G_Fix_x_sim)
            self.G_Fix_y_sim = [self.G_Y_fix_sim_sum_list[0] / (
                        self.G_X_fix_sim_sum_list[0] + self.G_Y_fix_sim_sum_list[0] + self.G_Z_fix_sim_sum_list[0])]
            print("self.G_Fix_y_sim", self.G_Fix_y_sim)
            self.G_Fix_T_sim = [((((G_fix_sim_spectrum_list / self.calculate_BLU())* self.CIE_spectrum_Series_Y).sum())/((self.calculate_BLU() * self.CIE_spectrum_Series_Y).sum()))*100]
            print("self.G_Fix_T_sim-自訂",self.G_Fix_T_sim)
            xy = [self.G_Fix_x_sim[0], self.G_Fix_y_sim[0]]
            xy_n = self.observer_D65
            # 計算主波長(
            dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
            dominant_wavelength_value = dominant_wavelength
            CIEcoordinate_value_1 = CIEcoordinate[0]
            CIEcoordinate_value_2 = CIEcoordinate[1]
            # 計算距離
            distance_from_white = math.sqrt((self.G_Fix_x_sim[0] - xy_n[0]) ** 2 + (self.G_Fix_y_sim[0] - xy_n[1]) ** 2)
            distance_from_black = math.sqrt(
                (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
            # 計算色純度
            purity = distance_from_white / distance_from_black * 100
            wave = str(dominant_wavelength_value)
            self.G_Fix_wave_sim = [wave]
            self.G_Fix_purity_sim = [purity]
            print("self.G_Fix_wave_sim", self.G_Fix_wave_sim)
            print("self.G_Fix_purity_sim", self.G_Fix_purity_sim)
            # 關閉連線
            connection_GCF_Fix.close()

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
                self.B_Fix_T_sim.append((((B_fix_sim_spectrum_list[i] /self.calculate_BLU()) * self.CIE_spectrum_Series_Y).sum())/((self.calculate_BLU() * self.CIE_spectrum_Series_Y).sum())*100)

            print("self.B_Fix_x_sim", self.B_Fix_x_sim)
            print("self.B_Fix_y_sim", self.B_Fix_y_sim)
            print("self.B_Fix_T_sim-模擬", self.B_Fix_T_sim)
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

        elif self.B_fix_mode.currentText() == "自訂":
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
            self.B_FIX_TK_sim_list = [self.BTK]
            self.BCF_Fix_Sim_headers_use = [self.B_fix_box.currentText()]
            self.B_TK_label.setText(f"{self.BTK:.3f} um")

            # 在計算完 C_Series 後，加上以下代碼
            new_BCF_Fix_spectrum_Series = new_BCF_Fix_spectrum_Series.reset_index(drop=True)
            print("new_BCF_Fix_spectrum_Series", new_BCF_Fix_spectrum_Series)
            B_fix_sim_spectrum_list = new_BCF_Fix_spectrum_Series * self.cell_blu_total_spectrum

            # 開始計算新的R座標
            self.B_Fix_X_list = [B_fix_sim_spectrum_list * self.CIE_spectrum_Series_X]
            print("self.B_Fix_X_list-自訂", self.B_Fix_X_list)
            self.B_Fix_Y_list = [B_fix_sim_spectrum_list * self.CIE_spectrum_Series_Y]
            print("self.B_Fix_Y_list-自訂", self.B_Fix_Y_list)
            self.B_Fix_Z_list = [B_fix_sim_spectrum_list * self.CIE_spectrum_Series_Z]
            self.B_X_fix_sim_sum_list = [self.B_Fix_X_list[0].sum()]
            self.B_Y_fix_sim_sum_list = [self.B_Fix_Y_list[0].sum()]
            self.B_Z_fix_sim_sum_list = [self.B_Fix_Z_list[0].sum()]
            self.B_Fix_x_sim = [self.B_X_fix_sim_sum_list[0] / (
                        self.B_X_fix_sim_sum_list[0] + self.B_Y_fix_sim_sum_list[0] + self.B_Z_fix_sim_sum_list[0])]
            print("self.B_Fix_x_sim", self.B_Fix_x_sim)
            self.B_Fix_y_sim = [self.B_Y_fix_sim_sum_list[0] / (
                        self.B_X_fix_sim_sum_list[0] + self.B_Y_fix_sim_sum_list[0] + self.B_Z_fix_sim_sum_list[0])]
            print("self.B_Fix_y_sim", self.B_Fix_y_sim)
            self.B_Fix_T_sim = [((((B_fix_sim_spectrum_list / self.calculate_BLU()) * self.CIE_spectrum_Series_Y).sum())/((self.calculate_BLU() * self.CIE_spectrum_Series_Y).sum()))*100]
            print(" self.B_Fix_T_sim-自訂", self.B_Fix_T_sim)
            xy = [self.B_Fix_x_sim[0], self.B_Fix_y_sim[0]]
            xy_n = self.observer_D65
            # 計算主波長(
            dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)
            dominant_wavelength_value = dominant_wavelength
            CIEcoordinate_value_1 = CIEcoordinate[0]
            CIEcoordinate_value_2 = CIEcoordinate[1]
            # 計算距離
            distance_from_white = math.sqrt((self.B_Fix_x_sim[0] - xy_n[0]) ** 2 + (self.B_Fix_y_sim[0] - xy_n[1]) ** 2)
            distance_from_black = math.sqrt(
                (CIEcoordinate_value_1 - xy_n[0]) ** 2 + (CIEcoordinate_value_2 - xy_n[1]) ** 2)
            # 計算色純度
            purity = distance_from_white / distance_from_black * 100
            wave = str(dominant_wavelength_value)
            self.B_Fix_wave_sim = [wave]
            self.B_Fix_purity_sim = [purity]
            print("self.B_Fix_wave_sim", self.B_Fix_wave_sim)
            print("self.B_Fix_purity_sim", self.B_Fix_purity_sim)
            # 關閉連線
            connection_BCF_Fix.close()

    def calculate_color_W_Fix_customize_sim(self):
        # 清除 color_sim_table 從第一行開始以下的所有內容
        for row in range(1, self.color_sim_table.rowCount()):
            for column in range(self.color_sim_table.columnCount()):
                self.color_sim_table.setItem(row, column, QTableWidgetItem(""))
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
        self.W_T_Fix = []
        for i in range(len(all_Fix_W_X_combinations)):
            W_x_total = self.W_X_Fix_sum_list[i]/(self.W_X_Fix_sum_list[i]+self.W_Y_Fix_sum_list[i]+self.W_Z_Fix_sum_list[i])
            self.W_x_Fix_sum_list.append(W_x_total)
        for i in range(len(all_Fix_W_Y_combinations)):
            W_y_total = self.W_Y_Fix_sum_list[i]/(self.W_X_Fix_sum_list[i]+self.W_Y_Fix_sum_list[i]+self.W_Z_Fix_sum_list[i])
            self.W_y_Fix_sum_list.append(W_y_total)
        all_Fix_T_combinations = list(
            itertools.product(self.R_Fix_T_sim, self.G_Fix_T_sim, self.B_Fix_T_sim))
        for i in range(len(all_Fix_T_combinations)):
            self.W_T_Fix.append((all_Fix_T_combinations[i][0]+all_Fix_T_combinations[i][1]+all_Fix_T_combinations[i][2])/3)
        print("self.W_x_Fix_sum_list",self.W_x_Fix_sum_list)
        print("self.W_y_Fix_sum_list",self.W_y_Fix_sum_list)
        print("all_Fix_T_combinations",all_Fix_T_combinations)
        print("self.W_T_Fix",self.W_T_Fix)

        # 每個組合都是一個元組形式 (r, g, b)
        all_Fix_x_combinations = list(itertools.product(self.R_Fix_x_sim, self.G_Fix_x_sim, self.B_Fix_x_sim))
        print("all_Fix_x_combinations", all_Fix_x_combinations)
        all_Fix_y_combinations = list(itertools.product(self.R_Fix_y_sim, self.G_Fix_y_sim, self.B_Fix_y_sim))
        # wave
        all_Fix_wave_combinations = list(
            itertools.product(self.R_Fix_wave_sim, self.G_Fix_wave_sim, self.B_Fix_wave_sim))
        # purity
        all_Fix_purity_combinations = list(
            itertools.product(self.R_Fix_purity_sim, self.G_Fix_purity_sim, self.B_Fix_purity_sim))

        # NTSC list
        self.RGB_Fix_NTSC_List = []
        for i in range(len(all_Fix_purity_combinations)):
            self.RGB_Fix_NTSC_List.append(100 * 0.5 * abs((all_Fix_x_combinations[i][0] * all_Fix_y_combinations[i][1] +
                                                           all_Fix_x_combinations[i][1] * all_Fix_y_combinations[i][2] +
                                                           all_Fix_x_combinations[i][2] * all_Fix_y_combinations[i][
                                                               0] - (all_Fix_x_combinations[i][1] *
                                                                     all_Fix_y_combinations[i][0]) - (
                                                                   all_Fix_x_combinations[i][2] *
                                                                   all_Fix_y_combinations[i][1]) - (
                                                                       all_Fix_x_combinations[i][0] *
                                                                       all_Fix_y_combinations[i][2]))) / 0.1582)

        # 重新設置row count
        self.color_sim_table.setRowCount(len(self.W_x_Fix_sum_list) + 1)
        # 開始放入table content隨便拿一個當數量迴圈
        rows_to_remove = set()  # 初始化集合記錄需要移除的行索引

        for i in range(len(self.W_x_Fix_sum_list)):
            self.color_sim_table.setItem(i + 1, 1, QTableWidgetItem(f"{self.W_x_Fix_sum_list[i]:.4f}"))
            self.color_sim_table.setItem(i + 1, 2, QTableWidgetItem(f"{self.W_y_Fix_sum_list[i]:.4f}"))
            self.color_sim_table.setItem(i + 1, 3, QTableWidgetItem(f"{self.W_T_Fix[i]:.4f}"))
            # RGB_x_Fix_sim
            # 生成所有可能組合的列表
            # 現在，all_combinations 包含所有可能的組合
            # print(len(all_Fix_x_combinations))
            # print(all_Fix_x_combinations[0])
            # print(all_Fix_x_combinations[0][0])
            self.color_sim_table.setItem(i + 1, 4, QTableWidgetItem(f"{all_Fix_x_combinations[i][0]:.4f}"))
            self.color_sim_table.setItem(i + 1, 9, QTableWidgetItem(f"{all_Fix_x_combinations[i][1]:.4f}"))
            self.color_sim_table.setItem(i + 1, 14, QTableWidgetItem(f"{all_Fix_x_combinations[i][2]:.4f}"))

            # RGB_y_Fix_sim

            self.color_sim_table.setItem(i + 1, 5, QTableWidgetItem(f"{all_Fix_y_combinations[i][0]:.4f}"))
            self.color_sim_table.setItem(i + 1, 10, QTableWidgetItem(f"{all_Fix_y_combinations[i][1]:.4f}"))
            self.color_sim_table.setItem(i + 1, 15, QTableWidgetItem(f"{all_Fix_y_combinations[i][2]:.4f}"))
            # RGB_T_Fix_sim
            all_Fix_Y_combinations = list(itertools.product(self.R_Fix_T_sim, self.G_Fix_T_sim, self.B_Fix_T_sim))
            self.color_sim_table.setItem(i + 1, 6, QTableWidgetItem(f"{all_Fix_Y_combinations[i][0]:.4f}"))
            self.color_sim_table.setItem(i + 1, 11, QTableWidgetItem(f"{all_Fix_Y_combinations[i][1]:.4f}"))
            self.color_sim_table.setItem(i + 1, 16, QTableWidgetItem(f"{all_Fix_Y_combinations[i][2]:.4f}"))

            # RGB_wave_Fix_sim
            self.color_sim_table.setItem(i + 1, 7, QTableWidgetItem(f"{all_Fix_wave_combinations[i][0]}"))
            self.color_sim_table.setItem(i + 1, 12, QTableWidgetItem(f"{all_Fix_wave_combinations[i][1]}"))
            self.color_sim_table.setItem(i + 1, 17, QTableWidgetItem(f"{all_Fix_wave_combinations[i][2]}"))

            # RGB_purity_Fix_sim
            self.color_sim_table.setItem(i + 1, 8, QTableWidgetItem(f"{all_Fix_purity_combinations[i][0]:.4f}"))
            self.color_sim_table.setItem(i + 1, 13, QTableWidgetItem(f"{all_Fix_purity_combinations[i][1]:.4f}"))
            self.color_sim_table.setItem(i + 1, 18, QTableWidgetItem(f"{all_Fix_purity_combinations[i][2]:.4f}"))

            # NTSC list
            #print("self.RGB_Fix_NTSC_List",self.RGB_Fix_NTSC_List)
            self.color_sim_table.setItem(i + 1, 19, QTableWidgetItem(f"{self.RGB_Fix_NTSC_List[i]:.4f}"))

            # RGB_Fix_CF_material
            all_Fix_material_combinations = list(
                itertools.product(self.RCF_Fix_Sim_headers_use, self.GCF_Fix_Sim_headers_use, self.BCF_Fix_Sim_headers_use))
            self.color_sim_table.setItem(i + 1, 22, QTableWidgetItem(f"{all_Fix_material_combinations[i][0]}"))
            self.color_sim_table.setItem(i + 1, 24, QTableWidgetItem(f"{all_Fix_material_combinations[i][1]}"))
            self.color_sim_table.setItem(i + 1, 26, QTableWidgetItem(f"{all_Fix_material_combinations[i][2]}"))

            # RGB_Fix_CF_TK
            all_Fix_TK_combinations = list(
                itertools.product(self.R_FIX_TK_sim_list, self.G_FIX_TK_sim_list, self.B_FIX_TK_sim_list))

            self.color_sim_table.setItem(i + 1, 23, QTableWidgetItem(f"{all_Fix_TK_combinations[i][0]:}"))
            self.color_sim_table.setItem(i + 1, 25, QTableWidgetItem(f"{all_Fix_TK_combinations[i][1]:}"))
            self.color_sim_table.setItem(i + 1, 27, QTableWidgetItem(f"{all_Fix_TK_combinations[i][2]:}"))


            # BLU x,y
            self.color_sim_table.setItem(i + 1, 20, QTableWidgetItem(f"{self.BLU_x}"))
            self.color_sim_table.setItem(i + 1, 21, QTableWidgetItem(f"{self.BLU_y}"))

            # Other layer
            self.color_sim_table.setItem(i + 1, 28, QTableWidgetItem(f"{self.CS_light_source.currentText()}"))
            if self.CS_layer1_mode.currentText() == "layer1_自訂":
                self.color_sim_table.setItem(i + 1, 29, QTableWidgetItem(f"{self.CS_layer1_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 29, QTableWidgetItem(""))
            if self.CS_layer2_mode.currentText() == "layer2_自訂":
                self.color_sim_table.setItem(i + 1, 30, QTableWidgetItem(f"{self.CS_layer2_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 30, QTableWidgetItem(""))
            if self.CS_layer3_mode.currentText() == "layer3_自訂":
                self.color_sim_table.setItem(i + 1, 31, QTableWidgetItem(f"{self.CS_layer3_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 31, QTableWidgetItem(""))
            if self.CS_layer4_mode.currentText() == "layer4_自訂":
                self.color_sim_table.setItem(i + 1, 32, QTableWidgetItem(f"{self.CS_layer4_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 32, QTableWidgetItem(""))
            if self.CS_layer5_mode.currentText() == "layer5_自訂":
                self.color_sim_table.setItem(i + 1, 33, QTableWidgetItem(f"{self.CS_layer5_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 33, QTableWidgetItem(""))
            if self.CS_layer6_mode.currentText() == "layer6_自訂":
                self.color_sim_table.setItem(i + 1, 34, QTableWidgetItem(f"{self.CS_layer6_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 34, QTableWidgetItem(""))
            # 開始篩選
            should_remove_row = False  # 初始化是否移除當前行的標記
            # 作為最後計算移除row的判斷
            self.count = 0
            # print("i",i)
            # W_x
            if self.W_checkbox.isChecked():
                # 檢查是否為空字符串
                # 檢查條件並決定是否標記當前行為移除
                if self.Wx_edit.text() and self.W_tolerance_edit.text():
                    # 現在只有當兩個條件都不為空，且值不在指定範圍內時，才標記當前行為移除
                    if not (float(self.W_x_Fix_sum_list[i]) >= (
                            float(self.Wx_edit.text()) - float(self.W_tolerance_edit.text())) and
                            float(self.W_x_Fix_sum_list[i]) <= (
                                    float(self.Wx_edit.text()) + float(self.W_tolerance_edit.text()))):
                        print("W_x_out_of_range")
                        should_remove_row = True

            # W_y
            if self.W_checkbox.isChecked():
                # 檢查是否為空字符串
                if self.Wy_edit.text() and self.W_tolerance_edit.text():
                    # 現在只有當兩個條件都不為空，且值不在指定範圍內時，才標記當前行為移除
                    if not (float(self.W_y_Fix_sum_list[i]) >= (
                            float(self.Wy_edit.text()) - float(self.W_tolerance_edit.text())) and
                            float(self.W_y_Fix_sum_list[i]) <= (
                                    float(self.Wy_edit.text()) + float(self.W_tolerance_edit.text()))):
                        print("W_x_out_of_range")
                        should_remove_row = True

            # R_x
            if self.R_checkbox.isChecked():
                # 檢查是否為空字符串
                print("R_checkbox_check")
                # 檢查是否為空字符串
                # 檢查是否為空字符串
                if self.Rx_edit.text() and self.R_tolerance_edit.text():
                    if not (float(all_Fix_x_combinations[i][0]) >= (
                            float(self.Rx_edit.text()) - float(self.R_tolerance_edit.text()))
                            and float(all_Fix_x_combinations[i][0]) <= (
                                    float(self.Rx_edit.text()) + float(self.R_tolerance_edit.text()))):
                        should_remove_row = True

            # R_y
            if self.R_checkbox.isChecked():
                # 檢查是否為空字符串
                if self.Ry_edit.text() and self.R_tolerance_edit.text():
                    if not (float(all_Fix_y_combinations[i][0]) >= (
                            float(self.Ry_edit.text()) - float(self.R_tolerance_edit.text()))
                            and float(all_Fix_y_combinations[i][0]) <= (
                                    float(self.Ry_edit.text()) + float(self.R_tolerance_edit.text()))):
                        should_remove_row = True

            # G_x
            if self.G_checkbox.isChecked():
                # 檢查是否為空字符串
                if self.Gx_edit.text() and self.G_tolerance_edit.text():
                    if not (float(all_Fix_x_combinations[i][1]) >= (
                            float(self.Gx_edit.text()) - float(self.G_tolerance_edit.text()))
                            and float(all_Fix_x_combinations[i][1]) <= (
                                    float(self.Gx_edit.text()) + float(self.G_tolerance_edit.text()))):
                        should_remove_row = True

            # G_y
            if self.G_checkbox.isChecked():
                # 檢查是否為空字符串
                if self.Gy_edit.text() and self.G_tolerance_edit.text():
                    if not (float(all_Fix_y_combinations[i][1]) >= (
                            float(self.Gy_edit.text()) - float(self.G_tolerance_edit.text()))
                            and float(all_Fix_y_combinations[i][1]) <= (
                                    float(self.Gy_edit.text()) + float(self.G_tolerance_edit.text()))):
                        should_remove_row = True

            # B_x
            if self.B_checkbox.isChecked():
                # 檢查是否為空字符串
                if self.Bx_edit.text() and self.B_tolerance_edit.text():
                    if not (float(all_Fix_x_combinations[i][2]) >= (
                            float(self.Bx_edit.text()) - float(self.B_tolerance_edit.text()))
                            and float(all_Fix_x_combinations[i][2]) <= (
                                    float(self.Bx_edit.text()) + float(self.B_tolerance_edit.text()))):
                        should_remove_row = True

            # B_y
            if self.B_checkbox.isChecked():
                # 檢查是否為空字符串
                if self.By_edit.text() and self.B_tolerance_edit.text():
                    if not (float(all_Fix_y_combinations[i][2]) >= (
                            float(self.By_edit.text()) - float(self.B_tolerance_edit.text()))
                            and float(all_Fix_y_combinations[i][2]) <= (
                                    float(self.By_edit.text()) + float(self.B_tolerance_edit.text()))):
                        should_remove_row = True
            # R_Wave & purity
            if self.R_wave_checkbox.isChecked():
                # 檢查是否為空字符串
                print("R_wave_checkbox_check")
                # 檢查是否為空字符串
                if not (self.R_wave_start_edit.text() and self.R_wave_end_edit.text() and
                        float(all_Fix_wave_combinations[i][0]) >= (float(self.R_wave_start_edit.text())) and float(
                            all_Fix_wave_combinations[i][0]) <= (float(self.R_wave_end_edit.text()))):
                    should_remove_row = True

            if self.R_purity_checkbox.isChecked():
                # 檢查是否為空字符串
                print("R_purity_checkbox_check")
                # 檢查是否為空字符串
                if not (self.R_purity_limit_edit.text() and
                        float(all_Fix_purity_combinations[i][0]) >= (float(self.R_purity_limit_edit.text()))):
                    should_remove_row = True

            # G_Wave & purity
            if self.G_wave_checkbox.isChecked():
                # 檢查是否為空字符串
                print("G_wave_checkbox_check")
                # 檢查是否為空字符串
                if not (self.G_wave_start_edit.text() and self.G_wave_end_edit.text() and
                        float(all_Fix_wave_combinations[i][1]) >= (float(self.G_wave_start_edit.text())) and float(
                            all_Fix_wave_combinations[i][1]) <= (float(self.G_wave_end_edit.text()))):
                    should_remove_row = True

            if self.G_purity_checkbox.isChecked():
                # 檢查是否為空字符串
                print("G_purity_checkbox_check")
                # 檢查是否為空字符串
                if not (self.G_purity_limit_edit.text() and
                        float(all_Fix_purity_combinations[i][1]) >= (float(self.G_purity_limit_edit.text()))):
                    should_remove_row = True

            # B_Wave & purity
            if self.B_wave_checkbox.isChecked():
                # 檢查是否為空字符串
                print("B_wave_checkbox_check")
                # 檢查是否為空字符串
                if not (self.B_wave_start_edit.text() and self.B_wave_end_edit.text() and
                        float(all_Fix_wave_combinations[i][2]) >= (float(self.B_wave_start_edit.text())) and float(
                            all_Fix_wave_combinations[i][2]) <= (float(self.B_wave_end_edit.text()))):
                    should_remove_row = True

            if self.B_purity_checkbox.isChecked():
                # 檢查是否為空字符串
                print("B_purity_checkbox_check")
                # 檢查是否為空字符串
                if not (self.B_purity_limit_edit.text() and
                        float(all_Fix_purity_combinations[i][2]) >= (float(self.B_purity_limit_edit.text()))):
                    should_remove_row = True

            # NTSC
            if self.NTSC_check.isChecked():
                print("NTSC_check.isChecked")
                # 檢查是否為空字符串
                if not (self.NTSC_check_edit.text() and
                        float(self.RGB_Fix_NTSC_List[i]) >= (float(self.NTSC_check_edit.text()))):
                    should_remove_row = True

            # TK
            # R=G=B
            if self.RGB_1_checkbox.isChecked():
                print("RGB_1_checkbox.isChecked")
                if not (float(all_Fix_TK_combinations[i][0]) == (float(all_Fix_TK_combinations[i][1])) and float(
                        all_Fix_TK_combinations[i][1]) == (float(all_Fix_TK_combinations[i][2]))):
                    should_remove_row = True

            # R=G>B
            if self.RGB_2_checkbox.isChecked():
                print("RGB_2_checkbox.isChecked")
                if not (float(all_Fix_TK_combinations[i][0]) == (float(all_Fix_TK_combinations[i][1])) and float(
                        all_Fix_TK_combinations[i][1]) > (float(all_Fix_TK_combinations[i][2]))):
                    should_remove_row = True

            # R=G>B
            if self.RGB_3_checkbox.isChecked():
                print("RGB_3_checkbox.isChecked")
                if not (float(all_Fix_TK_combinations[i][0]) > (float(all_Fix_TK_combinations[i][1])) and float(
                        all_Fix_TK_combinations[i][1]) == (float(all_Fix_TK_combinations[i][2]))):
                    should_remove_row = True

            # R>G>B
            if self.RGB_4_checkbox.isChecked():
                print("RGB_4_checkbox.isChecked")
                if not (float(all_Fix_TK_combinations[i][0]) > (float(all_Fix_TK_combinations[i][1])) and float(
                        all_Fix_TK_combinations[i][1]) > (float(all_Fix_TK_combinations[i][2]))):
                    should_remove_row = True

            # R<G=B
            if self.RGB_5_checkbox.isChecked():
                print("RGB_5_checkbox.isChecked")
                if not (float(all_Fix_TK_combinations[i][0]) < (float(all_Fix_TK_combinations[i][1])) and float(
                        all_Fix_TK_combinations[i][1]) == (float(all_Fix_TK_combinations[i][2]))):
                    should_remove_row = True

            # R<G<B
            if self.RGB_6_checkbox.isChecked():
                print("RGB_6_checkbox.isChecked")
                if not (float(all_Fix_TK_combinations[i][0]) < (float(all_Fix_TK_combinations[i][1])) and float(
                        all_Fix_TK_combinations[i][1]) < (float(all_Fix_TK_combinations[i][2]))):
                    should_remove_row = True

            # R=G<B
            if self.RGB_6_checkbox.isChecked():
                print("RGB_7_checkbox.isChecked")
                if not (float(all_Fix_TK_combinations[i][0]) == (float(all_Fix_TK_combinations[i][1])) and float(
                        all_Fix_TK_combinations[i][1]) < (float(all_Fix_TK_combinations[i][2]))):
                    should_remove_row = True

            # Free RTK
            if self.R_TK_check.isChecked():
                print("R_TK_check")
                # 檢查是否為空字符串
                if not (self.R_TK_start_edit.text() and self.R_TK_end_edit.text() and
                        float(all_Fix_TK_combinations[i][0]) >= (float(self.R_TK_start_edit.text())) and float(
                            all_Fix_TK_combinations[i][0]) <= (float(self.R_TK_end_edit.text()))):
                    print("i-pass", i)
                    print("self.RGB_TK_list[i][0]", self.RGB_TK_list[i][0])
                    should_remove_row = True

            # Free GTK
            if self.G_TK_check.isChecked():
                # 檢查是否為空字符串
                print("G_TK_check")
                # 檢查是否為空字符串
                if not (self.G_TK_start_edit.text() and self.G_TK_end_edit.text() and
                        float(all_Fix_TK_combinations[i][1]) >= (float(self.G_TK_start_edit.text())) and float(
                            all_Fix_TK_combinations[i][1]) <= (float(self.G_TK_end_edit.text()))):
                    should_remove_row = True

            # Free BTK
            if self.B_TK_check.isChecked():
                print("B_TK_check")
                # 檢查是否為空字符串
                if not (self.B_TK_start_edit.text() and self.B_TK_end_edit.text() and
                        float(all_Fix_TK_combinations[i][2]) >= (float(self.B_TK_start_edit.text())) and float(
                            all_Fix_TK_combinations[i][2]) <= (float(self.B_TK_end_edit.text()))):
                    should_remove_row = True

            # 如果當前行需要被移除，將其索引添加到移除集合中
            if should_remove_row:
                rows_to_remove.add(i + 1)

        # 移除標記為移除的行
        for row_index in sorted(rows_to_remove, reverse=True):
            self.color_sim_table.removeRow(row_index)

        # 重新設置剩餘行的行號
        for i in range(self.color_sim_table.rowCount()):
            self.color_sim_table.setItem(i, 0, QTableWidgetItem(str(i)))

        # 設定默認值
        column1_default_values = ["項目", "Wx", "Wy", "WY", "Rx", "Ry", "RY", "Rλ", "R_Purity", "Gx", "Gy", "GY",
                                  "Gλ", "G_Purity", "Bx", "By", "BY", "B_λ", "B_Purity", "NTSC%", "BLUx", "BLUy",
                                  "R色阻選擇", "R色阻厚度", "G色阻選擇", "G色阻厚度", "B色阻選擇", "B色阻厚度",
                                  "背光選擇", "Layer_1", "Layer_2", "Layer_3", "Layer_4",
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

        # 加載數據後自動調整所有欄位的寬度以適應內容
        self.color_sim_table.resizeColumnsToContents()

    def calculate_RCF_Change(self):
        if self.R_aK_mode.currentText() == "模擬":
            connection_RCF_Change = sqlite3.connect("RCF_Change_spectrum.db")
            cursor_RCF_Change = connection_RCF_Change.cursor()
            # 取得Change資料
            table_name_RCF_Change = self.R_aK_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_RCF_Change = f"SELECT * FROM '{table_name_RCF_Change}';"
            cursor_RCF_Change.execute(query_RCF_Change)
            result_RCF_Change = cursor_RCF_Change.fetchall()
            #print("result_RCF_Change",result_RCF_Change)
            # 找到指定標題的欄位索引
            header_RCF_Change = [column[0] for column in cursor_RCF_Change.description]

            # 移除所有標題中的編號
            header_RCF_Change = [re.sub(r'_\d+$', '', col) for col in header_RCF_Change]
            header_RCF_Change = header_RCF_Change[1:]
            print("header_RCF_Change", header_RCF_Change)

            # 創建一個字典來跟踪每個標題的索引
            header_indices = {}

            for index, header in enumerate(header_RCF_Change, start=1):
                if header not in header_indices:
                    header_indices[header] = []
                header_indices[header].append(index)

            # 創建不重複標題列表
            self.header_RCF_Change_list = list(header_indices.keys())

            # 創建對應的索引列表
            indices_without_number_list = list(header_indices.values())

            # 打印結果（可選）
            print("self.header_RCF_Change_list:", self.header_RCF_Change_list)
            print("indices_without_number_list:", indices_without_number_list)

            # 使用 pd.to_numeric 转换为数值类型，并处理无法转换的值
            SP_series_list = []
            for j in range(len(self.header_RCF_Change_list)):
                SP_series = []
                for i in indices_without_number_list[j]:
                    # 這裡假設 result_RCF_Change 是一個二維列表，其中每個子列表代表一行數據
                    SP_series.append(
                        pd.to_numeric(pd.Series([row[i] for row in result_RCF_Change]), errors='coerce').fillna(0))
                SP_series_list.append(SP_series)
            # 每個色組獨立一個list[[第一組色組list],[第二組色組list],[第三組色組list]]
            #print("SP_series_list", SP_series_list)

            # TK_SP準備分厚度順序
            self.TK_R_SP_list = []
            self.Sort_R_TK_SP_list = []
            for j in range(len(SP_series_list)):
                TK_SP = []
                for i in range(len(indices_without_number_list[j])):
                    TK_SP.append(SP_series_list[j][i].iloc[0])
                self.Sort_R_TK_SP = sorted(TK_SP,reverse=True)
                self.Sort_R_TK_SP_list.append(self.Sort_R_TK_SP)
                self.TK_R_SP_list.append(TK_SP)
            print("self.TK_R_SP_list", self.TK_R_SP_list)
            print("self.Sort_R_TK_SP_list", self.Sort_R_TK_SP_list)

            # 反轉找回原本的順序list
            Re_list = []
            for j in range(len(SP_series_list)):
                Re = []
                for i in range(len(indices_without_number_list[j])):
                    # 尋找 self.Sort_TK_SP_list[j][i] 在 self.TK_SP_list[j] 中的索引
                    index = self.TK_R_SP_list[j].index(self.Sort_R_TK_SP_list[j][i])
                    Re.append(index)
                Re_list.append(Re)

            #print("Re_list", Re_list)

            # 找到Alist
            A_list = []
            for j in range(len(SP_series_list)):
                A = []
                for i in range(len(indices_without_number_list[j]) - 1):
                    A.append((-1 / (float(self.Sort_R_TK_SP_list[j][i + 1]) - float(self.Sort_R_TK_SP_list[j][i])) * np.log(
                    SP_series_list[j][Re_list[j][i + 1]][1:] / SP_series_list[j][Re_list[j][i]][1:])).reset_index(drop=True))
                A_list.append(A)
            #print("A_list",A_list)

            # 找到K-list
            K_list = []
            for j in range(len(SP_series_list)):
                K = []
                for i in range(len(indices_without_number_list[j]) - 1):
                    K.append(SP_series_list[j][Re_list[j][i]][1:].reset_index(drop=True) / np.exp(
                    (-1 * A_list[j][i]) * self.Sort_R_TK_SP_list[j][i]))
                K_list.append(K)
            #print("K_list", K_list)

            # 厚度分割list-*
            self.RCF_TK_list = []
            try:
                if float(self.R_aK_TK_Start.text()) < float(self.R_aK_TK_End.text()):
                    i = 0
                    while float(self.R_aK_TK_Start.text()) + float(self.R_aK_TK_interval.text()) * i < float(self.R_aK_TK_End.text()):
                        print("self.R_aK_TK_interval.text()",self.R_aK_TK_interval.text())
                        self.RCF_TK_list.append(float(self.R_aK_TK_Start.text()) + float(self.R_aK_TK_interval.text()) * i)
                        i += 1
                        print("self.RCF_TK_list", self.RCF_TK_list)
                    self.RCF_TK_list.append(float(self.R_aK_TK_End.text()))

                elif float(self.R_aK_TK_End.text()) < float(self.R_aK_TK_Start.text()):
                    i = 0
                    while float(self.R_aK_TK_End.text()) + float(self.R_aK_TK_interval.text()) * i < float(self.R_aK_TK_Start.text()):
                        self.RCF_TK_list.append(float(self.R_aK_TK_End.text()) + float(self.R_aK_TK_interval.text()) * i)
                        i += 1
                    self.RCF_TK_list.append(float(self.R_aK_TK_Start.text()))

            except ValueError:
                print("R輸入的數值格式為空，請檢查後重新輸入。")
            print("self.RCF_TK_list", self.RCF_TK_list)

            # 找出AK_RCF_Change_spectrum_list
            self.AK_RCF_Change_spectrum_Series_list = []
            for j in range(len(SP_series_list)):
                AK_RCF_Change_spectrum_Series_2 = []
                for RCF_TK in self.RCF_TK_list:
                    # 先将 R_aK_TK 添加到 self.Sort_TK_SP_list
                    self.Sort_R_TK_SP_list[j].append(RCF_TK)
                    # 然后对更新后的列表进行排序
                    self.Sort_TK_SP_AK_list_2 = sorted(self.Sort_R_TK_SP_list[j], reverse=True)
                    print("self.Sort_TK_SP_AK_list_2",self.Sort_TK_SP_AK_list_2)
                    print("RCF_TK",RCF_TK)
                    # 4種情況
                    if RCF_TK == max(self.Sort_TK_SP_AK_list_2):
                        print("max")
                        AK_RCF_Change_spectrum_Series = K_list[j][0] * np.exp(-1 * A_list[j][0] * RCF_TK)
                        AK_RCF_Change_spectrum_Series_2.append(AK_RCF_Change_spectrum_Series)
                        # print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                        # return AK_RCF_Change_spectrum_Series
                    elif RCF_TK == min(self.Sort_TK_SP_AK_list_2):
                        print("min")
                        AK_RCF_Change_spectrum_Series = K_list[j][len(K_list[j]) - 1] * np.exp(
                            -1 * A_list[j][len(A_list[j]) - 1] * RCF_TK)
                        AK_RCF_Change_spectrum_Series_2.append(AK_RCF_Change_spectrum_Series)

                    elif RCF_TK in self.TK_R_SP_list[j]:
                        print("equal")
                        AK_RCF_Change_spectrum_Series = SP_series_list[j][self.TK_R_SP_list[j].index(RCF_TK)][1:].reset_index(
                            drop=True)
                        AK_RCF_Change_spectrum_Series_2.append(AK_RCF_Change_spectrum_Series)
                        # print("AK_RCF_Change_spectrum_Series",AK_RCF_Change_spectrum_Series)
                        # return AK_RCF_Change_spectrum_Series
                    else:
                        print("mid")
                        position = self.Sort_TK_SP_AK_list_2.index(RCF_TK)
                        print("position",position)
                        #print("K_list[j]",K_list[j])
                        AK_RCF_Change_spectrum_Series = K_list[j][position - 1] * np.exp(
                            -1 * A_list[j][position - 1] * RCF_TK)
                        AK_RCF_Change_spectrum_Series_2.append(AK_RCF_Change_spectrum_Series)
                    self.Sort_R_TK_SP_list[j].pop()
                self.AK_RCF_Change_spectrum_Series_list.append(AK_RCF_Change_spectrum_Series_2)
            self.CIEparameter()
            # BLU +Cell part
            self.cell_blu_total_spectrum = self.calculate_BLU() * self.calculate_layer1() * self.calculate_layer2() \
                                           * self.calculate_layer3() * self.calculate_layer4() * self.calculate_layer5() \
                                           * self.calculate_layer6()
            self.cell_other_spectrum =  self.calculate_layer1() * self.calculate_layer2() \
                                           * self.calculate_layer3() * self.calculate_layer4() * self.calculate_layer5() \
                                           * self.calculate_layer6()
            # 這個迴圈將遍歷 self.AK_RCF_Change_spectrum_Series_list 的每個子列表
            for i in range(len(self.AK_RCF_Change_spectrum_Series_list)):
                # 這個迴圈將遍歷當前子列表中的每個 Series 對象
                for j in range(len(self.AK_RCF_Change_spectrum_Series_list[i])):
                    # 將每個 Series 元素與 self.cell_blu_total_spectrum 相乘
                    self.AK_RCF_Change_spectrum_Series_list[i][j] = self.AK_RCF_Change_spectrum_Series_list[i][
                                                                        j] * self.cell_blu_total_spectrum

            # 此時 self.AK_RCF_Change_spectrum_Series_list 應已更新為乘以 self.cell_blu_total_spectrum 後的值
            self.R_Change_X_list = []

            for group in self.AK_RCF_Change_spectrum_Series_list:
                group_R_Change_X = []
                for series in group:
                    # 將每個 Series 元素與 self.CIE_spectrum_Series_X 相乘
                    multiplied_series = series.multiply(self.CIE_spectrum_Series_X)
                    group_R_Change_X.append(multiplied_series)
                self.R_Change_X_list.append(group_R_Change_X)
            # print("self.R_Change_X_list",self.R_Change_X_list)

            self.R_Change_Y_list = []

            for group in self.AK_RCF_Change_spectrum_Series_list:
                group_R_Change_Y = []
                for series in group:
                    # 將每個 Series 元素與 self.CIE_spectrum_Series_Y 相乘
                    multiplied_series = series.multiply(self.CIE_spectrum_Series_Y)
                    group_R_Change_Y.append(multiplied_series)
                self.R_Change_Y_list.append(group_R_Change_Y)
            print("self.R_Change_Y_list",self.R_Change_Y_list)



            self.R_Change_Z_list = []

            for group in self.AK_RCF_Change_spectrum_Series_list:
                group_R_Change_Z = []
                for series in group:
                    # 將每個 Series 元素與 self.CIE_spectrum_Series_Z 相乘
                    multiplied_series = series.multiply(self.CIE_spectrum_Series_Z)
                    group_R_Change_Z.append(multiplied_series)
                self.R_Change_Z_list.append(group_R_Change_Z)
            # X_Sum
            self.R_X_Change_sim_sum_list = []

            for group in self.R_Change_X_list:
                group_sums = []
                for series in group:
                    # 計算每個 Series 的總和
                    sum_series = series.sum()
                    group_sums.append(sum_series)
                self.R_X_Change_sim_sum_list.append(group_sums)
            print("self.R_X_Change_sim_sum_list",self.R_X_Change_sim_sum_list)
            # Y_Sum
            self.R_Y_Change_sim_sum_list = []

            for group in self.R_Change_Y_list:
                group_sums = []
                for series in group:
                    # 計算每個 Series 的總和
                    sum_series = series.sum()
                    group_sums.append(sum_series)
                self.R_Y_Change_sim_sum_list.append(group_sums)
            print("self.R_Y_Change_sim_sum_list",self.R_Y_Change_sim_sum_list)

            # Z_Sum
            self.R_Z_Change_sim_sum_list = []

            for group in self.R_Change_Z_list:
                group_sums = []
                for series in group:
                    # 計算每個 Series 的總和
                    sum_series = series.sum()
                    group_sums.append(sum_series)
                self.R_Z_Change_sim_sum_list.append(group_sums)
            print("self.R_Z_Change_sim_sum_list",self.R_Z_Change_sim_sum_list)

            # R_x
            self.R_Change_x_sim = []
            # 確保所有的列結構相同
            for group_index in range(len(self.R_X_Change_sim_sum_list)):
                group_sim = []
                for item_index in range(len(self.R_X_Change_sim_sum_list[group_index])):
                    X = self.R_X_Change_sim_sum_list[group_index][item_index]
                    Y = self.R_Y_Change_sim_sum_list[group_index][item_index]
                    Z = self.R_Z_Change_sim_sum_list[group_index][item_index]

                    # 计算比例，注意要避免除以零
                    total = X + Y + Z
                    if total > 0:
                        sim = X / total
                    else:
                        sim = 0
                    group_sim.append(sim)
                self.R_Change_x_sim.append(group_sim)

            # R_y
            self.R_Change_y_sim = []
            # 確保所有的列結構相同
            for group_index in range(len(self.R_Y_Change_sim_sum_list)):
                group_sim = []
                for item_index in range(len(self.R_Y_Change_sim_sum_list[group_index])):
                    X = self.R_X_Change_sim_sum_list[group_index][item_index]
                    Y = self.R_Y_Change_sim_sum_list[group_index][item_index]
                    Z = self.R_Z_Change_sim_sum_list[group_index][item_index]

                    # 计算比例，注意要避免除以零
                    total = X + Y + Z
                    if total > 0:
                        sim = Y / total
                    else:
                        sim = 0
                    group_sim.append(sim)
                self.R_Change_y_sim.append(group_sim)


            # # R_T
            # self.R_Change_T_sim = []
            # # 確保所有的列結構相同
            # for group_index in range(len(self.R_Z_Change_sim_sum_list)):
            #     group_sim = []
            #     for item_index in range(len(self.R_Z_Change_sim_sum_list[group_index])):
            #         X = self.R_X_Change_sim_sum_list[group_index][item_index]
            #         Y = self.R_Y_Change_sim_sum_list[group_index][item_index]
            #         Z = self.R_Z_Change_sim_sum_list[group_index][item_index]
            #
            #         # 计算比例，注意要避免除以零
            #         total = X + Y + Z
            #         if total > 0:
            #             sim = 3 * Y
            #         else:
            #             sim = 0
            #         group_sim.append(sim)
            #     self.R_Change_T_sim.append(group_sim)

            # 關閉連線
            connection_RCF_Change.close()
            print("self.R_Change_x_sim",self.R_Change_x_sim)
            print("self.R_Change_y_sim",self.R_Change_y_sim)
            # print("self.R_Change_T_sim-1",self.R_Change_T_sim)
            self.R_Change_wave_sim = []
            self.R_Change_purity_sim = []

            # xy_n 是觀察者使用的光源（例如 D65）
            xy_n = self.observer_D65

            for group_index in range(len(self.R_Change_x_sim)):
                group_wave = []
                group_purity = []

                for item_index in range(len(self.R_Change_x_sim[group_index])):
                    xy = [self.R_Change_x_sim[group_index][item_index], self.R_Change_y_sim[group_index][item_index]]

                    # 計算主波長和CIE坐標
                    dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)

                    # 計算距離
                    distance_from_white = math.sqrt((xy[0] - xy_n[0]) ** 2 + (xy[1] - xy_n[1]) ** 2)
                    distance_from_black = math.sqrt(
                        (CIEcoordinate[0] - xy_n[0]) ** 2 + (CIEcoordinate[1] - xy_n[1]) ** 2)

                    # 計算色純度
                    purity = distance_from_white / distance_from_black * 100 if distance_from_black != 0 else 0

                    group_wave.append(dominant_wavelength)
                    group_purity.append(purity)

                self.R_Change_wave_sim.append(group_wave)
                self.R_Change_purity_sim.append(group_purity)
            print("self.R_Change_wave_sim",self.R_Change_wave_sim)
            print("self.R_Change_wave_sim",self.R_Change_wave_sim[0][0])
            print("self.R_Change_purity_sim",self.R_Change_purity_sim)
            # 研究RT怎寫
            # 這個迴圈將遍歷 self.AK_RCF_Change_spectrum_Series_list 的每個子列表
            for i in range(len(self.AK_RCF_Change_spectrum_Series_list)):
                # 這個迴圈將遍歷當前子列表中的每個 Series 對象
                for j in range(len(self.AK_RCF_Change_spectrum_Series_list[i])):
                    # 將每個 Series 元素與 self.cell_blu_total_spectrum 相乘
                    self.AK_RCF_Change_spectrum_Series_list[i][j] = \
                        self.AK_RCF_Change_spectrum_Series_list[i][
                            j] / self.calculate_BLU() * self.CIE_spectrum_Series_Y

            self.BLUY = self.calculate_BLU() * self.CIE_spectrum_Series_Y
            # print("改過的self.AK_RCF_Change_spectrum_Series_list",
            #       self.AK_RCF_Change_spectrum_Series_list)
            # 先計算 self.BLUY 的總和
            bluy_sum = self.BLUY.sum()

            # 初始化 self.R_Change_T_sim 為空列表
            self.R_Change_T_sim = []

            # 遍歷 self.AK_RCF_Change_spectrum_Series_list 中的每一組 Series 列表
            for group in self.AK_RCF_Change_spectrum_Series_list:
                # 初始化當前組的結果列表
                group_results = []
                # 遍歷當前組中的每一個 Series 對象
                for series in group:
                    # 計算當前 Series 對象的總和乘以 self.BLUY 的總和，並將結果加到當前組的結果列表上
                    group_results.append(series.sum() / bluy_sum * 100)
                # 將當前組的結果列表加到 self.R_Change_T_sim 上
                self.R_Change_T_sim.append(group_results)
            print("self.R_Change_T_sim-2",self.R_Change_T_sim)

        elif self.R_aK_mode.currentText() == "自訂":
            print("R_aK_mode自訂")
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
            self.header_RCF_Change_list_2 = [column[0] for column in cursor_RCF_Change.description]

            # 移除所有標題中的編號
            self.header_RCF_Change_list_2 = [re.sub(r'_\d+$', '', col) for col in self.header_RCF_Change_list_2]
            # print("self.header_RCF_Change_list_2", self.header_RCF_Change_list_2)
            # 自訂標題list
            # 找到指定標題的欄位索引
            header_RCF_Change = [column[0] for column in cursor_RCF_Change.description]

            # 移除所有標題中的編號
            header_RCF_Change = [re.sub(r'_\d+$', '', col) for col in header_RCF_Change]
            header_RCF_Change = header_RCF_Change[1:]
            print("header_RCF_Change", header_RCF_Change)

            # 創建一個字典來跟踪每個標題的索引
            header_indices = {}

            for index, header in enumerate(header_RCF_Change, start=1):
                if header not in header_indices:
                    header_indices[header] = []
                header_indices[header].append(index)

            # 創建不重複標題列表
            self.header_RCF_Change_list = list(header_indices.keys())
            print("self.header_RCF_Change_list", self.header_RCF_Change_list)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(self.header_RCF_Change_list_2) if col == column_name_RCF_Change]
            print("indices_without_number", indices_without_number)

            # 使用 pd.to_numeric 转换为数值类型，并处理无法转换的值
            SP_series_list = [
                pd.to_numeric(pd.Series([row[i] for row in result_RCF_Change]), errors='coerce').fillna(0)
                for i in indices_without_number
            ]
            #print("SP_series_list", SP_series_list)

            # TK_SP準備分厚度順序
            TK_SP_list = []
            for i in range(len(indices_without_number)):
                TK_SP_list.append(SP_series_list[i].iloc[0])
            print("TK_SP_list", TK_SP_list)
            Sort_TK_SP_list = sorted(TK_SP_list, reverse=True)
            print("Sort_TK_SP_list", Sort_TK_SP_list)
            self.R_aK_TK_edit_label.setText(f"TKrange: {min(TK_SP_list):.2f}~{max(TK_SP_list):.2f}")

            # 反轉找回原本的順序list
            Re_list = []
            for i in range(len(indices_without_number)):
                Re_list.append(TK_SP_list.index(Sort_TK_SP_list[i]))
            print("Re_list", Re_list)

            A_list = []
            for i in range(len(indices_without_number) - 1):
                A_list.append((-1 / (float(Sort_TK_SP_list[i + 1]) - float(Sort_TK_SP_list[i])) * np.log(
                    SP_series_list[Re_list[i + 1]][1:] / SP_series_list[Re_list[i]][1:])).reset_index(drop=True))
            #print("A_list", A_list)
            K_list = []
            for i in range(len(indices_without_number) - 1):
                K_list.append(SP_series_list[Re_list[i]][1:].reset_index(drop=True) / np.exp(
                    (-1 * A_list[i]) * Sort_TK_SP_list[i]))
            #print("K_list", K_list)
            if self.R_aK_TK_edit.text() == "":
                AK_RCF_Change_spectrum_Series = 1
                print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                return AK_RCF_Change_spectrum_Series
            elif self.R_aK_TK_edit.text() != "":

                R_aK_TK = float(self.R_aK_TK_edit.text())
                # 厚度分割list-*
                self.RCF_TK_list = [R_aK_TK]
                print("self.RCF_TK_list",self.RCF_TK_list)
                # 先将 R_aK_TK 添加到 Sort_TK_SP_list
                Sort_TK_SP_list.append(R_aK_TK)
                # 然后对更新后的列表进行排序
                Sort_TK_SP_AK_list = sorted(Sort_TK_SP_list, reverse=True)
                print("Sort_TK_SP_AK_list", Sort_TK_SP_AK_list)

                # 4種情況
                if R_aK_TK == max(Sort_TK_SP_AK_list):
                    print("max")
                    AK_RCF_Change_spectrum_Series = K_list[0] * np.exp(-1 * A_list[0] * R_aK_TK)
                    # print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # return AK_RCF_Change_spectrum_Series
                elif R_aK_TK == min(Sort_TK_SP_AK_list):
                    print("min")
                    AK_RCF_Change_spectrum_Series = K_list[len(K_list) - 1] * np.exp(
                        -1 * A_list[len(A_list) - 1] * R_aK_TK)
                    #print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # 在Color_Table上顯示

                elif R_aK_TK in TK_SP_list:
                    print("equal")
                    AK_RCF_Change_spectrum_Series = SP_series_list[TK_SP_list.index(R_aK_TK)][1:].reset_index(drop=True)
                    print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # 在Color_Table上顯示

                else:
                    print("mid")
                    position = Sort_TK_SP_AK_list.index(R_aK_TK)
                    AK_RCF_Change_spectrum_Series = K_list[position - 1] * np.exp(-1 * A_list[position - 1] * R_aK_TK)
                    #print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                self.CIEparameter()
                # BLU +Cell part
                self.cell_blu_total_spectrum = self.calculate_BLU() * self.calculate_layer1() * self.calculate_layer2() \
                                               * self.calculate_layer3() * self.calculate_layer4() * self.calculate_layer5() \
                                               * self.calculate_layer6()
                AK_RCF_Change_spectrum_Series = AK_RCF_Change_spectrum_Series * self.cell_blu_total_spectrum
                R_Change_X_sum = (AK_RCF_Change_spectrum_Series * self.CIE_spectrum_Series_X).sum()
                self.R_X_Change_sim_sum_list = [[R_Change_X_sum]]
                print("self.R_X_Change_sim_sum_list",self.R_X_Change_sim_sum_list)
                R_Change_Y_sum = (AK_RCF_Change_spectrum_Series * self.CIE_spectrum_Series_Y).sum()
                self.R_Y_Change_sim_sum_list = [[R_Change_Y_sum]]
                print("self.R_Y_Change_sim_sum_list",self.R_Y_Change_sim_sum_list)
                R_Change_Z_sum = (AK_RCF_Change_spectrum_Series * self.CIE_spectrum_Series_Z).sum()
                self.R_Z_Change_sim_sum_list = [[R_Change_Z_sum]]
                print("self.R_Z_Change_sim_sum_list",self.R_Z_Change_sim_sum_list)
                R_Change_Total = R_Change_X_sum + R_Change_Y_sum +R_Change_Z_sum
                # R_x
                self.R_Change_x_sim = [[R_Change_X_sum/(R_Change_Total)]]
                print("self.R_Change_x_sim",self.R_Change_x_sim)
                # R_y
                self.R_Change_y_sim = [[R_Change_Y_sum / (R_Change_Total)]]
                print("self.R_Change_y_sim",self.R_Change_y_sim)
                # R_T
                #print("AK_RCF_Change_spectrum_Series",AK_RCF_Change_spectrum_Series)
                # print("AK_RCF_Change_spectrum_Series-2",AK_RCF_Change_spectrum_Series/self.calculate_BLU())

                R_Change_T_sum = (AK_RCF_Change_spectrum_Series /self.calculate_BLU() * self.CIE_spectrum_Series_Y).sum() / ((self.calculate_BLU()* self.CIE_spectrum_Series_Y).sum()) * 100
                self.R_Change_T_sim = [[R_Change_T_sum]]
                print("self.R_Change_T_sim",self.R_Change_T_sim)
                self.R_Change_wave_sim = []
                self.R_Change_purity_sim = []

                # xy_n 是觀察者使用的光源（例如 D65）
                xy_n = self.observer_D65

                for group_index in range(len(self.R_Change_x_sim)):
                    group_wave = []
                    group_purity = []

                    for item_index in range(len(self.R_Change_x_sim[group_index])):
                        xy = [self.R_Change_x_sim[group_index][item_index],
                              self.R_Change_y_sim[group_index][item_index]]

                        # 計算主波長和CIE坐標
                        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)

                        # 計算距離
                        distance_from_white = math.sqrt((xy[0] - xy_n[0]) ** 2 + (xy[1] - xy_n[1]) ** 2)
                        distance_from_black = math.sqrt(
                            (CIEcoordinate[0] - xy_n[0]) ** 2 + (CIEcoordinate[1] - xy_n[1]) ** 2)

                        # 計算色純度
                        purity = distance_from_white / distance_from_black * 100 if distance_from_black != 0 else 0

                        group_wave.append(dominant_wavelength)
                        group_purity.append(purity)

                    self.R_Change_wave_sim.append(group_wave)
                    self.R_Change_purity_sim.append(group_purity)
                print("self.R_Change_wave_sim", self.R_Change_wave_sim)
                print("self.R_Change_wave_sim", self.R_Change_wave_sim[0][0])
                print("self.R_Change_purity_sim", self.R_Change_purity_sim)


            # 關閉連線
            connection_RCF_Change.close()

    def calculate_GCF_Change(self):
        if self.G_aK_mode.currentText() == "模擬":
            connection_GCF_Change = sqlite3.connect("GCF_Change_spectrum.db")
            cursor_GCF_Change = connection_GCF_Change.cursor()
            # 取得Change資料
            table_name_GCF_Change = self.G_aK_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_GCF_Change = f"SELECT * FROM '{table_name_GCF_Change}';"
            cursor_GCF_Change.execute(query_GCF_Change)
            result_GCF_Change = cursor_GCF_Change.fetchall()
            #print("result_GCF_Change",result_GCF_Change)
            # 找到指定標題的欄位索引
            header_GCF_Change = [column[0] for column in cursor_GCF_Change.description]

            # 移除所有標題中的編號
            header_GCF_Change = [re.sub(r'_\d+$', '', col) for col in header_GCF_Change]
            header_GCF_Change = header_GCF_Change[1:]
            print("header_GCF_Change", header_GCF_Change)

            # 創建一個字典來跟踪每個標題的索引
            header_indices = {}

            for index, header in enumerate(header_GCF_Change, start=1):
                if header not in header_indices:
                    header_indices[header] = []
                header_indices[header].append(index)

            # 創建不重複標題列表
            self.header_GCF_Change_list = list(header_indices.keys())

            # 創建對應的索引列表
            indices_without_number_list = list(header_indices.values())

            # 打印結果（可選）
            print("self.header_GCF_Change_list:", self.header_GCF_Change_list)
            print("indices_without_number_list:", indices_without_number_list)

            # 使用 pd.to_numeric 转换为数值类型，并处理无法转换的值
            SP_series_list = []
            for j in range(len(self.header_GCF_Change_list)):
                SP_series = []
                for i in indices_without_number_list[j]:
                    # 這裡假設 result_RCF_Change 是一個二維列表，其中每個子列表代表一行數據
                    SP_series.append(
                        pd.to_numeric(pd.Series([row[i] for row in result_GCF_Change]), errors='coerce').fillna(0))
                SP_series_list.append(SP_series)
            # 每個色組獨立一個list[[第一組色組list],[第二組色組list],[第三組色組list]]
            #print("SP_series_list", SP_series_list)

            # TK_SP準備分厚度順序
            self.TK_G_SP_list = []
            self.Sort_G_TK_SP_list = []
            for j in range(len(SP_series_list)):
                TK_SP = []
                for i in range(len(indices_without_number_list[j])):
                    TK_SP.append(SP_series_list[j][i].iloc[0])
                self.Sort_G_TK_SP = sorted(TK_SP,reverse=True)
                self.Sort_G_TK_SP_list.append(self.Sort_G_TK_SP)
                self.TK_G_SP_list.append(TK_SP)
            print(" self.TK_G_SP_list", self.TK_G_SP_list)
            print(" self.Sort_G_TK_SP_list", self.Sort_G_TK_SP_list)

            # 反轉找回原本的順序list
            Re_list = []
            for j in range(len(SP_series_list)):
                Re = []
                for i in range(len(indices_without_number_list[j])):
                    # 尋找 self.Sort_TK_SP_list[j][i] 在 self.TK_SP_list[j] 中的索引
                    index = self.TK_G_SP_list[j].index(self.Sort_G_TK_SP_list[j][i])
                    Re.append(index)
                Re_list.append(Re)

            print("Re_list", Re_list)

            # 找到Alist
            A_list = []
            for j in range(len(SP_series_list)):
                A = []
                for i in range(len(indices_without_number_list[j]) - 1):
                    A.append((-1 / (float(self.Sort_G_TK_SP_list[j][i + 1]) - float(self.Sort_G_TK_SP_list[j][i])) * np.log(
                    SP_series_list[j][Re_list[j][i + 1]][1:] / SP_series_list[j][Re_list[j][i]][1:])).reset_index(drop=True))
                A_list.append(A)
            #print("A_list",A_list)

            # 找到K-list
            K_list = []
            for j in range(len(SP_series_list)):
                K = []
                for i in range(len(indices_without_number_list[j]) - 1):
                    K.append(SP_series_list[j][Re_list[j][i]][1:].reset_index(drop=True) / np.exp(
                    (-1 * A_list[j][i]) * self.Sort_G_TK_SP_list[j][i]))
                K_list.append(K)
            #print("K_list", K_list)

            # 厚度分割list-*
            self.GCF_TK_list = []
            try:
                if float(self.G_aK_TK_Start.text()) < float(self.G_aK_TK_End.text()):
                    i = 0
                    while float(self.G_aK_TK_Start.text()) + float(self.G_aK_TK_interval.text()) * i < float(self.G_aK_TK_End.text()):
                        self.GCF_TK_list.append(float(self.G_aK_TK_Start.text()) + float(self.G_aK_TK_interval.text()) * i)
                        i += 1
                    self.GCF_TK_list.append(float(self.G_aK_TK_End.text()))

                elif float(self.G_aK_TK_End.text()) < float(self.G_aK_TK_Start.text()):
                    i = 0
                    while float(self.G_aK_TK_End.text()) + float(self.G_aK_TK_interval.text()) * i < float(self.G_aK_TK_Start.text()):
                        self.GCF_TK_list.append(float(self.G_aK_TK_End.text()) + float(self.G_aK_TK_interval.text()) * i)
                        i += 1
                    self.GCF_TK_list.append(float(self.G_aK_TK_Start.text()))

            except ValueError:
                print("R輸入的數值格式為空，請檢查後重新輸入。")
            print("self.GCF_TK_list", self.GCF_TK_list)

            # 找出AK_GCF_Change_spectrum_list
            self.AK_GCF_Change_spectrum_Series_list = []
            for j in range(len(SP_series_list)):
                AK_GCF_Change_spectrum_Series_2 = []
                for GCF_TK in self.GCF_TK_list:
                    # 先将 R_aK_TK 添加到 self.Sort_TK_SP_list
                    self.Sort_G_TK_SP_list[j].append(GCF_TK)
                    # 然后对更新后的列表进行排序
                    self.Sort_TK_SP_AK_list_2 = sorted(self.Sort_G_TK_SP_list[j], reverse=True)
                    print("self.Sort_TK_SP_AK_list_2",self.Sort_TK_SP_AK_list_2)
                    print("GCF_TK",GCF_TK)
                    # 4種情況
                    if GCF_TK == max(self.Sort_TK_SP_AK_list_2):
                        print("max")
                        AK_GCF_Change_spectrum_Series = K_list[j][0] * np.exp(-1 * A_list[j][0] * GCF_TK)
                        AK_GCF_Change_spectrum_Series_2.append(AK_GCF_Change_spectrum_Series)
                        # print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                        # return AK_GCF_Change_spectrum_Series
                    elif GCF_TK == min(self.Sort_TK_SP_AK_list_2):
                        print("min")
                        AK_GCF_Change_spectrum_Series = K_list[j][len(K_list[j]) - 1] * np.exp(
                            -1 * A_list[j][len(A_list[j]) - 1] * GCF_TK)
                        AK_GCF_Change_spectrum_Series_2.append(AK_GCF_Change_spectrum_Series)

                    elif GCF_TK in self.TK_G_SP_list[j]:
                        print("equal")
                        AK_GCF_Change_spectrum_Series = SP_series_list[j][self.TK_G_SP_list[j].index(GCF_TK)][1:].reset_index(
                            drop=True)
                        AK_GCF_Change_spectrum_Series_2.append(AK_GCF_Change_spectrum_Series)
                        # print("AK_GCF_Change_spectrum_Series",AK_GCF_Change_spectrum_Series)
                        # return AK_GCF_Change_spectrum_Series
                    else:
                        print("mid")
                        position = self.Sort_TK_SP_AK_list_2.index(GCF_TK)
                        print("position",position)
                        #print("K_list[j]",K_list[j])
                        AK_GCF_Change_spectrum_Series = K_list[j][position - 1] * np.exp(
                            -1 * A_list[j][position - 1] * GCF_TK)
                        AK_GCF_Change_spectrum_Series_2.append(AK_GCF_Change_spectrum_Series)
                    self.Sort_G_TK_SP_list[j].pop()
                self.AK_GCF_Change_spectrum_Series_list.append(AK_GCF_Change_spectrum_Series_2)
            self.CIEparameter()
            # BLU +Cell part
            self.cell_blu_total_spectrum = self.calculate_BLU() * self.calculate_layer1() * self.calculate_layer2() \
                                           * self.calculate_layer3() * self.calculate_layer4() * self.calculate_layer5() \
                                           * self.calculate_layer6()
            # 這個迴圈將遍歷 self.AK_RCF_Change_spectrum_Series_list 的每個子列表
            for i in range(len(self.AK_GCF_Change_spectrum_Series_list)):
                # 這個迴圈將遍歷當前子列表中的每個 Series 對象
                for j in range(len(self.AK_GCF_Change_spectrum_Series_list[i])):
                    # 將每個 Series 元素與 self.cell_blu_total_spectrum 相乘
                    self.AK_GCF_Change_spectrum_Series_list[i][j] = self.AK_GCF_Change_spectrum_Series_list[i][
                                                                        j] * self.cell_blu_total_spectrum

            # 此時 self.AK_RCF_Change_spectrum_Series_list 應已更新為乘以 self.cell_blu_total_spectrum 後的值
            self.G_Change_X_list = []

            for group in self.AK_GCF_Change_spectrum_Series_list:
                group_G_Change_X = []
                for series in group:
                    # 將每個 Series 元素與 self.CIE_spectrum_Series_X 相乘
                    multiplied_series = series.multiply(self.CIE_spectrum_Series_X)
                    group_G_Change_X.append(multiplied_series)
                self.G_Change_X_list.append(group_G_Change_X)
            # print("self.G_Change_X_list",self.G_Change_X_list)

            self.G_Change_Y_list = []

            for group in self.AK_GCF_Change_spectrum_Series_list:
                group_G_Change_Y = []
                for series in group:
                    # 將每個 Series 元素與 self.CIE_spectrum_Series_Y 相乘
                    multiplied_series = series.multiply(self.CIE_spectrum_Series_Y)
                    group_G_Change_Y.append(multiplied_series)
                self.G_Change_Y_list.append(group_G_Change_Y)

            self.G_Change_Z_list = []

            for group in self.AK_GCF_Change_spectrum_Series_list:
                group_G_Change_Z = []
                for series in group:
                    # 將每個 Series 元素與 self.CIE_spectrum_Series_Z 相乘
                    multiplied_series = series.multiply(self.CIE_spectrum_Series_Z)
                    group_G_Change_Z.append(multiplied_series)
                self.G_Change_Z_list.append(group_G_Change_Z)
            # X_Sum
            self.G_X_Change_sim_sum_list = []

            for group in self.G_Change_X_list:
                group_sums = []
                for series in group:
                    # 計算每個 Series 的總和
                    sum_series = series.sum()
                    group_sums.append(sum_series)
                self.G_X_Change_sim_sum_list.append(group_sums)
            print("self.G_X_Change_sim_sum_list",self.G_X_Change_sim_sum_list)
            # Y_Sum
            self.G_Y_Change_sim_sum_list = []

            for group in self.G_Change_Y_list:
                group_sums = []
                for series in group:
                    # 計算每個 Series 的總和
                    sum_series = series.sum()
                    group_sums.append(sum_series)
                self.G_Y_Change_sim_sum_list.append(group_sums)
            print("self.G_Y_Change_sim_sum_list",self.G_Y_Change_sim_sum_list)

            # Z_Sum
            self.G_Z_Change_sim_sum_list = []

            for group in self.G_Change_Z_list:
                group_sums = []
                for series in group:
                    # 計算每個 Series 的總和
                    sum_series = series.sum()
                    group_sums.append(sum_series)
                self.G_Z_Change_sim_sum_list.append(group_sums)
            print("self.G_Z_Change_sim_sum_list",self.G_Z_Change_sim_sum_list)

            # G_x
            self.G_Change_x_sim = []
            # 確保所有的列結構相同
            for group_index in range(len(self.G_X_Change_sim_sum_list)):
                group_sim = []
                for item_index in range(len(self.G_X_Change_sim_sum_list[group_index])):
                    X = self.G_X_Change_sim_sum_list[group_index][item_index]
                    Y = self.G_Y_Change_sim_sum_list[group_index][item_index]
                    Z = self.G_Z_Change_sim_sum_list[group_index][item_index]

                    # 计算比例，注意要避免除以零
                    total = X + Y + Z
                    if total > 0:
                        sim = X / total
                    else:
                        sim = 0
                    group_sim.append(sim)
                self.G_Change_x_sim.append(group_sim)

            # G_y
            self.G_Change_y_sim = []
            # 確保所有的列結構相同
            for group_index in range(len(self.G_Y_Change_sim_sum_list)):
                group_sim = []
                for item_index in range(len(self.G_Y_Change_sim_sum_list[group_index])):
                    X = self.G_X_Change_sim_sum_list[group_index][item_index]
                    Y = self.G_Y_Change_sim_sum_list[group_index][item_index]
                    Z = self.G_Z_Change_sim_sum_list[group_index][item_index]

                    # 计算比例，注意要避免除以零
                    total = X + Y + Z
                    if total > 0:
                        sim = Y / total
                    else:
                        sim = 0
                    group_sim.append(sim)
                self.G_Change_y_sim.append(group_sim)
            # # G_T
            # self.G_Change_T_sim = []
            # # 確保所有的列結構相同
            # for group_index in range(len(self.G_Z_Change_sim_sum_list)):
            #     group_sim = []
            #     for item_index in range(len(self.G_Z_Change_sim_sum_list[group_index])):
            #         X = self.G_X_Change_sim_sum_list[group_index][item_index]
            #         Y = self.G_Y_Change_sim_sum_list[group_index][item_index]
            #         Z = self.G_Z_Change_sim_sum_list[group_index][item_index]
            #
            #         # 计算比例，注意要避免除以零
            #         total = X + Y + Z
            #         if total > 0:
            #             sim = 3 * Y
            #         else:
            #             sim = 0
            #         group_sim.append(sim)
            #     self.G_Change_T_sim.append(group_sim)

            # 關閉連線
            connection_GCF_Change.close()
            print("self.G_Change_x_sim",self.G_Change_x_sim)
            print("self.G_Change_y_sim",self.G_Change_y_sim)
            # print("self.G_Change_T_sim",self.G_Change_T_sim)
            self.G_Change_wave_sim = []
            self.G_Change_purity_sim = []

            # xy_n 是觀察者使用的光源（例如 D65）
            xy_n = self.observer_D65

            for group_index in range(len(self.G_Change_x_sim)):
                group_wave = []
                group_purity = []

                for item_index in range(len(self.G_Change_x_sim[group_index])):
                    xy = [self.G_Change_x_sim[group_index][item_index], self.G_Change_y_sim[group_index][item_index]]

                    # 計算主波長和CIE坐標
                    dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)

                    # 計算距離
                    distance_from_white = math.sqrt((xy[0] - xy_n[0]) ** 2 + (xy[1] - xy_n[1]) ** 2)
                    distance_from_black = math.sqrt(
                        (CIEcoordinate[0] - xy_n[0]) ** 2 + (CIEcoordinate[1] - xy_n[1]) ** 2)

                    # 計算色純度
                    purity = distance_from_white / distance_from_black * 100 if distance_from_black != 0 else 0

                    group_wave.append(dominant_wavelength)
                    group_purity.append(purity)

                self.G_Change_wave_sim.append(group_wave)
                self.G_Change_purity_sim.append(group_purity)
            print("self.G_Change_wave_sim",self.G_Change_wave_sim)
            print("self.G_Change_wave_sim",self.G_Change_wave_sim[0][0])
            print("self.G_Change_purity_sim",self.G_Change_purity_sim)
            # 研究GT怎寫
            # 這個迴圈將遍歷 self.AK_GCF_Change_spectrum_Series_list 的每個子列表
            for i in range(len(self.AK_GCF_Change_spectrum_Series_list)):
                # 這個迴圈將遍歷當前子列表中的每個 Series 對象
                for j in range(len(self.AK_GCF_Change_spectrum_Series_list[i])):
                    # 將每個 Series 元素與 self.cell_blu_total_spectrum 相乘
                    self.AK_GCF_Change_spectrum_Series_list[i][j] = \
                        self.AK_GCF_Change_spectrum_Series_list[i][
                            j] / self.calculate_BLU() * self.CIE_spectrum_Series_Y

            self.BLUY = self.calculate_BLU() * self.CIE_spectrum_Series_Y
            # print("改過的self.AK_RCF_Change_spectrum_Series_list",
            #       self.AK_RCF_Change_spectrum_Series_list)
            # 先計算 self.BLUY 的總和
            bluy_sum = self.BLUY.sum()

            # 初始化 self.R_Change_T_sim 為空列表
            self.G_Change_T_sim = []

            # 遍歷 self.AK_RCF_Change_spectrum_Series_list 中的每一組 Series 列表
            for group in self.AK_GCF_Change_spectrum_Series_list:
                # 初始化當前組的結果列表
                group_results = []
                # 遍歷當前組中的每一個 Series 對象
                for series in group:
                    # 計算當前 Series 對象的總和乘以 self.BLUY 的總和，並將結果加到當前組的結果列表上
                    group_results.append(series.sum() / bluy_sum * 100)
                # 將當前組的結果列表加到 self.R_Change_T_sim 上
                self.G_Change_T_sim.append(group_results)
            print("self.G_Change_T_sim-2", self.G_Change_T_sim)

        elif self.G_aK_mode.currentText() == "自訂":
            connection_GCF_Change = sqlite3.connect("GCF_Change_spectrum.db")
            cursor_GCF_Change = connection_GCF_Change.cursor()
            # 取得Rdiffer資料
            column_name_GCF_Change = self.G_aK_box.currentText()
            # print("column_name_RCF_Change",column_name_RCF_Change)
            table_name_GCF_Change = self.G_aK_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_GCF_Change = f"SELECT * FROM '{table_name_GCF_Change}';"
            cursor_GCF_Change.execute(query_GCF_Change)
            result_GCF_Change = cursor_GCF_Change.fetchall()

            # 找到指定標題的欄位索引
            self.header_GCF_Change_list_2 = [column[0] for column in cursor_GCF_Change.description]

            # 移除所有標題中的編號
            self.header_GCF_Change_list_2 = [re.sub(r'_\d+$', '', col) for col in self.header_GCF_Change_list_2]
            # print("self.header_GCF_Change_list_2", self.header_GCF_Change_list_2)
            # 自訂標題list
            # 找到指定標題的欄位索引
            header_GCF_Change = [column[0] for column in cursor_GCF_Change.description]

            # 移除所有標題中的編號
            header_GCF_Change = [re.sub(r'_\d+$', '', col) for col in header_GCF_Change]
            header_GCF_Change = header_GCF_Change[1:]
            print("header_GCF_Change", header_GCF_Change)

            # 創建一個字典來跟踪每個標題的索引
            header_indices = {}

            for index, header in enumerate(header_GCF_Change, start=1):
                if header not in header_indices:
                    header_indices[header] = []
                header_indices[header].append(index)

            # 創建不重複標題列表
            self.header_GCF_Change_list = list(header_indices.keys())
            print("self.header_GCF_Change_list",self.header_GCF_Change_list)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(self.header_GCF_Change_list_2) if col == column_name_GCF_Change]
            print("indices_without_number", indices_without_number)

            # 使用 pd.to_numeric 转换为数值类型，并处理无法转换的值
            SP_series_list = [
                pd.to_numeric(pd.Series([row[i] for row in result_GCF_Change]), errors='coerce').fillna(0)
                for i in indices_without_number
            ]
            #print("SP_series_list", SP_series_list)

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
            #print("A_list", A_list)
            K_list = []
            for i in range(len(indices_without_number) - 1):
                K_list.append(SP_series_list[Re_list[i]][1:].reset_index(drop=True) / np.exp(
                    (-1 * A_list[i]) * Sort_TK_SP_list[i]))
            #print("K_list", K_list)
            if self.G_aK_TK_edit.text() == "":
                AK_GCF_Change_spectrum_Series = 1
                print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                return AK_GCF_Change_spectrum_Series
            elif self.G_aK_TK_edit.text() != "":

                G_aK_TK = float(self.G_aK_TK_edit.text())
                # 厚度分割list-*
                self.GCF_TK_list = [G_aK_TK]
                print("self.GCF_TK_list",self.GCF_TK_list)
                # 先将 R_aK_TK 添加到 Sort_TK_SP_list
                Sort_TK_SP_list.append(G_aK_TK)
                # 然后对更新后的列表进行排序
                Sort_TK_SP_AK_list = sorted(Sort_TK_SP_list, reverse=True)
                print("Sort_TK_SP_AK_list", Sort_TK_SP_AK_list)

                # 4種情況
                if G_aK_TK == max(Sort_TK_SP_AK_list):
                    print("max")
                    AK_GCF_Change_spectrum_Series = K_list[0] * np.exp(-1 * A_list[0] * G_aK_TK)
                    # print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                    # return AK_GCF_Change_spectrum_Series
                elif G_aK_TK == min(Sort_TK_SP_AK_list):
                    print("min")
                    AK_GCF_Change_spectrum_Series = K_list[len(K_list) - 1] * np.exp(
                        -1 * A_list[len(A_list) - 1] * G_aK_TK)
                    #print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # 在Color_Table上顯示

                elif G_aK_TK in TK_SP_list:
                    print("equal")
                    AK_GCF_Change_spectrum_Series = SP_series_list[TK_SP_list.index(G_aK_TK)][1:].reset_index(drop=True)
                    #print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # 在Color_Table上顯示

                else:
                    print("mid")
                    position = Sort_TK_SP_AK_list.index(G_aK_TK)
                    AK_GCF_Change_spectrum_Series = K_list[position - 1] * np.exp(-1 * A_list[position - 1] * G_aK_TK)
                    #print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                self.CIEparameter()
                # BLU +Cell part
                self.cell_blu_total_spectrum = self.calculate_BLU() * self.calculate_layer1() * self.calculate_layer2() \
                                               * self.calculate_layer3() * self.calculate_layer4() * self.calculate_layer5() \
                                               * self.calculate_layer6()
                AK_GCF_Change_spectrum_Series = AK_GCF_Change_spectrum_Series * self.cell_blu_total_spectrum
                G_Change_X_sum = (AK_GCF_Change_spectrum_Series * self.CIE_spectrum_Series_X).sum()
                self.G_X_Change_sim_sum_list = [[G_Change_X_sum]]
                print("self.G_X_Change_sim_sum_list",self.G_X_Change_sim_sum_list)
                G_Change_Y_sum = (AK_GCF_Change_spectrum_Series * self.CIE_spectrum_Series_Y).sum()
                self.G_Y_Change_sim_sum_list = [[G_Change_Y_sum]]
                print("self.G_Y_Change_sim_sum_list",self.G_Y_Change_sim_sum_list)
                G_Change_Z_sum = (AK_GCF_Change_spectrum_Series * self.CIE_spectrum_Series_Z).sum()
                self.G_Z_Change_sim_sum_list = [[G_Change_Z_sum]]
                print("self.G_Z_Change_sim_sum_list",self.G_Z_Change_sim_sum_list)
                G_Change_Total = G_Change_X_sum + G_Change_Y_sum + G_Change_Z_sum
                # G_x
                self.G_Change_x_sim = [[G_Change_X_sum/(G_Change_Total)]]
                print("self.G_Change_x_sim",self.G_Change_x_sim)
                # G_y
                self.G_Change_y_sim = [[G_Change_Y_sum / (G_Change_Total)]]
                print("self.G_Change_y_sim",self.G_Change_y_sim)
                # G_T
                G_Change_T_sum = (AK_GCF_Change_spectrum_Series /self.calculate_BLU() * self.CIE_spectrum_Series_Y).sum() / ((self.calculate_BLU()* self.CIE_spectrum_Series_Y).sum()) * 100
                self.G_Change_T_sim = [[G_Change_T_sum]]
                print("self.G_Change_T_sim",self.G_Change_T_sim)
                self.G_Change_wave_sim = []
                self.G_Change_purity_sim = []

                # xy_n 是觀察者使用的光源（例如 D65）
                xy_n = self.observer_D65

                for group_index in range(len(self.G_Change_x_sim)):
                    group_wave = []
                    group_purity = []

                    for item_index in range(len(self.G_Change_x_sim[group_index])):
                        xy = [self.G_Change_x_sim[group_index][item_index],
                              self.G_Change_y_sim[group_index][item_index]]

                        # 計算主波長和CIE坐標
                        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)

                        # 計算距離
                        distance_from_white = math.sqrt((xy[0] - xy_n[0]) ** 2 + (xy[1] - xy_n[1]) ** 2)
                        distance_from_black = math.sqrt(
                            (CIEcoordinate[0] - xy_n[0]) ** 2 + (CIEcoordinate[1] - xy_n[1]) ** 2)

                        # 計算色純度
                        purity = distance_from_white / distance_from_black * 100 if distance_from_black != 0 else 0

                        group_wave.append(dominant_wavelength)
                        group_purity.append(purity)

                    self.G_Change_wave_sim.append(group_wave)
                    self.G_Change_purity_sim.append(group_purity)
                print("self.G_Change_wave_sim", self.G_Change_wave_sim)
                print("self.G_Change_wave_sim", self.G_Change_wave_sim[0][0])
                print("self.G_Change_purity_sim", self.G_Change_purity_sim)


            # 關閉連線
            connection_GCF_Change.close()

    def calculate_BCF_Change(self):
        if self.B_aK_mode.currentText() == "模擬":
            connection_BCF_Change = sqlite3.connect("BCF_Change_spectrum.db")
            cursor_BCF_Change = connection_BCF_Change.cursor()
            # 取得Change資料
            table_name_BCF_Change = self.B_aK_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_BCF_Change = f"SELECT * FROM '{table_name_BCF_Change}';"
            cursor_BCF_Change.execute(query_BCF_Change)
            result_BCF_Change = cursor_BCF_Change.fetchall()
            #print("result_BCF_Change",result_BCF_Change)
            # 找到指定標題的欄位索引
            header_BCF_Change = [column[0] for column in cursor_BCF_Change.description]

            # 移除所有標題中的編號
            header_BCF_Change = [re.sub(r'_\d+$', '', col) for col in header_BCF_Change]
            header_BCF_Change = header_BCF_Change[1:]
            print("header_BCF_Change", header_BCF_Change)

            # 創建一個字典來跟踪每個標題的索引
            header_indices = {}

            for index, header in enumerate(header_BCF_Change, start=1):
                if header not in header_indices:
                    header_indices[header] = []
                header_indices[header].append(index)

            # 創建不重複標題列表
            self.header_BCF_Change_list = list(header_indices.keys())

            # 創建對應的索引列表
            indices_without_number_list = list(header_indices.values())

            # 打印結果（可選）
            print("self.header_BCF_Change_list:", self.header_BCF_Change_list)
            print("indices_without_number_list:", indices_without_number_list)

            # 使用 pd.to_numeric 转换为数值类型，并处理无法转换的值
            SP_series_list = []
            for j in range(len(self.header_BCF_Change_list)):
                SP_series = []
                for i in indices_without_number_list[j]:
                    # 這裡假設 result_RCF_Change 是一個二維列表，其中每個子列表代表一行數據
                    SP_series.append(
                        pd.to_numeric(pd.Series([row[i] for row in result_BCF_Change]), errors='coerce').fillna(0))
                SP_series_list.append(SP_series)
            # 每個色組獨立一個list[[第一組色組list],[第二組色組list],[第三組色組list]]
            #print("SP_series_list", SP_series_list)

            # TK_SP準備分厚度順序
            self.TK_B_SP_list = []
            self.Sort_B_TK_SP_list = []
            for j in range(len(SP_series_list)):
                TK_SP = []
                for i in range(len(indices_without_number_list[j])):
                    TK_SP.append(SP_series_list[j][i].iloc[0])
                self.Sort_B_TK_SP = sorted(TK_SP,reverse=True)
                self.Sort_B_TK_SP_list.append(self.Sort_B_TK_SP)
                self.TK_B_SP_list.append(TK_SP)
            # self.Sort_B_TK_SP_list = sorted(self.TK_B_SP_list)
            print(" self.TK_B_SP_list", self.TK_B_SP_list)
            print(" self.Sort_B_TK_SP_list", self.Sort_B_TK_SP_list)

            # 反轉找回原本的順序list
            Re_list = []
            for j in range(len(SP_series_list)):
                Re = []
                for i in range(len(indices_without_number_list[j])):
                    # 尋找 self.Sort_TK_SP_list[j][i] 在 self.TK_SP_list[j] 中的索引
                    print("self.TK_B_SP_list[j]",self.TK_B_SP_list[j])
                    print("self.Sort_B_TK_SP_list[j][i]",self.Sort_B_TK_SP_list[j][i])
                    index = self.TK_B_SP_list[j].index(self.Sort_B_TK_SP_list[j][i])
                    Re.append(index)
                Re_list.append(Re)

            #print("Re_list", Re_list)

            # 找到Alist
            A_list = []
            for j in range(len(SP_series_list)):
                A = []
                for i in range(len(indices_without_number_list[j]) - 1):
                    A.append((-1 / (float(self.Sort_B_TK_SP_list[j][i + 1]) - float(self.Sort_B_TK_SP_list[j][i])) * np.log(
                    SP_series_list[j][Re_list[j][i + 1]][1:] / SP_series_list[j][Re_list[j][i]][1:])).reset_index(drop=True))
                A_list.append(A)
            #print("A_list",A_list)

            # 找到K-list
            K_list = []
            for j in range(len(SP_series_list)):
                K = []
                for i in range(len(indices_without_number_list[j]) - 1):
                    K.append(SP_series_list[j][Re_list[j][i]][1:].reset_index(drop=True) / np.exp(
                    (-1 * A_list[j][i]) * self.Sort_B_TK_SP_list[j][i]))
                K_list.append(K)
            #print("K_list", K_list)

            # 厚度分割list-*
            self.BCF_TK_list = []
            try:
                if float(self.B_aK_TK_Start.text()) < float(self.B_aK_TK_End.text()):
                    i = 0
                    while float(self.B_aK_TK_Start.text()) + float(self.B_aK_TK_interval.text()) * i < float(self.B_aK_TK_End.text()):
                        self.BCF_TK_list.append(float(self.B_aK_TK_Start.text()) + float(self.B_aK_TK_interval.text()) * i)
                        i += 1
                    self.BCF_TK_list.append(float(self.B_aK_TK_End.text()))

                elif float(self.B_aK_TK_End.text()) < float(self.B_aK_TK_Start.text()):
                    i = 0
                    while float(self.B_aK_TK_End.text()) + float(self.B_aK_TK_interval.text()) * i < float(self.B_aK_TK_Start.text()):
                        self.BCF_TK_list.append(float(self.B_aK_TK_End.text()) + float(self.B_aK_TK_interval.text()) * i)
                        i += 1
                    self.BCF_TK_list.append(float(self.B_aK_TK_Start.text()))

            except ValueError:
                print("R輸入的數值格式為空，請檢查後重新輸入。")
            print("self.BCF_TK_list", self.BCF_TK_list)

            # 找出AK_BCF_Change_spectrum_list
            self.AK_BCF_Change_spectrum_Series_list = []
            for j in range(len(SP_series_list)):
                AK_BCF_Change_spectrum_Series_2 = []
                for BCF_TK in self.BCF_TK_list:
                    # 先将 B_aK_TK 添加到 self.Sort_TK_SP_list
                    self.Sort_B_TK_SP_list[j].append(BCF_TK)
                    # 然后对更新后的列表进行排序
                    self.Sort_TK_SP_AK_list_2 = sorted(self.Sort_B_TK_SP_list[j], reverse=True)
                    print("self.Sort_TK_SP_AK_list_2",self.Sort_TK_SP_AK_list_2)
                    print("BCF_TK",BCF_TK)
                    # 4種情況
                    if BCF_TK == max(self.Sort_TK_SP_AK_list_2):
                        print("max")
                        AK_BCF_Change_spectrum_Series = K_list[j][0] * np.exp(-1 * A_list[j][0] * BCF_TK)
                        AK_BCF_Change_spectrum_Series_2.append(AK_BCF_Change_spectrum_Series)
                        # print("AK_GCF_Change_spectrum_Series", AK_GCF_Change_spectrum_Series)
                        # return AK_GCF_Change_spectrum_Series
                    elif BCF_TK == min(self.Sort_TK_SP_AK_list_2):
                        print("min")
                        #print("K_list",K_list)
                        #print("K_list[j]", K_list[j])
                        print()
                        AK_BCF_Change_spectrum_Series = K_list[j][len(K_list[j]) - 1] * np.exp(
                            -1 * A_list[j][len(A_list[j]) - 1] * BCF_TK)
                        AK_BCF_Change_spectrum_Series_2.append(AK_BCF_Change_spectrum_Series)

                    elif BCF_TK in self.TK_B_SP_list[j]:
                        print("equal")
                        AK_BCF_Change_spectrum_Series = SP_series_list[j][self.TK_B_SP_list[j].index(BCF_TK)][1:].reset_index(
                            drop=True)
                        AK_BCF_Change_spectrum_Series_2.append(AK_BCF_Change_spectrum_Series)
                        # print("AK_BCF_Change_spectrum_Series",AK_BCF_Change_spectrum_Series)
                        # return AK_BCF_Change_spectrum_Series
                    else:
                        print("mid")
                        position = self.Sort_TK_SP_AK_list_2.index(BCF_TK)
                        print("position",position)
                        #print("K_list[j]",K_list[j])
                        AK_BCF_Change_spectrum_Series = K_list[j][position - 1] * np.exp(
                            -1 * A_list[j][position - 1] * BCF_TK)
                        AK_BCF_Change_spectrum_Series_2.append(AK_BCF_Change_spectrum_Series)
                    self.Sort_B_TK_SP_list[j].pop()
                self.AK_BCF_Change_spectrum_Series_list.append(AK_BCF_Change_spectrum_Series_2)
            self.CIEparameter()
            # BLU +Cell part
            self.cell_blu_total_spectrum = self.calculate_BLU() * self.calculate_layer1() * self.calculate_layer2() \
                                           * self.calculate_layer3() * self.calculate_layer4() * self.calculate_layer5() \
                                           * self.calculate_layer6()
            # 這個迴圈將遍歷 self.AK_RCF_Change_spectrum_Series_list 的每個子列表
            for i in range(len(self.AK_BCF_Change_spectrum_Series_list)):
                # 這個迴圈將遍歷當前子列表中的每個 Series 對象
                for j in range(len(self.AK_BCF_Change_spectrum_Series_list[i])):
                    # 將每個 Series 元素與 self.cell_blu_total_spectrum 相乘
                    self.AK_BCF_Change_spectrum_Series_list[i][j] = self.AK_BCF_Change_spectrum_Series_list[i][
                                                                        j] * self.cell_blu_total_spectrum

            # 此時 self.AK_RCF_Change_spectrum_Series_list 應已更新為乘以 self.cell_blu_total_spectrum 後的值
            self.B_Change_X_list = []

            for group in self.AK_BCF_Change_spectrum_Series_list:
                group_B_Change_X = []
                for series in group:
                    # 將每個 Series 元素與 self.CIE_spectrum_Series_X 相乘
                    multiplied_series = series.multiply(self.CIE_spectrum_Series_X)
                    group_B_Change_X.append(multiplied_series)
                self.B_Change_X_list.append(group_B_Change_X)
            # print("self.B_Change_X_list",self.B_Change_X_list)

            self.B_Change_Y_list = []

            for group in self.AK_BCF_Change_spectrum_Series_list:
                group_B_Change_Y = []
                for series in group:
                    # 將每個 Series 元素與 self.CIE_spectrum_Series_Y 相乘
                    multiplied_series = series.multiply(self.CIE_spectrum_Series_Y)
                    group_B_Change_Y.append(multiplied_series)
                self.B_Change_Y_list.append(group_B_Change_Y)

            self.B_Change_Z_list = []

            for group in self.AK_BCF_Change_spectrum_Series_list:
                group_B_Change_Z = []
                for series in group:
                    # 將每個 Series 元素與 self.CIE_spectrum_Series_Z 相乘
                    multiplied_series = series.multiply(self.CIE_spectrum_Series_Z)
                    group_B_Change_Z.append(multiplied_series)
                self.B_Change_Z_list.append(group_B_Change_Z)
            # X_Sum
            self.B_X_Change_sim_sum_list = []

            for group in self.B_Change_X_list:
                group_sums = []
                for series in group:
                    # 計算每個 Series 的總和
                    sum_series = series.sum()
                    group_sums.append(sum_series)
                self.B_X_Change_sim_sum_list.append(group_sums)
            print("self.B_X_Change_sim_sum_list",self.B_X_Change_sim_sum_list)
            # Y_Sum
            self.B_Y_Change_sim_sum_list = []

            for group in self.B_Change_Y_list:
                group_sums = []
                for series in group:
                    # 計算每個 Series 的總和
                    sum_series = series.sum()
                    group_sums.append(sum_series)
                self.B_Y_Change_sim_sum_list.append(group_sums)
            print("self.B_Y_Change_sim_sum_list",self.B_Y_Change_sim_sum_list)

            # Z_Sum
            self.B_Z_Change_sim_sum_list = []

            for group in self.B_Change_Z_list:
                group_sums = []
                for series in group:
                    # 計算每個 Series 的總和
                    sum_series = series.sum()
                    group_sums.append(sum_series)
                self.B_Z_Change_sim_sum_list.append(group_sums)
            print("self.B_Z_Change_sim_sum_list",self.B_Z_Change_sim_sum_list)

            # B_x
            self.B_Change_x_sim = []
            # 確保所有的列結構相同
            for group_index in range(len(self.B_X_Change_sim_sum_list)):
                group_sim = []
                for item_index in range(len(self.B_X_Change_sim_sum_list[group_index])):
                    X = self.B_X_Change_sim_sum_list[group_index][item_index]
                    Y = self.B_Y_Change_sim_sum_list[group_index][item_index]
                    Z = self.B_Z_Change_sim_sum_list[group_index][item_index]

                    # 计算比例，注意要避免除以零
                    total = X + Y + Z
                    if total > 0:
                        sim = X / total
                    else:
                        sim = 0
                    group_sim.append(sim)
                self.B_Change_x_sim.append(group_sim)

            # B_y
            self.B_Change_y_sim = []
            # 確保所有的列結構相同
            for group_index in range(len(self.B_Y_Change_sim_sum_list)):
                group_sim = []
                for item_index in range(len(self.B_Y_Change_sim_sum_list[group_index])):
                    X = self.B_X_Change_sim_sum_list[group_index][item_index]
                    Y = self.B_Y_Change_sim_sum_list[group_index][item_index]
                    Z = self.B_Z_Change_sim_sum_list[group_index][item_index]

                    # 计算比例，注意要避免除以零
                    total = X + Y + Z
                    if total > 0:
                        sim = Y / total
                    else:
                        sim = 0
                    group_sim.append(sim)
                self.B_Change_y_sim.append(group_sim)
            # # B_T
            # self.B_Change_T_sim = []
            # # 確保所有的列結構相同
            # for group_index in range(len(self.B_Z_Change_sim_sum_list)):
            #     group_sim = []
            #     for item_index in range(len(self.B_Z_Change_sim_sum_list[group_index])):
            #         X = self.B_X_Change_sim_sum_list[group_index][item_index]
            #         Y = self.B_Y_Change_sim_sum_list[group_index][item_index]
            #         Z = self.B_Z_Change_sim_sum_list[group_index][item_index]
            #
            #         # 计算比例，注意要避免除以零
            #         total = X + Y + Z
            #         if total > 0:
            #             sim = 3 * Y
            #         else:
            #             sim = 0
            #         group_sim.append(sim)
            #     self.B_Change_T_sim.append(group_sim)

            # 關閉連線
            connection_BCF_Change.close()
            print("self.B_Change_x_sim",self.B_Change_x_sim)
            print("self.B_Change_y_sim",self.B_Change_y_sim)
            # print("self.B_Change_T_sim",self.B_Change_T_sim)
            self.B_Change_wave_sim = []
            self.B_Change_purity_sim = []

            # xy_n 是觀察者使用的光源（例如 D65）
            xy_n = self.observer_D65

            for group_index in range(len(self.B_Change_x_sim)):
                group_wave = []
                group_purity = []

                for item_index in range(len(self.B_Change_x_sim[group_index])):
                    xy = [self.B_Change_x_sim[group_index][item_index], self.B_Change_y_sim[group_index][item_index]]

                    # 計算主波長和CIE坐標
                    dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)

                    # 計算距離
                    distance_from_white = math.sqrt((xy[0] - xy_n[0]) ** 2 + (xy[1] - xy_n[1]) ** 2)
                    distance_from_black = math.sqrt(
                        (CIEcoordinate[0] - xy_n[0]) ** 2 + (CIEcoordinate[1] - xy_n[1]) ** 2)

                    # 計算色純度
                    purity = distance_from_white / distance_from_black * 100 if distance_from_black != 0 else 0

                    group_wave.append(dominant_wavelength)
                    group_purity.append(purity)

                self.B_Change_wave_sim.append(group_wave)
                self.B_Change_purity_sim.append(group_purity)
            print("self.B_Change_wave_sim",self.B_Change_wave_sim)
            print("self.B_Change_wave_sim",self.B_Change_wave_sim[0][0])
            print("self.B_Change_purity_sim",self.B_Change_purity_sim)
            # 研究BT怎寫
            # 這個迴圈將遍歷 self.AK_BCF_Change_spectrum_Series_list 的每個子列表
            for i in range(len(self.AK_BCF_Change_spectrum_Series_list)):
                # 這個迴圈將遍歷當前子列表中的每個 Series 對象
                for j in range(len(self.AK_BCF_Change_spectrum_Series_list[i])):
                    # 將每個 Series 元素與 self.cell_blu_total_spectrum 相乘
                    self.AK_BCF_Change_spectrum_Series_list[i][j] = \
                        self.AK_BCF_Change_spectrum_Series_list[i][
                            j] / self.calculate_BLU() * self.CIE_spectrum_Series_Y

            self.BLUY = self.calculate_BLU() * self.CIE_spectrum_Series_Y
            # print("改過的self.AK_RCF_Change_spectrum_Series_list",
            #       self.AK_RCF_Change_spectrum_Series_list)
            # 先計算 self.BLUY 的總和
            bluy_sum = self.BLUY.sum()

            # 初始化 self.R_Change_T_sim 為空列表
            self.B_Change_T_sim = []

            # 遍歷 self.AK_BCF_Change_spectrum_Series_list 中的每一組 Series 列表
            for group in self.AK_BCF_Change_spectrum_Series_list:
                # 初始化當前組的結果列表
                group_results = []
                # 遍歷當前組中的每一個 Series 對象
                for series in group:
                    # 計算當前 Series 對象的總和乘以 self.BLUY 的總和，並將結果加到當前組的結果列表上
                    group_results.append(series.sum() / bluy_sum * 100)
                # 將當前組的結果列表加到 self.R_Change_T_sim 上
                self.B_Change_T_sim.append(group_results)
            print("self.B_Change_T_sim-2", self.B_Change_T_sim)

        elif self.B_aK_mode.currentText() == "自訂":
            connection_BCF_Change = sqlite3.connect("BCF_Change_spectrum.db")
            cursor_BCF_Change = connection_BCF_Change.cursor()
            # 取得Rdiffer資料
            column_name_BCF_Change = self.B_aK_box.currentText()
            # print("column_name_RCF_Change",column_name_RCF_Change)
            table_name_BCF_Change = self.B_aK_table.currentText()
            # 使用正確的引號包裹表名和列名
            query_BCF_Change = f"SELECT * FROM '{table_name_BCF_Change}';"
            cursor_BCF_Change.execute(query_BCF_Change)
            result_BCF_Change = cursor_BCF_Change.fetchall()

            # 找到指定標題的欄位索引
            self.header_BCF_Change_list_2 = [column[0] for column in cursor_BCF_Change.description]

            # 移除所有標題中的編號
            self.header_BCF_Change_list_2 = [re.sub(r'_\d+$', '', col) for col in self.header_BCF_Change_list_2]
            #print("self.header_BCF_Change_list_2", self.header_BCF_Change_list_2)
            # 自訂標題list
            # 找到指定標題的欄位索引
            header_BCF_Change = [column[0] for column in cursor_BCF_Change.description]

            # 移除所有標題中的編號
            header_BCF_Change = [re.sub(r'_\d+$', '', col) for col in header_BCF_Change]
            header_BCF_Change = header_BCF_Change[1:]
            print("header_BCF_Change", header_BCF_Change)

            # 創建一個字典來跟踪每個標題的索引
            header_indices = {}

            for index, header in enumerate(header_BCF_Change, start=1):
                if header not in header_indices:
                    header_indices[header] = []
                header_indices[header].append(index)

            # 創建不重複標題列表
            self.header_BCF_Change_list = list(header_indices.keys())
            print("self.header_BCF_Change_list",self.header_BCF_Change_list)

            # 找到移除編號後的名稱在標題的位置
            indices_without_number = [i for i, col in enumerate(self.header_BCF_Change_list_2) if col == column_name_BCF_Change]
            print("indices_without_number", indices_without_number)

            # 使用 pd.to_numeric 转换为数值类型，并处理无法转换的值
            SP_series_list = [
                pd.to_numeric(pd.Series([row[i] for row in result_BCF_Change]), errors='coerce').fillna(0)
                for i in indices_without_number
            ]
            #print("SP_series_list", SP_series_list)

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
            #print("A_list", A_list)
            K_list = []
            for i in range(len(indices_without_number) - 1):
                K_list.append(SP_series_list[Re_list[i]][1:].reset_index(drop=True) / np.exp(
                    (-1 * A_list[i]) * Sort_TK_SP_list[i]))
            #print("K_list", K_list)
            if self.B_aK_TK_edit.text() == "":
                AK_BCF_Change_spectrum_Series = 1
                print("AK_BCF_Change_spectrum_Series", AK_BCF_Change_spectrum_Series)
                return AK_BCF_Change_spectrum_Series
            elif self.B_aK_TK_edit.text() != "":

                B_aK_TK = float(self.B_aK_TK_edit.text())
                # 厚度分割list-*
                self.BCF_TK_list = [B_aK_TK]
                print("self.BCF_TK_list",self.BCF_TK_list)
                # 先将 B_aK_TK 添加到 Sort_TK_SP_list
                Sort_TK_SP_list.append(B_aK_TK)
                # 然后对更新后的列表进行排序
                Sort_TK_SP_AK_list = sorted(Sort_TK_SP_list, reverse=True)
                print("Sort_TK_SP_AK_list", Sort_TK_SP_AK_list)

                # 4種情況
                if B_aK_TK == max(Sort_TK_SP_AK_list):
                    print("max")
                    AK_BCF_Change_spectrum_Series = K_list[0] * np.exp(-1 * A_list[0] * B_aK_TK)
                    # print("AK_BCF_Change_spectrum_Series", AK_BCF_Change_spectrum_Series)
                    # return AK_BCF_Change_spectrum_Series
                elif B_aK_TK == min(Sort_TK_SP_AK_list):
                    print("min")
                    AK_BCF_Change_spectrum_Series = K_list[len(K_list) - 1] * np.exp(
                        -1 * A_list[len(A_list) - 1] * B_aK_TK)
                    #print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                    # 在Color_Table上顯示

                elif B_aK_TK in TK_SP_list:
                    print("equal")
                    AK_BCF_Change_spectrum_Series = SP_series_list[TK_SP_list.index(B_aK_TK)][1:].reset_index(drop=True)
                    #print("AK_BCF_Change_spectrum_Series", AK_BCF_Change_spectrum_Series)
                    # 在Color_Table上顯示

                else:
                    print("mid")
                    position = Sort_TK_SP_AK_list.index(B_aK_TK)
                    AK_BCF_Change_spectrum_Series = K_list[position - 1] * np.exp(-1 * A_list[position - 1] * B_aK_TK)
                    #print("AK_RCF_Change_spectrum_Series", AK_RCF_Change_spectrum_Series)
                self.CIEparameter()
                # BLU +Cell part
                self.cell_blu_total_spectrum = self.calculate_BLU() * self.calculate_layer1() * self.calculate_layer2() \
                                               * self.calculate_layer3() * self.calculate_layer4() * self.calculate_layer5() \
                                               * self.calculate_layer6()
                AK_BCF_Change_spectrum_Series = AK_BCF_Change_spectrum_Series * self.cell_blu_total_spectrum
                B_Change_X_sum = (AK_BCF_Change_spectrum_Series * self.CIE_spectrum_Series_X).sum()
                self.B_X_Change_sim_sum_list = [[B_Change_X_sum]]
                B_Change_Y_sum = (AK_BCF_Change_spectrum_Series * self.CIE_spectrum_Series_Y).sum()
                self.B_Y_Change_sim_sum_list = [[B_Change_Y_sum]]
                B_Change_Z_sum = (AK_BCF_Change_spectrum_Series * self.CIE_spectrum_Series_Z).sum()
                self.B_Z_Change_sim_sum_list = [[B_Change_Z_sum]]
                B_Change_Total = B_Change_X_sum + B_Change_Y_sum + B_Change_Z_sum
                # B_x
                self.B_Change_x_sim = [[B_Change_X_sum/(B_Change_Total)]]
                print("self.B_Change_x_sim",self.B_Change_x_sim)
                # B_y
                self.B_Change_y_sim = [[B_Change_Y_sum / (B_Change_Total)]]
                print("self.B_Change_y_sim",self.B_Change_y_sim)
                # B_T
                B_Change_T_sum = (AK_BCF_Change_spectrum_Series /self.calculate_BLU() * self.CIE_spectrum_Series_Y).sum() / ((self.calculate_BLU()* self.CIE_spectrum_Series_Y).sum()) * 100
                self.B_Change_T_sim = [[B_Change_T_sum]]
                print("self.B_Change_T_sim",self.B_Change_T_sim)
                self.B_Change_wave_sim = []
                self.B_Change_purity_sim = []

                # xy_n 是觀察者使用的光源（例如 D65）
                xy_n = self.observer_D65

                for group_index in range(len(self.B_Change_x_sim)):
                    group_wave = []
                    group_purity = []

                    for item_index in range(len(self.B_Change_x_sim[group_index])):
                        xy = [self.B_Change_x_sim[group_index][item_index],
                              self.B_Change_y_sim[group_index][item_index]]

                        # 計算主波長和CIE坐標
                        dominant_wavelength, CIEcoordinate, _ = colour.dominant_wavelength(xy, xy_n)

                        # 計算距離
                        distance_from_white = math.sqrt((xy[0] - xy_n[0]) ** 2 + (xy[1] - xy_n[1]) ** 2)
                        distance_from_black = math.sqrt(
                            (CIEcoordinate[0] - xy_n[0]) ** 2 + (CIEcoordinate[1] - xy_n[1]) ** 2)

                        # 計算色純度
                        purity = distance_from_white / distance_from_black * 100 if distance_from_black != 0 else 0

                        group_wave.append(dominant_wavelength)
                        group_purity.append(purity)

                    self.B_Change_wave_sim.append(group_wave)
                    self.B_Change_purity_sim.append(group_purity)
                print("self.B_Change_wave_sim", self.B_Change_wave_sim)
                print("self.B_Change_wave_sim", self.B_Change_wave_sim[0][0])
                print("self.B_Change_purity_sim", self.B_Change_purity_sim)


            # 關閉連線
            connection_BCF_Change.close()

    def calculate_color_W_Change_customize_sim(self):
        # 清除 color_sim_table 從第一行開始以下的所有內容
        for row in range(1, self.color_sim_table.rowCount()):
            for column in range(self.color_sim_table.columnCount()):
                self.color_sim_table.setItem(row, column, QTableWidgetItem(""))
        self.calculate_BLU()
        self.calculate_RCF_Change()
        self.calculate_GCF_Change()
        self.calculate_BCF_Change()
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
        self.BLU_x = RxSxxl_sum_k / (RxSxxl_sum_k + RxSxyl_sum_k + RxSxzl_sum_k)
        self.BLU_y = RxSxyl_sum_k / (RxSxxl_sum_k + RxSxyl_sum_k + RxSxzl_sum_k)

        # W_X_Change_Sim
        self.all_Change_W_X_combinations = []

        # 遍历 R 的所有子列表
        for r_sublist in self.R_X_Change_sim_sum_list:
            # 遍历 G 的所有子列表
            for g_sublist in self.G_X_Change_sim_sum_list:
                # 遍历 B 的所有子列表
                for b_sublist in self.B_X_Change_sim_sum_list:
                    # 遍历 R 子列表中的每个元素
                    for r_value in r_sublist:
                        # 遍历 G 子列表中的每个元素
                        for g_value in g_sublist:
                            # 遍历 B 子列表中的每个元素
                            for b_value in b_sublist:
                                # 将選取的值組合成一個新列表
                                combination = [r_value, g_value, b_value]
                                self.all_Change_W_X_combinations.append(combination)

        print("self.all_Change_W_X_combinations", self.all_Change_W_X_combinations)

        # W_Y_Change_Sim
        self.all_Change_W_Y_combinations = []

        # 遍历 R 的所有子列表
        for r_sublist in self.R_Y_Change_sim_sum_list:
            # 遍历 G 的所有子列表
            for g_sublist in self.G_Y_Change_sim_sum_list:
                # 遍历 B 的所有子列表
                for b_sublist in self.B_Y_Change_sim_sum_list:
                    # 遍历 R 子列表中的每个元素
                    for r_value in r_sublist:
                        # 遍历 G 子列表中的每个元素
                        for g_value in g_sublist:
                            # 遍历 B 子列表中的每个元素
                            for b_value in b_sublist:
                                # 将選取的值組合成一個新列表
                                combination = [r_value, g_value, b_value]
                                self.all_Change_W_Y_combinations.append(combination)

        print("self.all_Change_W_Y_combinations", self.all_Change_W_Y_combinations)

        # W_Z_Change_Sim
        self.all_Change_W_Z_combinations = []

        # 遍历 R 的所有子列表
        for r_sublist in self.R_Z_Change_sim_sum_list:
            # 遍历 G 的所有子列表
            for g_sublist in self.G_Z_Change_sim_sum_list:
                # 遍历 B 的所有子列表
                for b_sublist in self.B_Z_Change_sim_sum_list:
                    # 遍历 R 子列表中的每个元素
                    for r_value in r_sublist:
                        # 遍历 G 子列表中的每个元素
                        for g_value in g_sublist:
                            # 遍历 B 子列表中的每个元素
                            for b_value in b_sublist:
                                # 将選取的值組合成一個新列表
                                combination = [r_value, g_value, b_value]
                                self.all_Change_W_Z_combinations.append(combination)

        print("self.all_Change_W_Z_combinations", self.all_Change_W_Z_combinations)

        self.W_x_Change_sum_list = []
        self.W_y_Change_sum_list = []
        # self.W_T_Change_sum_list = []

        # Wx & Wy------------------------------------------------
        for i in range(len(self.all_Change_W_X_combinations)):
            self.W_x_Change_sum_list.append((self.all_Change_W_X_combinations[i][0] + self.all_Change_W_X_combinations[i][1] +self.all_Change_W_X_combinations[i][2])/
                                            (self.all_Change_W_X_combinations[i][0] + self.all_Change_W_X_combinations[i][1] +self.all_Change_W_X_combinations[i][2] +
                                            self.all_Change_W_Y_combinations[i][0] + self.all_Change_W_Y_combinations[i][1] +self.all_Change_W_Y_combinations[i][2]
                                             + self.all_Change_W_Z_combinations[i][0] + self.all_Change_W_Z_combinations[i][1]+self.all_Change_W_Z_combinations[i][2]))
            self.W_y_Change_sum_list.append((self.all_Change_W_Y_combinations[i][0] + self.all_Change_W_Y_combinations[i][1] +self.all_Change_W_Y_combinations[i][2])/
                                            (self.all_Change_W_X_combinations[i][0] + self.all_Change_W_X_combinations[i][1] +self.all_Change_W_X_combinations[i][2] +
                                            self.all_Change_W_Y_combinations[i][0] + self.all_Change_W_Y_combinations[i][1] +self.all_Change_W_Y_combinations[i][2]
                                             + self.all_Change_W_Z_combinations[i][0] + self.all_Change_W_Z_combinations[i][1]+self.all_Change_W_Z_combinations[i][2]))
            # self.W_T_Change_sum_list.append(self.all_Change_W_Y_combinations[i][0] + self.all_Change_W_Y_combinations[i][1] +self.all_Change_W_Y_combinations[i][2])
        print("self.W_x_Change_sum_list", self.W_x_Change_sum_list)
        print("self.W_y_Change_sum_list", self.W_y_Change_sum_list)
        # print("self.W_T_Change_sum_list", self.W_T_Change_sum_list)

        # RGB_x_Change_sim

        self.RGB_x_change_sim_list = []

        # 遍历 R 的所有子列表
        for r_sublist in self.R_Change_x_sim:
            # 遍历 G 的所有子列表
            for g_sublist in self.G_Change_x_sim:
                # 遍历 B 的所有子列表
                for b_sublist in self.B_Change_x_sim:
                    # 遍历 R 子列表中的每个元素
                    for r_value in r_sublist:
                        # 遍历 G 子列表中的每个元素
                        for g_value in g_sublist:
                            # 遍历 B 子列表中的每个元素
                            for b_value in b_sublist:
                                # 将選取的值組合成一個新列表
                                combination = [r_value, g_value, b_value]
                                self.RGB_x_change_sim_list.append(combination)

        # print(len(self.RGB_x_change_sim_list))

        print("self.RGB_x_change_sim_list",self.RGB_x_change_sim_list)
        print("len(self.RGB_x_change_sim_list)",len(self.RGB_x_change_sim_list))

        # RGB_y_Change_sim

        self.RGB_y_change_sim_list = []

        # 遍历 R 的所有子列表
        for r_sublist in self.R_Change_y_sim:
            # 遍历 G 的所有子列表
            for g_sublist in self.G_Change_y_sim:
                # 遍历 B 的所有子列表
                for b_sublist in self.B_Change_y_sim:
                    # 遍历 R 子列表中的每个元素
                    for r_value in r_sublist:
                        # 遍历 G 子列表中的每个元素
                        for g_value in g_sublist:
                            # 遍历 B 子列表中的每个元素
                            for b_value in b_sublist:
                                # 将選取的值組合成一個新列表
                                combination = [r_value, g_value, b_value]
                                self.RGB_y_change_sim_list.append(combination)

        print("self.RGB_y_change_sim_list", self.RGB_y_change_sim_list)

        # RGB_T_Change_sim

        self.RGB_T_change_sim_list = []

        # 遍历 R 的所有子列表
        for r_sublist in self.R_Change_T_sim:
            # 遍历 G 的所有子列表
            for g_sublist in self.G_Change_T_sim:
                # 遍历 B 的所有子列表
                for b_sublist in self.B_Change_T_sim:
                    # 遍历 R 子列表中的每个元素
                    for r_value in r_sublist:
                        # 遍历 G 子列表中的每个元素
                        for g_value in g_sublist:
                            # 遍历 B 子列表中的每个元素
                            for b_value in b_sublist:
                                # 将選取的值組合成一個新列表
                                combination = [r_value, g_value, b_value]
                                self.RGB_T_change_sim_list.append(combination)

        print("self.RGB_T_change_sim_list", self.RGB_T_change_sim_list)

        # header and TK
        if self.R_aK_mode.currentText() == "模擬":
            self.New_header_RCF_Change_list = [[header] * len(self.RCF_TK_list) for header in self.header_RCF_Change_list]
            print("self.New_header_RCF_Change_list",self.New_header_RCF_Change_list)
            self.New_RCF_TK_list = [list(self.RCF_TK_list) for _ in self.header_RCF_Change_list]
            print("self.New_RCF_TK_list",self.New_RCF_TK_list)
        elif self.R_aK_mode.currentText() == "自訂":
            self.New_header_RCF_Change_list = [[self.R_aK_box.currentText()]]
            print("self.New_header_RCF_Change_list-自訂", self.New_header_RCF_Change_list)
            self.New_RCF_TK_list = [[self.R_aK_TK_edit.text()]]
            print("self.New_RCF_TK_list-自訂", self.New_RCF_TK_list)
        if self.G_aK_mode.currentText() == "模擬":
            self.New_header_GCF_Change_list = [[header] * len(self.GCF_TK_list) for header in self.header_GCF_Change_list]
            print("self.New_header_GCF_Change_list", self.New_header_GCF_Change_list)
            self.New_GCF_TK_list = [list(self.GCF_TK_list) for _ in self.header_GCF_Change_list]
            print("self.New_GCF_TK_list", self.New_GCF_TK_list)
        elif self.G_aK_mode.currentText() == "自訂":
            self.New_header_GCF_Change_list = [[self.G_aK_box.currentText()]]
            print("self.New_header_GCF_Change_list-自訂", self.New_header_GCF_Change_list)
            self.New_GCF_TK_list = [[self.G_aK_TK_edit.text()]]
            print("self.New_GCF_TK_list-自訂", self.New_GCF_TK_list)
        if self.B_aK_mode.currentText() == "模擬":
            self.New_header_BCF_Change_list = [[header] * len(self.BCF_TK_list) for header in self.header_BCF_Change_list]
            print("self.New_header_BCF_Change_list", self.New_header_BCF_Change_list)
            self.New_BCF_TK_list = [list(self.BCF_TK_list) for _ in self.header_BCF_Change_list]
            print("self.New_BCF_TK_list", self.New_BCF_TK_list)
        elif self.B_aK_mode.currentText() == "自訂":
            self.New_header_BCF_Change_list = [[self.B_aK_box.currentText()]]
            print("self.New_header_BCF_Change_list-自訂", self.New_header_BCF_Change_list)
            self.New_BCF_TK_list = [[self.B_aK_TK_edit.text()]]
            print("self.New_BCF_TK_list-自訂", self.New_BCF_TK_list)

        # 找出Header組合list
        self.RGB_Header_list = []

        # 遍历 R 的所有子列表
        for r_sublist in self.New_header_RCF_Change_list:
            # 遍历 G 的所有子列表
            for g_sublist in self.New_header_GCF_Change_list:
                # 遍历 B 的所有子列表
                for b_sublist in self.New_header_BCF_Change_list:
                    # 遍历 R 子列表中的每个元素
                    for r_value in r_sublist:
                        # 遍历 G 子列表中的每个元素
                        for g_value in g_sublist:
                            # 遍历 B 子列表中的每个元素
                            for b_value in b_sublist:
                                # 将選取的值組合成一個新列表
                                combination = [r_value, g_value, b_value]
                                self.RGB_Header_list.append(combination)
        print("self.RGB_Header_list",self.RGB_Header_list)

        # 找出TK組合
        self.RGB_TK_list = []

        # 遍历 R 的所有子列表
        for r_sublist in self.New_RCF_TK_list:
            # 遍历 G 的所有子列表
            for g_sublist in self.New_GCF_TK_list:
                # 遍历 B 的所有子列表
                for b_sublist in self.New_BCF_TK_list:
                    # 遍历 R 子列表中的每个元素
                    for r_value in r_sublist:
                        # 遍历 G 子列表中的每个元素
                        for g_value in g_sublist:
                            # 遍历 B 子列表中的每个元素
                            for b_value in b_sublist:
                                # 将選取的值組合成一個新列表
                                combination = [r_value, g_value, b_value]
                                self.RGB_TK_list.append(combination)
        print("self.RGB_TK_list", self.RGB_TK_list)

        # 找出wave組合
        self.RGB_Wave_list = []

        # 遍历 R 的所有子列表
        for r_sublist in self.R_Change_wave_sim:
            # 遍历 G 的所有子列表
            for g_sublist in self.G_Change_wave_sim:
                # 遍历 B 的所有子列表
                for b_sublist in self.B_Change_wave_sim:
                    # 遍历 R 子列表中的每个元素
                    for r_value in r_sublist:
                        # 遍历 G 子列表中的每个元素
                        for g_value in g_sublist:
                            # 遍历 B 子列表中的每个元素
                            for b_value in b_sublist:
                                # 将選取的值組合成一個新列表
                                combination = [r_value, g_value, b_value]
                                self.RGB_Wave_list.append(combination)
        print("self.RGB_Wave_list", self.RGB_Wave_list)

        # 找出Purity組合
        self.RGB_Purity_list = []
        # 遍历 R 的所有子列表
        for r_sublist in self.R_Change_purity_sim:
            # 遍历 G 的所有子列表
            for g_sublist in self.G_Change_purity_sim:
                # 遍历 B 的所有子列表
                for b_sublist in self.B_Change_purity_sim:
                    # 遍历 R 子列表中的每个元素
                    for r_value in r_sublist:
                        # 遍历 G 子列表中的每个元素
                        for g_value in g_sublist:
                            # 遍历 B 子列表中的每个元素
                            for b_value in b_sublist:
                                # 将選取的值組合成一個新列表
                                combination = [r_value, g_value, b_value]
                                self.RGB_Purity_list.append(combination)
        print("self.RGB_Purity_list", self.RGB_Purity_list)

        # NTSC list
        self.RGB_NTSC_Change_list = []
        for i in range(len(self.RGB_x_change_sim_list)):
            R_x = self.RGB_x_change_sim_list[i][0]
            R_y = self.RGB_y_change_sim_list[i][0]
            G_x = self.RGB_x_change_sim_list[i][1]
            G_y = self.RGB_y_change_sim_list[i][1]
            B_x = self.RGB_x_change_sim_list[i][2]
            B_y = self.RGB_y_change_sim_list[i][2]
            self.RGB_NTSC_Change_list.append(100 * 0.5 * abs((R_x * G_y + G_x * B_y + B_x * R_y)-
                                                         (G_x*R_y)-(B_x*G_y)-(R_x*B_y))/0.1582)

        print("self.RGB_NTSC_Change_list",self.RGB_NTSC_Change_list)
        # 重新設置row count
        self.color_sim_table.setRowCount(len(self.W_x_Change_sum_list)+1)
        print("len(self.W_x_Change_sum_list)",len(self.W_x_Change_sum_list))
        # 開始放入table content隨便拿一個當數量迴圈
        rows_to_remove = set()  # 初始化集合記錄需要移除的行索引
        for i in range(len(self.W_x_Change_sum_list)):
            # W
            self.color_sim_table.setItem(i + 1, 1, QTableWidgetItem(f"{self.W_x_Change_sum_list[i]:.4f}"))
            self.color_sim_table.setItem(i + 1, 2, QTableWidgetItem(f"{self.W_y_Change_sum_list[i]:.4f}"))
            self.color_sim_table.setItem(i + 1, 3, QTableWidgetItem(f"{(self.RGB_T_change_sim_list[i][0]+self.RGB_T_change_sim_list[i][1] +self.RGB_T_change_sim_list[i][2])/3:.4f}"))
            # R
            self.color_sim_table.setItem(i + 1, 4, QTableWidgetItem(f"{self.RGB_x_change_sim_list[i][0]:.4f}"))
            self.color_sim_table.setItem(i + 1, 5, QTableWidgetItem(f"{self.RGB_y_change_sim_list[i][0]:.4f}"))
            self.color_sim_table.setItem(i + 1, 6, QTableWidgetItem(f"{self.RGB_T_change_sim_list[i][0]:.4f}"))
            self.color_sim_table.setItem(i + 1, 7, QTableWidgetItem(f"{self.RGB_Wave_list[i][0]:.1f}"))
            self.color_sim_table.setItem(i + 1, 8, QTableWidgetItem(f"{self.RGB_Purity_list[i][0]:.2f}"))
            # G
            self.color_sim_table.setItem(i + 1, 9, QTableWidgetItem(f"{self.RGB_x_change_sim_list[i][1]:.4f}"))
            self.color_sim_table.setItem(i + 1, 10, QTableWidgetItem(f"{self.RGB_y_change_sim_list[i][1]:.4f}"))
            self.color_sim_table.setItem(i + 1, 11, QTableWidgetItem(f"{self.RGB_T_change_sim_list[i][1]:.4f}"))
            self.color_sim_table.setItem(i + 1, 12, QTableWidgetItem(f"{self.RGB_Wave_list[i][1]:.1f}"))
            self.color_sim_table.setItem(i + 1, 13, QTableWidgetItem(f"{self.RGB_Purity_list[i][1]:.2f}"))
            # B
            self.color_sim_table.setItem(i + 1, 14, QTableWidgetItem(f"{self.RGB_x_change_sim_list[i][2]:.4f}"))
            self.color_sim_table.setItem(i + 1, 15, QTableWidgetItem(f"{self.RGB_y_change_sim_list[i][2]:.4f}"))
            self.color_sim_table.setItem(i + 1, 16, QTableWidgetItem(f"{self.RGB_T_change_sim_list[i][2]:.4f}"))
            self.color_sim_table.setItem(i + 1, 17, QTableWidgetItem(f"{self.RGB_Wave_list[i][2]:.1f}"))
            self.color_sim_table.setItem(i + 1, 18, QTableWidgetItem(f"{self.RGB_Purity_list[i][2]:.2f}"))

            # 色阻選擇
            self.color_sim_table.setItem(i + 1, 22, QTableWidgetItem(f"{self.RGB_Header_list[i][0]}"))
            self.color_sim_table.setItem(i + 1, 24, QTableWidgetItem(f"{self.RGB_Header_list[i][1]}"))
            self.color_sim_table.setItem(i + 1, 26, QTableWidgetItem(f"{self.RGB_Header_list[i][2]}"))

            # 厚度
            self.color_sim_table.setItem(i + 1, 23, QTableWidgetItem(f"{self.RGB_TK_list[i][0]:}"))
            print("self.RGB_TK_list[i][0]-check",self.RGB_TK_list[i][0])
            self.color_sim_table.setItem(i + 1, 25, QTableWidgetItem(f"{self.RGB_TK_list[i][1]:}"))
            print("i",i)
            print("self.RGB_TK_list[i][1]-check", self.RGB_TK_list[i][1])
            self.color_sim_table.setItem(i + 1, 27, QTableWidgetItem(f"{self.RGB_TK_list[i][2]:}"))

            # NTSC
            self.color_sim_table.setItem(i + 1, 19, QTableWidgetItem(f"{self.RGB_NTSC_Change_list[i]:.3f}"))

            # BLU
            self.color_sim_table.setItem(i + 1, 20, QTableWidgetItem(f"{self.BLU_x:.3f}"))
            self.color_sim_table.setItem(i + 1, 21, QTableWidgetItem(f"{self.BLU_y:.3f}"))

            self.color_sim_table.setItem(i + 1, 28, QTableWidgetItem(f"{self.CS_light_source.currentText()}"))
            if self.CS_layer1_mode.currentText() == "layer1_自訂":
                self.color_sim_table.setItem(i + 1, 29, QTableWidgetItem(f"{self.CS_layer1_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 29, QTableWidgetItem(""))
            if self.CS_layer2_mode.currentText() == "layer2_自訂":
                self.color_sim_table.setItem(i + 1, 30, QTableWidgetItem(f"{self.CS_layer2_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 30, QTableWidgetItem(""))
            if self.CS_layer3_mode.currentText() == "layer3_自訂":
                self.color_sim_table.setItem(i + 1, 31, QTableWidgetItem(f"{self.CS_layer3_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 31, QTableWidgetItem(""))
            if self.CS_layer4_mode.currentText() == "layer4_自訂":
                self.color_sim_table.setItem(i + 1, 32, QTableWidgetItem(f"{self.CS_layer4_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 32, QTableWidgetItem(""))
            if self.CS_layer5_mode.currentText() == "layer5_自訂":
                self.color_sim_table.setItem(i + 1, 33, QTableWidgetItem(f"{self.CS_layer5_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 33, QTableWidgetItem(""))
            if self.CS_layer6_mode.currentText() == "layer6_自訂":
                self.color_sim_table.setItem(i + 1, 34, QTableWidgetItem(f"{self.CS_layer6_box.currentText()}"))
            else:
                self.color_sim_table.setItem(i + 1, 34, QTableWidgetItem(""))

            should_remove_row = False  # 初始化是否移除當前行的標記
            # 開始篩選
            # W_x
            # 檢查條件並決定是否標記當前行為移除
            if self.W_checkbox.isChecked():
                print("W_x_checkbox_check")
                # 確保兩個條件都不為空
                if self.Wx_edit.text() and self.W_tolerance_edit.text():
                    # 現在只有當兩個條件都不為空，且值不在指定範圍內時，才標記當前行為移除
                    if not (float(self.W_x_Change_sum_list[i]) >= (
                            float(self.Wx_edit.text()) - float(self.W_tolerance_edit.text())) and
                            float(self.W_x_Change_sum_list[i]) <= (
                                    float(self.Wx_edit.text()) + float(self.W_tolerance_edit.text()))):
                        print("W_x_out_of_range")
                        should_remove_row = True

            # W_y
            if self.W_checkbox.isChecked():
                print("W_y_checkbox_check")
                # 確保兩個條件都不為空
                if self.Wy_edit.text() and self.W_tolerance_edit.text():
                    # 現在只有當兩個條件都不為空，且值不在指定範圍內時，才標記當前行為移除
                    if not (float(self.W_y_Change_sum_list[i]) >= (
                            float(self.Wy_edit.text()) - float(self.W_tolerance_edit.text())) and
                            float(self.W_y_Change_sum_list[i]) <= (
                                    float(self.Wy_edit.text()) + float(self.W_tolerance_edit.text()))):
                        print("W_y_out_of_range")
                        should_remove_row = True

            # R_x
            if self.R_checkbox.isChecked():
                print("R_checkbox_check")
                # 檢查是否為空字符串
                if self.Rx_edit.text() and self.R_tolerance_edit.text():
                    if not (float(self.RGB_x_change_sim_list[i][0]) >= (
                            float(self.Rx_edit.text()) - float(self.R_tolerance_edit.text()))
                            and float(self.RGB_x_change_sim_list[i][0]) <= (
                                    float(self.Rx_edit.text()) + float(self.R_tolerance_edit.text()))):
                        should_remove_row = True


            # R_y
            if self.R_checkbox.isChecked():
                print("R_checkbox_check")
                # 檢查是否為空字符串
                if self.Ry_edit.text() and self.R_tolerance_edit.text():
                    if not (float(self.RGB_y_change_sim_list[i][0]) >= (
                            float(self.Ry_edit.text()) - float(self.R_tolerance_edit.text()))
                    and float(self.RGB_y_change_sim_list[i][0]) <= (
                            float(self.Ry_edit.text()) + float(self.R_tolerance_edit.text()))):
                        should_remove_row = True

            # G_x
            if self.G_checkbox.isChecked():
                print("G_checkbox_check")
                # 檢查是否為空字符串
                if self.Gx_edit.text() and self.G_tolerance_edit.text():
                    if not (float(self.RGB_x_change_sim_list[i][1]) >= (
                            float(self.Gx_edit.text()) - float(self.G_tolerance_edit.text()))
                            and float(self.RGB_x_change_sim_list[i][1]) <= (
                                    float(self.Gx_edit.text()) + float(self.G_tolerance_edit.text()))):
                        should_remove_row = True

            # G_y
            if self.G_checkbox.isChecked():
                print("G_checkbox_check")
                # 檢查是否為空字符串
                if self.Gy_edit.text() and self.G_tolerance_edit.text():
                    if not (float(self.RGB_y_change_sim_list[i][1]) >= (
                            float(self.Gy_edit.text()) - float(self.G_tolerance_edit.text()))
                            and float(self.RGB_y_change_sim_list[i][1]) <= (
                                    float(self.Gy_edit.text()) + float(self.G_tolerance_edit.text()))):
                        should_remove_row = True

            # B_x
            if self.B_checkbox.isChecked():
                print("B_checkbox_check")
                # 檢查是否為空字符串
                if self.Bx_edit.text() and self.B_tolerance_edit.text():
                    if not (float(self.RGB_x_change_sim_list[i][2]) >= (
                            float(self.Bx_edit.text()) - float(self.B_tolerance_edit.text()))
                            and float(self.RGB_x_change_sim_list[i][2]) <= (
                                    float(self.Bx_edit.text()) + float(self.B_tolerance_edit.text()))):
                        should_remove_row = True

            # B_y
            if self.B_checkbox.isChecked():
                print("B_checkbox_check")
                # 檢查是否為空字符串
                if self.By_edit.text() and self.B_tolerance_edit.text():
                    if not (float(self.RGB_y_change_sim_list[i][2]) >= (
                            float(self.By_edit.text()) - float(self.B_tolerance_edit.text()))
                            and float(self.RGB_y_change_sim_list[i][2]) <= (
                                    float(self.By_edit.text()) + float(self.B_tolerance_edit.text()))):
                        should_remove_row = True

            # R_Wave & purity
            if self.R_wave_checkbox.isChecked():
                print("R_wave_checkbox_check")
                # 檢查是否為空字符串
                if not (self.R_wave_start_edit.text() and self.R_wave_end_edit.text() and
                     float(self.RGB_Wave_list[i][0]) >= (float(self.R_wave_start_edit.text())) and float(
                            self.RGB_Wave_list[i][0]) <= (float(self.R_wave_end_edit.text()))):
                    should_remove_row = True

            if self.R_purity_checkbox.isChecked():
                print("R_purity_checkbox_check")
                # 檢查是否為空字符串
                if not (self.R_purity_limit_edit.text() and
                    float(self.RGB_Purity_list[i][0]) >= (float(self.R_purity_limit_edit.text()))):
                    should_remove_row = True


            # G_Wave & purity
            if self.G_wave_checkbox.isChecked():
                print("G_wave_checkbox_check")
                # 檢查是否為空字符串
                if not (self.G_wave_start_edit.text() and self.G_wave_end_edit.text() and
                        float(self.RGB_Wave_list[i][1]) >= (float(self.G_wave_start_edit.text())) and float(
                            self.RGB_Wave_list[i][1]) <= (float(self.G_wave_end_edit.text()))):
                    should_remove_row = True

            if self.G_purity_checkbox.isChecked():
                print("G_purity_checkbox_check")
                # 檢查是否為空字符串
                if not (self.G_purity_limit_edit.text() and
                        float(self.RGB_Purity_list[i][1]) >= (float(self.G_purity_limit_edit.text()))):
                    should_remove_row = True

            # B_Wave & purity
            if self.B_wave_checkbox.isChecked():
                print("B_wave_checkbox_check")
                # 檢查是否為空字符串
                if not (self.B_wave_start_edit.text() and self.B_wave_end_edit.text() and
                        float(self.RGB_Wave_list[i][2]) >= (float(self.B_wave_start_edit.text())) and float(
                            self.RGB_Wave_list[i][2]) <= (float(self.B_wave_end_edit.text()))):
                    should_remove_row = True

            if self.B_purity_checkbox.isChecked():
                print("B_purity_checkbox_check")
                # 檢查是否為空字符串
                if not (self.B_purity_limit_edit.text() and
                        float(self.RGB_Purity_list[i][2]) >= (float(self.B_purity_limit_edit.text()))):
                    should_remove_row = True

            # NTSC
            if self.NTSC_check.isChecked():
                print("NTSC_check.isChecked")
                # 檢查是否為空字符串
                if not (self.NTSC_check_edit.text() and
                    float(self.RGB_NTSC_Change_list[i]) >= (float(self.NTSC_check_edit.text()))):
                    should_remove_row = True


            # TK
            # R=G=B
            if self.RGB_1_checkbox.isChecked():
                print("RGB_1_checkbox.isChecked")
                if not (float(self.RGB_TK_list[i][0]) == (float(self.RGB_TK_list[i][1])) and float(self.RGB_TK_list[i][1]) ==(float(self.RGB_TK_list[i][2]))):
                    should_remove_row = True

            # R=G>B
            if self.RGB_2_checkbox.isChecked():
                print("RGB_2_checkbox.isChecked")
                if not(float(self.RGB_TK_list[i][0]) == (float(self.RGB_TK_list[i][1])) and float(
                        self.RGB_TK_list[i][1]) > (float(self.RGB_TK_list[i][2]))):
                    should_remove_row = True

            # R=G>B
            if self.RGB_3_checkbox.isChecked():
                print("RGB_3_checkbox.isChecked")
                if not (float(self.RGB_TK_list[i][0]) > (float(self.RGB_TK_list[i][1])) and float(
                        self.RGB_TK_list[i][1]) == (float(self.RGB_TK_list[i][2]))):
                    should_remove_row = True
            # R>G>B
            if self.RGB_4_checkbox.isChecked():
                print("RGB_4_checkbox.isChecked")
                if not(float(self.RGB_TK_list[i][0]) > (float(self.RGB_TK_list[i][1])) and float(
                        self.RGB_TK_list[i][1]) > (float(self.RGB_TK_list[i][2]))):
                    should_remove_row = True

            # R<G=B
            if self.RGB_5_checkbox.isChecked():
                print("RGB_5_checkbox.isChecked")
                if not (float(self.RGB_TK_list[i][0]) < (float(self.RGB_TK_list[i][1])) and float(
                        self.RGB_TK_list[i][1]) == (float(self.RGB_TK_list[i][2]))):
                    should_remove_row = True

            # R<G<B
            if self.RGB_6_checkbox.isChecked():
                print("RGB_6_checkbox.isChecked")
                if not(float(self.RGB_TK_list[i][0]) < (float(self.RGB_TK_list[i][1])) and float(
                        self.RGB_TK_list[i][1]) < (float(self.RGB_TK_list[i][2]))):
                    should_remove_row = True

            # R=G<B
            if self.RGB_7_checkbox.isChecked():
                print("RGB_7_checkbox.isChecked")
                if not(float(self.RGB_TK_list[i][0]) == (float(self.RGB_TK_list[i][1])) and float(
                        self.RGB_TK_list[i][1]) < (float(self.RGB_TK_list[i][2]))):
                    should_remove_row = True

            # Free RTK
            if self.R_TK_check.isChecked():
                print("R_TK_check")
                # 檢查是否為空字符串
                if not(self.R_TK_start_edit.text() and self.R_TK_end_edit.text() and
                    float(self.RGB_TK_list[i][0]) >= (float(self.R_TK_start_edit.text())) and float(
                            self.RGB_TK_list[i][0]) <= (float(self.R_TK_end_edit.text()))):
                    print("i-pass",i)
                    print("self.RGB_TK_list[i][0]",self.RGB_TK_list[i][0])
                    should_remove_row = True

            # Free GTK
            if self.G_TK_check.isChecked():
                print("G_TK_check")
                # 檢查是否為空字符串
                if not(self.G_TK_start_edit.text() and self.G_TK_end_edit.text() and
                    float(self.RGB_TK_list[i][1]) >= (float(self.G_TK_start_edit.text())) and float(
                            self.RGB_TK_list[i][1]) <= (float(self.G_TK_end_edit.text()))):
                    should_remove_row = True

            # Free BTK
            if self.B_TK_check.isChecked():
                print("B_TK_check")
                # 檢查是否為空字符串
                if not(self.B_TK_start_edit.text() and self.B_TK_end_edit.text() and
                    float(self.RGB_TK_list[i][2]) >= (float(self.B_TK_start_edit.text())) and float(
                            self.RGB_TK_list[i][2]) <= (float(self.B_TK_end_edit.text()))):
                    should_remove_row = True
            # 如果當前行需要被移除，將其索引添加到移除集合中
            if should_remove_row:
                rows_to_remove.add(i + 1)
        # 移除標記為移除的行
        for row_index in sorted(rows_to_remove, reverse=True):
            self.color_sim_table.removeRow(row_index)

        # 重新設置剩餘行的行號
        for i in range(self.color_sim_table.rowCount()):
            self.color_sim_table.setItem(i, 0, QTableWidgetItem(str(i)))
        # 設定默認值
        column1_default_values = ["項目", "Wx", "Wy", "WY", "Rx", "Ry", "RY", "Rλ", "R_Purity", "Gx", "Gy", "GY",
                                  "Gλ", "G_Purity", "Bx", "By", "BY", "B_λ", "B_Purity", "NTSC%", "BLUx", "BLUy",
                                  "R色阻選擇", "R色阻厚度", "G色阻選擇", "G色阻厚度", "B色阻選擇", "B色阻厚度",
                                  "背光選擇", "Layer_1", "Layer_2", "Layer_3", "Layer_4",
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

        # 加載數據後自動調整所有欄位的寬度以適應內容
        self.color_sim_table.resizeColumnsToContents()






























