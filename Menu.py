from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QPushButton, QLabel
from PySide6.QtGui import QIcon
from Color_Simulation import *
from Light_Source_BLU import *
from LED_Spectrum import Led_Spectrum
from CIE1931 import CIE_Spectrum
from Cell_total import Cell_Spectrum
from BSITO import BSITO_Spectrum
from RCF_Fix import RCF_Fix_Spectrum
from GCF_Fix import GCF_Fix_Spectrum
from BCF_Fix import BCF_Fix_Spectrum
from RCF_Change import RCF_Change_Spectrum
from GCF_Change import GCF_Change_Spectrum
from BCF_Change import BCF_Change_Spectrum

class Menu(QMainWindow):
    def __init__(self):
        super().__init__()

        # 創建主窗口
        self.setWindowTitle("Color_Simulation")
        self.setGeometry(100, 100, 1300, 900)

        # 创建一个选项卡窗口
        self.tab_widget = QTabWidget(self)

        # Page1
        self.page1 = Color_Simulation()
        self.tab_widget.addTab(self.page1, "Color_Simulation")


        # Page2
        self.page2 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page2, "Light_Source_BLU")

        # self.page2_table = Light_Source_BLU_Table()
        self.page2_table = Light_Source_BLU()

        self.page2_layout = QVBoxLayout()
        self.page2_layout.addWidget(self.page2_table)
        # self.page2_layout.addWidget(self.page2_table)

        # 將 page2_layout 設置為 page2 的佈局
        self.page2.setLayout(self.page2_layout)

        # Page3
        self.page3 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page3, "LED")
        self.page3_table = Led_Spectrum()

        self.page3_layout = QVBoxLayout()
        self.page3_layout.addWidget(self.page3_table)

        # 將 page3_layout 設置為 page3 的佈局
        self.page3.setLayout(self.page3_layout)

        # Page4
        self.page4 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page4, "CIE_Spectrum")
        self.page4_table = CIE_Spectrum()

        self.page4_layout = QVBoxLayout()
        self.page4_layout.addWidget(self.page4_table)

        # 將 page4_layout 設置為 page4 的佈局
        self.page4.setLayout(self.page4_layout)

        # Page5
        self.page5 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page5, "Cell")
        self.page5_table = Cell_Spectrum()

        self.page5_layout = QVBoxLayout()
        self.page5_layout.addWidget(self.page5_table)

        # 將 page5_layout 設置為 page5 的佈局
        self.page5.setLayout(self.page5_layout)

        # Page6
        self.page6 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page6, "BSITO")
        self.page6_table = BSITO_Spectrum()

        self.page6_layout = QVBoxLayout()
        self.page6_layout.addWidget(self.page6_table)

        # 將 page6_layout 設置為 page6 的佈局
        self.page6.setLayout(self.page6_layout)

        # Page7
        self.page7 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page7, "RCF_Fix")
        self.page7_table = RCF_Fix_Spectrum()

        self.page7_layout = QVBoxLayout()
        self.page7_layout.addWidget(self.page7_table)

        # 將 page7_layout 設置為 page7 的佈局
        self.page7.setLayout(self.page7_layout)

        # Page8
        self.page8 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page8, "GCF_Fix")
        self.page8_table = GCF_Fix_Spectrum()

        self.page8_layout = QVBoxLayout()
        self.page8_layout.addWidget(self.page8_table)

        # 將 page8_layout 設置為 page8 的佈局
        self.page8.setLayout(self.page8_layout)

        # Page9
        self.page9 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page9, "BCF_Fix")
        self.page9_table = BCF_Fix_Spectrum()

        self.page9_layout = QVBoxLayout()
        self.page9_layout.addWidget(self.page9_table)

        # 將 page9_layout 設置為 page9 的佈局
        self.page9.setLayout(self.page9_layout)

        # Page10
        self.page10 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page10, "RCF_Change")
        self.page10_table = RCF_Change_Spectrum()

        self.page10_layout = QVBoxLayout()
        self.page10_layout.addWidget(self.page10_table)

        # 將 page10_layout 設置為 page11 的佈局
        self.page10.setLayout(self.page10_layout)

        # Page11
        self.page11 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page11, "GCF_Change")
        self.page11_table = GCF_Change_Spectrum()

        self.page11_layout = QVBoxLayout()
        self.page11_layout.addWidget(self.page11_table)

        # 將 page11_layout 設置為 page11 的佈局
        self.page11.setLayout(self.page11_layout)

        # Page12
        self.page12 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page12, "BCF_Change")
        self.page12_table = BCF_Change_Spectrum()

        self.page12_layout = QVBoxLayout()
        self.page12_layout.addWidget(self.page12_table)

        # 將 page12_layout 設置為 page12 的佈局
        self.page12.setLayout(self.page12_layout)





        # 添加选项卡窗口到主窗口
        self.setCentralWidget(self.tab_widget)

        # Set Background
        self.setStyleSheet("background-color: #F5F5DC;")