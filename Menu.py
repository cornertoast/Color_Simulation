from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QPushButton, QLabel
from PySide6.QtGui import QIcon
from Color_Simulation import *
from Light_Source_BLU import *
from LED_Spectrum import Led_Spectrum
from CIE1931 import CIE_Spectrum
from Cell_total import Cell_Spectrum
from BSITO import BSITO_Spectrum

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
        self.tab_widget.addTab(self.page3, "LED_Spectrum")
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
        self.tab_widget.addTab(self.page5, "Cell_Spectrum")
        self.page5_table = Cell_Spectrum()

        self.page5_layout = QVBoxLayout()
        self.page5_layout.addWidget(self.page5_table)

        # 將 page5_layout 設置為 page5 的佈局
        self.page5.setLayout(self.page5_layout)

        # Page6
        self.page6 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page6, "BSITO_Spectrum")
        self.page6_table = BSITO_Spectrum()

        self.page6_layout = QVBoxLayout()
        self.page6_layout.addWidget(self.page6_table)

        # 將 page6_layout 設置為 page6 的佈局
        self.page6.setLayout(self.page6_layout)

        # 添加选项卡窗口到主窗口
        self.setCentralWidget(self.tab_widget)

        # Set Background
        self.setStyleSheet("background-color: #F5F5DC;")