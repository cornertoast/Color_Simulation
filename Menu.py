from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QPushButton, QLabel
from PySide6.QtGui import QIcon
from Color_Simulation import *
from Light_Source_BLU import *
from LED_Spectrum import *

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
        self.page2_BLU = Light_Source_BLU()

        self.page2_layout = QVBoxLayout()
        self.page2_layout.addWidget(self.page2_BLU)
        # self.page2_layout.addWidget(self.page2_table)

        # 將 page2_layout 設置為 page2 的佈局
        self.page2.setLayout(self.page2_layout)


        # Page3
        self.page3 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page3, "LED_Spectrum")

        self.page3_LED = Led_Spectrum()

        self.page3_layout = QVBoxLayout()
        self.page3_layout.addWidget(self.page3_LED)

        # 將 page3_layout 設置為 page3 的佈局
        self.page3.setLayout(self.page3_layout)

        # 添加选项卡窗口到主窗口
        self.setCentralWidget(self.tab_widget)

        # # Set Background
        # self.setStyleSheet("background-color: #F5F5DC;")