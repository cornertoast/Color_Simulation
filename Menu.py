from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QPushButton, QLabel
from PySide6.QtGui import QIcon
from Color_Simulation import *
from Light_Source_BLU import *

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
        # Page2
        self.page2 = QWidget()  # 創建一個新的 QWidget 實例
        self.tab_widget.addTab(self.page2, "Light_Source_BLU")

        self.page2_table = Light_Source_BLU_Table()
        self.page2_button = Light_Source_BLU_button()

        self.page2_layout = QVBoxLayout()
        self.page2_layout.addWidget(self.page2_button)
        self.page2_layout.addWidget(self.page2_table)

        # 將 page2_layout 設置為 page2 的佈局
        self.page2.setLayout(self.page2_layout)

        # # Page3
        # self.page3 = Stock()
        # self.tab_widget.addTab(self.page3, "Stock")
        #
        # #Page4
        # self.page4 = Bank()
        # self.tab_widget.addTab(self.page4,"Bank")

        # 添加选项卡窗口到主窗口
        self.setCentralWidget(self.tab_widget)

        # Set Background
        self.setStyleSheet("background-color: #F5F5DC;")