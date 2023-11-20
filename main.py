from PySide6.QtWidgets import QApplication
from Menu import Menu

import sys

#連結sys讓Gui內部起作用
app = QApplication(sys.argv)
app.setStyle("Fusion")

Menu_Window = Menu()
Menu_Window.show()

app.exec()
