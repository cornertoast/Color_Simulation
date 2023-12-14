from PySide6.QtCore import QObject, Signal

class SignalManager(QObject):
    # 定義一個信號
    databaseUpdated = Signal()

# 創建一個全局的信號管理器實例
global_signal_manager = SignalManager()
