import sys
from PySide6.QtWidgets import QApplication
from gui import MainWindow

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setApplicationName("出货数据合并工具")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
