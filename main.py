import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import Qt
from ui_courtstats1 import Ui_CourtStats  # your generated file

class CourtStatsApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_CourtStats()
        self.ui.setupUi(self)  # populate this QMainWindow

def main():

    # Enable high-DPI scaling before app initialization
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    window = CourtStatsApp()
    window.show()  
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()