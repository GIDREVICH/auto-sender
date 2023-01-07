from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow

import sys

def application():
    app = QApplication(sys.argv)
    main_window = QMainWindow

    main_window.setWindowTitle("Рассылка трек-номеров")
    main_window.setGeometry(300, 250, 350, 200)

    main_window.show()
    sys.exit(app.exec_())