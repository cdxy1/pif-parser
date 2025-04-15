import sys

from PyQt5.QtWidgets import QApplication

from src.pif_parser.app import ExcelMergerApp

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelMergerApp()
    window.show()
    sys.exit(app.exec_())
