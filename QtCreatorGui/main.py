import sys
import os

from PySide2.QtWidgets import QApplication, QWidget, QPushButton
from PySide2.QtCore import QFile
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import Slot

# pyuic5 -x .\form.ui -o output.py

class MyApp(QWidget):
    def __init__(self):
        super(MyApp, self).__init__()
        self.load_ui()

        self.window.buttonVerify.clicked.connect(self.verify)
        self.window.buttonLoadData.clicked.connect(self.loadData)

    def load_ui(self):
        loader = QUiLoader()
        path = os.path.join(os.path.dirname(__file__), "form.ui")
        ui_file = QFile(path)
        ui_file.open(QFile.ReadOnly)
        self.window = loader.load(ui_file, self)
        ui_file.close()

    def verify(self):
        print("Verifying...")

        serial = self.window.lineEditSerial.text()
        proNum = self.window.lineEditPro.text()
        chassis = self.window.lineEditChassis.text()

        print(serial, proNum, chassis)

    def loadData(self):
        serial = self.window.lineEditSerialStart.text()

        print("Loading data for {}".format(serial))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyApp()
    window.show()
    sys.exit(app.exec_())
