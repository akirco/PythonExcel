import sys
import time
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *


class Stream(QObject):
    newText = pyqtSignal(str)

    def write(self, text):
        self.newText.emit(str(text))


class GenMast(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = '微赞数据筛选'
        self.width = 600
        self.height = 300
        self.icon = 'Excel.ico'

        self.initUI()
        sys.stdout = Stream(newText=self.onUpdateText)

    def onUpdateText(self, text):
        cursor = self.process.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertText(text)
        self.process.setTextCursor(cursor)
        self.process.ensureCursorVisible()

    def closeEvent(self, event):
        sys.stdout = sys.__stdout__
        super().closeEvent(event)

    def initUI(self):
        btnGenMast = QPushButton('Run', self)
        btnGenMast.move(450, 50)
        btnGenMast.resize(100, 200)
        btnGenMast.clicked.connect(self.genMastClicked)

        self.process = QTextEdit(self, readOnly=True)
        self.process.ensureCursorVisible()
        self.process.setLineWrapColumnOrWidth(500)
        self.process.setLineWrapMode(QTextEdit.FixedPixelWidth)
        self.process.setFixedHeight(200)
        self.process.setFixedWidth(400)
        self.process.move(30, 50)

        self.setWindowTitle(self.title)
        self.setGeometry(300, 300, self.width, self.height)
        self.setWindowIcon(QIcon(self.icon))
        self.show()

    def printHello(self):
        print('hello')

    def genMastClicked(self):
        print('Running...')
        self.printHello()
        loop = QEventLoop()
        QTimer.singleShot(2000, loop.quit())
        loop.exec_()
        print('done')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.aboutToQuit.connect(app.deleteLater)
    gui = GenMast()
    sys.exit(app.exec())
