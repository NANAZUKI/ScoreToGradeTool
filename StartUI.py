from PyQt5 import uic
import GlobalManager as gm
from PyQt5.QtGui import QIcon
import ctypes

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myappid")


class StartPanel:

    def __init__(self):
        self.window = uic.loadUi('qt5/StartPanel.ui')
        self.window.setWindowIcon(QIcon("./finicon.ico"))
        self.window.setFixedSize(self.window.width(), self.window.height())
        self.window.pushButton.clicked.connect(lambda: self.Start())

    def Start(self):
        gm.mainPanel.file = gm.mainPanel.OpenXlsx()
        gm.mainPanel.LoadXlsx()
        gm.mainPanel.window.show()
        gm.rulesPanel.Preparing()
        self.window.close()