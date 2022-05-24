from PyQt5 import QtWidgets
import MainUI, RulesUI, StartUI, EditUI, GlobalManager

if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    GlobalManager.app = app

    mainWindow = MainUI.MainPanel()
    GlobalManager.mainPanel = mainWindow
    rulesWindow = RulesUI.RulesPanel()
    GlobalManager.rulesPanel = rulesWindow
    editWindow = EditUI.EditPanel()
    GlobalManager.editPanel = editWindow
    startWindow = StartUI.StartPanel()
    startWindow.window.show()

    app.exec_()