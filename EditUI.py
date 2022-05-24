from curses import savetty
from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtGui import QIcon
import GlobalManager as gm
import ctypes

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myappid")


class EditPanel:
    editing = []
    editname = ''
    editLessFirst = False
    editIndex = 0

    def __init__(self):
        self.window = uic.loadUi('qt5/EditPanel.ui')
        self.window.setWindowFlags(QtCore.Qt.WindowMinimizeButtonHint)
        self.window.setWindowIcon(QIcon("./finicon.ico"))

        self.window.creatButton.clicked.connect(lambda: self.CreatClip())
        self.window.deleteButton.clicked.connect(lambda: self.RemoveCLip())
        self.window.modeComboBox.currentIndexChanged.connect(
            lambda: self.SetMode())
        self.window.saveButton.clicked.connect(lambda: self.SaveRule())

    def SetEditing(self, rules, index):
        self.editIndex = index
        self.editing = rules[index]['content']
        self.editname = rules[index]['name']
        self.editLessFirst = rules[index]['lessFirst']
        self.window.contentTable.setRowCount(len(self.editing))
        j = 0
        for i in self.editing:
            self.window.contentTable.setItem(
                j, 0, QtWidgets.QTableWidgetItem(str(i[0])))
            self.window.contentTable.setItem(
                j, 1, QtWidgets.QTableWidgetItem(str(i[1])))
            j += 1
        self.window.ruleTitleLineEdit.setText(self.editname)
        if (self.editLessFirst):
            self.window.modeComboBox.setCurrentIndex(0)
        else:
            self.window.modeComboBox.setCurrentIndex(1)
        self.SetMode()

    def CreatClip(self):
        index = self.window.contentTable.rowCount()
        self.window.contentTable.setRowCount(index + 1)
        self.window.contentTable.setItem(index, 0,
                                         QtWidgets.QTableWidgetItem("0"))
        self.window.contentTable.setItem(index, 1,
                                         QtWidgets.QTableWidgetItem("A"))

    def RemoveCLip(self):
        for item in self.window.contentTable.selectedItems():
            selectIndex = item.row()
            self.window.contentTable.removeRow(selectIndex)
            del self.editing[selectIndex]

    def SetMode(self):
        mode = self.window.modeComboBox.currentIndex()
        if (mode == 0):
            self.window.contentTable.setHorizontalHeaderLabels(
                ['区间最小值', '对应等级'])
        else:
            self.window.contentTable.setHorizontalHeaderLabels(
                ['区间最大值', '对应等级'])

    def SaveRule(self):
        res = []
        for i in range(0, self.window.contentTable.rowCount()):
            res.append([
                float(self.window.contentTable.item(i, 0).text()),
                str(self.window.contentTable.item(i, 1).text())
            ])
        self.editing = res
        self.editname = self.window.ruleTitleLineEdit.text()
        mode = self.window.modeComboBox.currentIndex()
        if (mode == 0):
            self.editLessFirst = True
        else:
            self.editLessFirst = False
        sult = {
            'name': self.editname,
            'lessFirst': self.editLessFirst,
            'content': self.editing
        }
        gm.rulesPanel.ruleObj[self.editIndex] = sult
        gm.rulesPanel.SaveRules()
        gm.rulesPanel.ReloadRules()
        self.window.close()