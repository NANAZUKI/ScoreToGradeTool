from PyQt5 import uic, QtWidgets
from PyQt5.QtGui import QBrush, QColor
from PyQt5.QtGui import QIcon
import os, GlobalManager
import XlsxManager as xm
import Caculate
import ctypes

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myappid")


class MainPanel:
    file = ''
    scoreContent = []
    gradeItems = []

    def __init__(self):
        self.window = uic.loadUi('qt5/MainPanel.ui')
        self.window.setWindowIcon(QIcon("./finicon.ico"))
        self.window.setFixedSize(self.window.width(), self.window.height())
        self.window.sheetEdit.setRowCount(0)
        self.window.sheetEdit.setColumnCount(0)

        self.window.saveButton.clicked.connect(lambda: self.WriteXlsx())
        self.window.sheetEdit.cellDoubleClicked.connect(
            lambda: self.Purefresh())
        self.window.markScoreButton.clicked.connect(lambda: self.MarkScore())
        self.window.markGradeButton.clicked.connect(lambda: self.MarkGrade())
        self.window.rulesButton.clicked.connect(lambda: self.SelectRules())
        self.window.writeInButton.clicked.connect(lambda: self.WorkOut())

    #Import xls

    def OpenXlsx(self):
        fileName, fileType = QtWidgets.QFileDialog.getOpenFileName(
            self.window, '选取文件', os.getcwd(), 'Excel工作簿(*.xlsx)')
        self.file = fileName
        return fileName

    #Load xls to table
    def LoadXlsx(self):
        self.window.info_lable.setText(
            self.file.split('/')[-1] + ' - ' + '未选择规则')
        #Get content
        try:
            xm.OpenXlsx(self.file)
        except:
            msg = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Critical,
                                        "打开文件错误", "必须打开一个Excel工作簿！")
            msg.exec_()
            GlobalManager.app.quit()
        else:
            content = xm.ReadXlsx()
            #Convert range
            RCrange = xm.CutRCRange(xm.ToR1C1(xm.usedRange.address))
            #Creat table
            self.window.sheetEdit.setRowCount(RCrange[1][0])
            self.window.sheetEdit.setColumnCount(RCrange[1][1])
            #Write contents
            for r in range(0, RCrange[1][0]):
                for c in range(0, RCrange[1][1]):
                    if str(content[r][c]) == 'None':
                        content[r][c] = ''
                    self.window.sheetEdit.setItem(
                        r, c, QtWidgets.QTableWidgetItem(str(content[r][c])))
            xm.SnQ()
            #Set title
            name = self.file.split('/')[-1]
            self.window.setWindowTitle('Grading ' + name)
            self.window.xlsTitle_lable.setText(name)

    def WriteXlsx(self):
        #Generate content
        xm.OpenXlsx(self.file)
        content = []
        for r in range(0, self.window.sheetEdit.rowCount()):
            row = []
            for c in range(0, self.window.sheetEdit.columnCount()):
                row.append(str(self.window.sheetEdit.item(r, c).text()))
            content.append(row)
        xm.WriteXlsx(content)
        xm.SnQ()

    def Refresh(self, color):
        if (color == QBrush(QColor(255, 200, 200))):
            self.scoreContent = []
            self.window.scoreCount_label.setText('0')
        if (color == QBrush(QColor(200, 200, 255))):
            self.gradeItems = []
            self.window.gradeCount_label.setText('0')
        for r in range(0, self.window.sheetEdit.rowCount()):
            for c in range(0, self.window.sheetEdit.columnCount()):
                if (self.window.sheetEdit.item(r, c).background() == color):
                    self.window.sheetEdit.item(r, c).setBackground(
                        QBrush(QColor(255, 255, 255)))

    def MarkScore(self):
        self.Refresh(QBrush(QColor(255, 200, 200)))
        sItems = self.window.sheetEdit.selectedItems()
        i = 0
        for item in sItems:
            if item.background() != QBrush(QColor(200, 200, 255)):
                item.setBackground(QBrush(QColor(255, 200, 200)))
                self.scoreContent.append(item.text())
                i += 1

        self.window.scoreCount_label.setText(str(i))

    def MarkGrade(self):
        self.Refresh(QBrush(QColor(200, 200, 255)))
        gItems = self.window.sheetEdit.selectedItems()
        i = 0
        for item in gItems:
            if item.background() != QBrush(QColor(255, 200, 200)):
                item.setBackground(QBrush(QColor(200, 200, 255)))
                self.gradeItems.append(item)
                i += 1
        self.window.gradeCount_label.setText(str(i))

    def SelectRules(self):
        GlobalManager.rulesPanel.window.show()
        GlobalManager.rulesPanel.ReloadRules()

    def WorkOut(self):
        if (GlobalManager.rule == []):
            QtWidgets.QMessageBox.critical(self.window, '请选择规则', '没有选择规则！')
        else:
            if (self.gradeItems == [] or self.scoreContent == []):
                QtWidgets.QMessageBox.critical(self.window, '空区域',
                                               '没有选择分数区域或等级区域！')
            else:
                if (len(self.gradeItems) != len(self.scoreContent)):
                    QtWidgets.QMessageBox.warning(self.window, '提示',
                                                  '提示：分数区域和等级区域不匹配，可能导致转换不完全')
                sult = []
                sult = Caculate.CaculateByRule(
                    self.scoreContent, GlobalManager.rule,
                    self.window.convertCheckBox.isChecked())
                for i in range(0, len(self.gradeItems)):
                    if (i >= len(sult)):
                        s = ''
                    else:
                        s = sult[i]
                    self.gradeItems[i].setText(s)

    def Purefresh(self):
        for item in self.window.sheetEdit.selectedItems():
            self.Refresh(item.background())
