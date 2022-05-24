from PyQt5 import uic, QtWidgets, QtCore
from PyQt5.QtWidgets import QMessageBox, QFileIconProvider
from PyQt5.QtGui import QIcon
import os
import GlobalManager
import JsonManager as jm
import ctypes

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('myappid')


class RulesPanel:

    def __init__(self):
        self.window = uic.loadUi('qt5/RulesPanel.ui')
        self.window.setWindowIcon(QIcon('./finicon.ico'))
        self.window.setFixedSize(self.window.width(), self.window.height())

        self.window.importButton.clicked.connect((lambda: self.LoadRules()))
        self.window.creatButton.clicked.connect((lambda: self.CreatRule()))
        self.window.removeButton.clicked.connect((lambda: self.RemoveRule()))
        self.window.selectButton.clicked.connect((lambda: self.SaveNClose()))
        self.window.exportButton.clicked.connect((lambda: self.ExportRules()))
        self.window.editButton.clicked.connect((lambda: self.EditRule()))
        self.window.rulesTable.cellClicked.connect(
            (lambda: self.UpdateRuleLabel(self.ruleObj[
                self.window.rulesTable.selectedItems()[0].row()]['name'])))
        try:
            self.config = jm.ReadNConvert('config.json')
        except FileNotFoundError:
            QMessageBox.critical(GlobalManager.rulesPanel.window, '文件错误',
                                 '配置文件丢失！请重新安装软件！')
            GlobalManager.app.quit()
        else:
            try:
                self.memoryPath = str(self.config['rulesMemoryPath'])
            except KeyError:
                QMessageBox.critical(GlobalManager.rulesPanel.window, '文件错误',
                                     '配置已损坏！请重新安装软件！')

    def Preparing(self):
        self.UpdateRuleLabel('无')
        try:
            self.LoadRules(self.memoryPath)
        except FileNotFoundError:
            QMessageBox.information(self.window, '找不到规则方案', '规则方案文件被删除或更改！')
            self.LoadRules()

    # WARNING: Decompyle incomplete

    def LoadRules(self, *Path):
        try:
            self.window.rulesTable.setRowCount(0)
            if Path == ():
                fileName, ileType = QtWidgets.QFileDialog.getOpenFileName(
                    self.window, '选取规则方案文件', os.getcwd(), '规则方案文件(*.rulepack)')
            else:
                fileName = Path[0]
            self.ruleObj = jm.ReadNConvert(fileName)
            print(self.ruleObj)
            self.ViewRules()
        except FileNotFoundError:
            QMessageBox.critical(GlobalManager.rulesPanel.window, '找不到规则方案',
                                 '必须选择规则方案！')

    # WARNING: Decompyle incomplete

    def ReloadRules(self):
        self.LoadRules(self.config['rulesMemoryPath'])

    def ViewRules(self):
        try:
            self.window.rulesTable.setRowCount(len(self.ruleObj))
            i = 0
            for row in self.ruleObj:
                self.window.rulesTable.setItem(
                    i, 0, QtWidgets.QTableWidgetItem(row['name']))
                if (row['lessFirst']):
                    head = '区间最小值'
                    self.window.rulesTable.setItem(
                        i, 1, QtWidgets.QTableWidgetItem('较小值优先'))
                else:
                    head = '区间最大值'
                    self.window.rulesTable.setItem(
                        i, 1, QtWidgets.QTableWidgetItem('较大值优先'))
                overview = ''
                for e in row['content']:
                    overview = overview + head + "：" + str(
                        e[0]) + " 对应分数：" + str(e[1]) + "；"
                self.window.rulesTable.setItem(
                    i, 2, QtWidgets.QTableWidgetItem(overview))
                i += 1
        except KeyError:
            QMessageBox.critical(GlobalManager.rulesPanel.window, '文件错误',
                                 '规则方案不正确或已损坏，请联系规则方案提供者')

    def CreatRule(self):
        index = self.window.rulesTable.rowCount()
        self.window.rulesTable.setRowCount(index + 1)
        self.window.rulesTable.setItem(
            index, 0, QtWidgets.QTableWidgetItem('新规则' + str(index + 1)))
        self.window.rulesTable.setItem(
            index, 1, QtWidgets.QTableWidgetItem(str('较大值优先')))
        self.window.rulesTable.setItem(
            index, 2,
            QtWidgets.QTableWidgetItem(
                str('区间最大值：60 对应等级D； 区间最大值：80 对应等级B； 区间最大值：90 对应等级C； 区间最大值：100 对应等级A；'
                    )))
        self.ruleObj.append({
            'name':
            '新规则' + str(index + 1),
            'lessFirst':
            False,
            'content': [[60, 'D'], [80, 'B'], [90, 'C'], [100, 'A']]
        })

    def RemoveRule(self):
        if len(self.window.rulesTable.selectedItems()) != 0:
            selectIndex = self.window.rulesTable.selectedItems()[0].row()
            self.window.rulesTable.removeRow(selectIndex)
            del self.ruleObj[selectIndex]

    def SaveRules(self):
        try:
            jm.ConvertNWrite(self.ruleObj, str(self.config['rulesMemoryPath']))
            if len(self.window.rulesTable.selectedItems()) == 0:
                QMessageBox.information(self.window, '请选择规则', '没有选择规则！')
            else:
                GlobalManager.rule = self.ruleObj[
                    self.window.rulesTable.selectedItems()[0].row()]
                name = self.ruleObj[self.window.rulesTable.selectedItems()
                                    [0].row()]['name']
                GlobalManager.ruleName = name
                self.UpdateRuleLabel(name)
                GlobalManager.UpdateInfo()
        except KeyError:
            QMessageBox.critical(GlobalManager.rulesPanel.window, '文件错误',
                                 '规则方案不正确或已损坏，请联系规则方案提供者')

    def SaveNClose(self):
        self.SaveRules()
        self.window.close()

    def ExportRules(self):
        try:
            content = jm.ConvertNWrite(self.ruleObj,
                                       str(self.config['rulesMemoryPath']))
        except FileNotFoundError:
            QMessageBox.critical(GlobalManager.rulesPanel.window, '文件错误',
                                 '规则方案文件被删除或更改！')
        try:
            fileName, fileType = QtWidgets.QFileDialog.getSaveFileName(
                self.window, '选择保存路径', 'C:\\Users', '规则方案文件(*.rulepack)')
        except FileNotFoundError:
            QMessageBox.information(self.window, '导出失败', '导出位置不存在')
        else:
            if (fileName == ''):
                QMessageBox.information(self.window, '导出失败', '必须选择导出位置')
            else:
                try:
                    jm.ConvertNWrite(self.ruleObj, fileName)
                except (FileNotFoundError, FileExistsError):
                    QMessageBox.information(self.window, '导出失败', '导出位置不存在或被更改')
                else:
                    QMessageBox.information(self.window, '导出成功', '导出成功')

    # WARNING: Decompyle incomplete

    def EditRule(self):
        if len(self.window.rulesTable.selectedItems()) != 0:
            GlobalManager.editPanel.SetEditing(
                self.ruleObj,
                self.window.rulesTable.selectedItems()[0].row())
            GlobalManager.editPanel.window.show()

    def UpdateRuleLabel(self, name):
        self.window.ruleLable.setText(name)
