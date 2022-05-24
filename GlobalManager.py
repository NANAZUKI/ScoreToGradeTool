global mainPanel
global rulesPanel
global ruleName
global editPanel
global app
global rule
rule = []
ruleName = []


def UpdateInfo():
    mainPanel.window.info_lable.setText(
        mainPanel.file.split('/')[-1] + ' - ' + ruleName)
