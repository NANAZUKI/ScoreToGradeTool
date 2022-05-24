import xlwings as xw


def ToA1(R1C1):
    return xw.apps[app.pid].api.ConvertFormula(R1C1, -4150, 1, 4)


def ToR1C1(A1):
    return xw.apps[app.pid].api.ConvertFormula(A1, 1, -4150, 1)


def CutRCRange(RC):
    beganPiece = RC.split(':')[0]
    endingPiece = RC.split(':')[1]
    bR = int(beganPiece.split('R')[1].split('C')[0])
    bC = int(beganPiece.split('C')[1])
    eR = int(endingPiece.split('R')[1].split('C')[0])
    eC = int(endingPiece.split('C')[1])
    ret = [[bR, bC], [eR, eC]]
    return ret


def OpenXlsx(path):
    global app
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    global wb
    wb = app.books.open(path)
    global sheet
    sheet = wb.sheets[0]
    global usedRange
    usedRange = sheet.used_range


def ReadXlsx():
    return sheet.range(usedRange).value


def WriteXlsx(content):
    sheet.range('A1').value = content
    wb.save()


def SnQ():
    wb.save()
    wb.close()
    app.quit()
