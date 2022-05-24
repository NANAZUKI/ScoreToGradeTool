from PyQt5.QtWidgets import QMessageBox
import GlobalManager as gm


def ConvertTime(content):
    result = []
    for item in content:
        sult = str(item)
        sult = sult.replace('′', '.')
        sult = sult.replace('″', '')
        sult = sult.replace(':', '')
        sult = sult.replace('min', '.')
        sult = sult.replace('s', '')
        sult = sult.replace('h', '')
        try:
            sult = float(sult)
        except ValueError:
            result.append('NotNumber')
        else:
            result.append(sult)
    return result


def IsNum(str_number):
    str_number = str(str_number)
    if (str_number.split('.')[0]).isdigit() or str_number.isdigit() or (
            str_number.split('-')[-1]).split('.')[-1].isdigit():
        return True
    else:
        return False


def CaculateByRule(content, rule, autoConvert):
    if (autoConvert):
        digged = ConvertTime(content)
    else:
        digged = content

    for item in digged:
        try:
            dig = float(item)
        except ValueError:
            digged.append('NotNumber')
        else:
            digged.append(dig)
    result = []
    for item in digged:
        result.append(Cliping(item, rule['content'], rule['lessFirst'], 0))
    return result


def Cliping(item, rule, mode, index):
    if (IsNum(item)):
        if (not mode):
            if (index < len(rule) and item <= rule[index][0]):
                return str(rule[index][1])
            else:
                if index >= len(rule):
                    return '原始数据超出规则范围'
                else:
                    return Cliping(item, rule, mode, index + 1)
        else:
            if (index < len(rule) and item >= rule[index][0]):
                return str(rule[index][1])
            else:
                if index >= len(rule):
                    return '原始数据超出规则范围'
                else:
                    return Cliping(item, rule, mode, index + 1)
    else:
        return '原始数据不符合转换规范或不是数字'