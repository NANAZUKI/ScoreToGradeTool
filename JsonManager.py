import simplejson as json
import GlobalManager
from PyQt5.QtWidgets import QMessageBox


def ConvertNWrite(obj, path):
    info = str(json.dumps(obj))
    file = open(path, encoding='utf-8', mode='w')
    file.write(info)
    file.close


def ReadNConvert(path):
    file = open(path, encoding='utf-8')
    info = file.read()
    obj = json.loads(info)
    return obj