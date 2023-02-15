from PyQt6.QtCore import pyqtSignal,QObject
from PyQt6.QtWidgets import QMessageBox
from enum import Enum

class Mysingal(QObject):
    setResult = pyqtSignal(str)
    setmessagebox = pyqtSignal(Enum,str,str)
    setprogressbar = pyqtSignal(int)

