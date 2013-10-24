#-*- coding: utf-8 -*-
import sys
import Schedule
from PyQt4 import QtGui
from PyQt4 import QtCore

def main():
    app = QtGui.QApplication(sys.argv)
    ex =  Schedule.import_excel()
    ex.saveFile1.test()
    app.exec_()
    
if __name__ == '__main__':
    main()



