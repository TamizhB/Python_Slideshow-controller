# -*- coding: utf-8 -*-
"""
Created on Wed Apr  4 15:11:15 2018

@author: Tamil SB
"""
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton,QFileDialog,QLabel
from PyQt5.QtGui import QIcon,QPixmap
#from PyQt5.QtCore import pyqtSlot
import ctrlexe
file=""
class App(QWidget):
 
    def __init__(self):
        super().__init__()
        self.title = 'ppt control'
        self.left = 50
        self.top = 50
        self.width = 275
        self.height = 70
        self.initUI()
 
    def openFileNameDialog(self):    
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","PowerPoint Files (*.pptx | *.ppt)", options=options)
        if fileName:
            global file
            file=fileName
            print(fileName)


        
    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        label = QLabel(self)
        pixmap = QPixmap('bg.jpg')
        label.setPixmap(pixmap)
        button = QPushButton('Start', self)
        button.setStyleSheet('QPushButton {background-color: #008000;color: #FFFFFF;}')
        button.setToolTip('Press to start!!!')
        button.move(20,30) 
        button.clicked.connect(self.on_click)
        button1 = QPushButton('Exit', self)
        button1.setStyleSheet('QPushButton {background-color: red;color: #FFFFFF;}')
        button1.setToolTip('Press to stop')
        button1.move(160,30) 
        button1.clicked.connect(self.on_click1)
 
        self.show()
 
    def on_click(self):
        self.openFileNameDialog()
        ctrlexe.run(file)

    def on_click1(self):
        sys.exit(app.exec_())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    exit(app.exec_())