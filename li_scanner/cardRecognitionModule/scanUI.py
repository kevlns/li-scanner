# -*- coding: utf-8 -*-
import time

import numpy as np
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QImage
import os
import cv2
import xlwings as xw
from li_scanner.cardRecognitionModule.utils import getStuID, get_complete_card, getAnswers, SubjectiveSegmentation


class Ui_Form(QtWidgets.QWidget):
    def setupUi(self):
        self.setObjectName("Form")
        self.resize(955, 714)
        self.line = QtWidgets.QFrame(self)
        self.line.setGeometry(QtCore.QRect(373, 60, 20, 541))
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.msg2 = QtWidgets.QTextEdit(self)
        self.msg2.setGeometry(QtCore.QRect(460, 200, 431, 351))
        self.msg2.setObjectName("msg2")
        self.createScores = QtWidgets.QPushButton(self)
        self.createScores.setGeometry(QtCore.QRect(710, 70, 121, 28))
        self.createScores.setObjectName("createScores")
        self.chooseAnswer = QtWidgets.QPushButton(self)
        self.chooseAnswer.setGeometry(QtCore.QRect(510, 70, 121, 28))
        self.chooseAnswer.setObjectName("chooseAnswer")
        self.msg1 = QtWidgets.QLineEdit(self)
        self.msg1.setGeometry(QtCore.QRect(460, 140, 421, 21))
        self.msg1.setObjectName("msg1")
        self.camIdx = QtWidgets.QSpinBox(self)
        self.camIdx.setGeometry(QtCore.QRect(240, 80, 71, 22))
        self.camIdx.setObjectName("camIdx")
        self.label = QtWidgets.QLabel(self)
        self.label.setGeometry(QtCore.QRect(140, 80, 81, 16))
        self.label.setObjectName("label")

        # -------------------------------------------------------------------------
        self.label_show_camera = QtWidgets.QLabel()
        self.label_move = QtWidgets.QLabel()
        self.label_move.setFixedSize(100, 100)

        self.label_show_camera.setFixedSize(641, 481)
        self.label_show_camera.setAutoFillBackground(False)

        self.button_open_camera = QtWidgets.QPushButton(self)
        self.button_open_camera.setObjectName("button_open_camera")

        self.button_close = QtWidgets.QPushButton(u'退出')
        # 信息显示
        self.label_show_camera = QtWidgets.QLabel()
        self.label_show_camera.setGeometry(50, 120, 300, 500)
        self.label_move = QtWidgets.QLabel()
        self.label_move.setFixedSize(100, 100)

        self.timer_camera = QtCore.QTimer()  # 初始化定时器
        self.label_show_camera.setFixedSize(641, 481)
        self.label_show_camera.setAutoFillBackground(False)
        self.cap = cv2.VideoCapture()  # 初始化摄像头
        self.CAM_NUM = 0
        self.slot_init()
        self.__flag_work = 0
        self.x = 0
        self.count = 0




        # -------------------------------------------------------------------------

        self.retranslateUi(self)
        QtCore.QMetaObject.connectSlotsByName(self)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.createScores.setText(_translate("Form", "导出成绩单"))
        self.chooseAnswer.setText(_translate("Form", "选择标准答案"))
        self.label.setText(_translate("Form", "摄像头序号"))
        self.button_open_camera.setText(_translate("Form", "打开摄像头"))

    def slot_init(self):  # 建立通信连接
        self.button_open_camera.clicked.connect(self.button_open_camera_click)
        self.timer_camera.timeout.connect(self.show_camera)
        self.button_close.clicked.connect(self.close)

    def button_open_camera_click(self):
        if self.timer_camera.isActive() == False:
            flag = self.cap.open(self.CAM_NUM)
            if flag == False:
                msg = QtWidgets.QMessageBox.Warning(self, u'Warning', u'请检测相机与电脑是否连接正确',
                                                    buttons=QtWidgets.QMessageBox.Ok,
                                                    defaultButton=QtWidgets.QMessageBox.Ok)
                # if msg==QtGui.QMessageBox.Cancel:
                #                     pass
            else:
                self.timer_camera.start(30)
                self.button_open_camera.setText(u'关闭相机')
        else:
            self.timer_camera.stop()
            self.cap.release()
            self.label_show_camera.clear()
            self.button_open_camera.setText(u'打开相机')

    def show_camera(self):
        flag, self.image = self.cap.read()
        show = cv2.resize(self.image, (640, 480))
        show = cv2.cvtColor(show, cv2.COLOR_BGR2RGB)
        showImage = QtGui.QImage(show.data, show.shape[1], show.shape[0], QtGui.QImage.Format_RGB888)
        self.label_show_camera.setPixmap(QtGui.QPixmap.fromImage(showImage))

    def closeEvent(self, event):
        ok = QtWidgets.QPushButton()
        cancel = QtWidgets.QPushButton()
        msg = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning, u'关闭', u'是否关闭！')
        msg.addButton(ok, QtWidgets.QMessageBox.ActionRole)
        msg.addButton(cancel, QtWidgets.QMessageBox.RejectRole)
        ok.setText(u'确定')
        cancel.setText(u'取消')
        if msg.exec_() == QtWidgets.QMessageBox.RejectRole:
            event.ignore()
        else:
            if self.cap.isOpened():
                self.cap.release()
            if self.timer_camera.isActive():
                self.timer_camera.stop()
            event.accept()

import sys

pyapp = QtWidgets.QApplication(sys.argv)

ui = Ui_Form()
ui.setupUi()
ui.show()
sys.exit(pyapp.exec_())