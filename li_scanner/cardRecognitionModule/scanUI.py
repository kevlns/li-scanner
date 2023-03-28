import sys
import os
import cv2

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import QPalette, QBrush, QPixmap


class Ui_MainWindow(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super(Ui_MainWindow, self).__init__(parent)

        self.timer_camera = QtCore.QTimer()  # 初始化定时器
        self.cap = cv2.VideoCapture()  # 初始化摄像头
        self.CAM_NUM = 0
        self.set_ui()
        self.slot_init()
        self.__flag_work = 0
        self.x = 0
        self.count = 0

    def set_ui(self):
        #*****************************************************************************************
        # 分割线
        self.line = QtWidgets.QFrame(self)
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")

        self.line.setMinimumWidth(100)
        self.line.setMinimumHeight(600)

        # 单行输入区域
        self.msg1 = QtWidgets.QLineEdit(self)
        self.msg1.setObjectName("msg1")
        self.msg1.setMinimumWidth(400)
        # self.msg1.setMinimumHeight(50)
        # self.msg1.setMaximumWidth(400)
        # 多行输入区域
        self.msg2 = QtWidgets.QTextEdit(self)
        self.msg2.setObjectName("msg2")
        # 导出成绩单
        self.createScores = QtWidgets.QPushButton('导出成绩单')
        self.createScores.setObjectName("createScores")
        # 选择标准答案
        self.chooseAnswer = QtWidgets.QPushButton('选择标准答案')
        self.chooseAnswer.setObjectName("chooseAnswer")


        # 摄像头序号
        self.camIdx = QtWidgets.QSpinBox(self)
        self.camIdx.setObjectName("camIdx")
        # 摄像头序号标签
        self.label = QtWidgets.QLabel('摄像头序号')
        self.label.setObjectName("label")
        #*****************************************************************************************


        self.__layout_main = QtWidgets.QHBoxLayout()  # 采用QHBoxLayout类，按照从左到右的顺序来添加控件
        # self.__layout_fun_button = QtWidgets.QVBoxLayout()




        # self.__layout_data_show = QtWidgets.QVBoxLayout()  # QVBoxLayout类垂直地摆放小部件

        self.button_open_camera = QtWidgets.QPushButton(u'打开相机')


        # self.button_open_camera.setMinimumHeight(10)

        # move()方法是移动窗口在屏幕上的位置到x = 500，y = 500的位置上
        self.move(500, 200)

        # 信息显示
        self.label_show_camera = QtWidgets.QLabel()
        # self.label_show_camera.setFixedSize(641, 481)
        self.label_show_camera.setMinimumWidth(300)
        self.label_show_camera.setMaximumWidth(300)
        # self.label_show_camera.setAutoFillBackground(True)
        self.label_show_camera.setScaledContents (True)

        self.setWindowTitle(u'摄像头')

        # 一些布局
        self.left = QtWidgets.QVBoxLayout()
        self.left_top = QtWidgets.QHBoxLayout()
        self.left_top_right = QtWidgets.QHBoxLayout()
        self.left_top_left = QtWidgets.QHBoxLayout()
        self.right_top = QtWidgets.QHBoxLayout()
        self.right_bottom = QtWidgets.QVBoxLayout()
        self.right = QtWidgets.QVBoxLayout()

        self.left_top_right.addWidget(self.label)
        self.left_top_right.addWidget(self.camIdx)
        self.left_top_left.addWidget(self.button_open_camera)

        # self.left_top.addWidget(self.button_open_camera)
        self.left.addLayout(self.left_top_left)
        self.left.addLayout(self.left_top_right)


        self.left.addLayout(self.left_top)
        self.left.addWidget(self.label_show_camera)
        # self.left.setSizeConstraint()

        self.right_top.addWidget(self.chooseAnswer)
        self.right_top.addWidget(self.createScores)

        self.right_bottom.addWidget(self.msg1)
        self.right_bottom.addSpacing(30)
        self.right_bottom.addWidget(self.msg2)

        self.right.addLayout(self.right_top)
        self.right.addSpacing(30)
        self.right.addLayout(self.right_bottom)


        self.setLayout(self.__layout_main)
        self.__layout_main.addLayout(self.left)
        self.__layout_main.addWidget(self.line)
        self.__layout_main.addLayout(self.right)

    def slot_init(self):  # 建立通信连接
        self.button_open_camera.clicked.connect(self.button_open_camera_click)
        self.timer_camera.timeout.connect(self.show_camera)
        # self.button_close.clicked.connect(self.close)

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
        # print(show.shape)
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


if __name__ == '__main__':
    App = QApplication(sys.argv)
    win = Ui_MainWindow()
    win.show()
    sys.exit(App.exec_())

