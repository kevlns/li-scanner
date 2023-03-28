import os
import sys

import cv2
import numpy as np
import xlwings as xw
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication

from li_scanner.cardRecognitionModule.SubjectUI import SubUI
from li_scanner.cardRecognitionModule.utils import getStuID, get_complete_card, getAnswers, SubjectiveSegmentation


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

        # ***************************************************************************************************
        # 每题选项个数
        self.divLines = []
        print(os.path.abspath('.'))
        f = open("../doc/divLines", "r")
        lines = f.readlines()
        for i in range(len(lines)):
            tmp = int(lines[i].strip('\n'))
            if tmp < 66:
                self.divLines.append(tmp)

        self.optNumOfSelQList = []
        f = open("../doc/optNumOfSelQList.txt", "r")
        lines = f.readlines()
        for i in range(len(lines)):
            self.optNumOfSelQList.append(int(lines[i].strip('\n')))

        # 获取每个题的坐标
        self.cors = []
        f1 = open("../doc/cors", "r")
        lines = f1.readlines()
        for i in range(len(lines)):
            tmp = lines[i].strip('\n').split(' ')
            self.cors.append([int(tmp[0]), int(tmp[1])])  # 删除\n

        self.idDigits = 0
        f1 = open("../doc/idDigits", "r")
        lines = f1.readlines()
        self.idDigits = int(lines[0][0])


        self.app = xw.App(visible=True, add_book=False)
        self.app.display_alerts = False  # 关闭一些提示信息，可以加快运行速度。 默认为 True。
        self.app.screen_updating = True  # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度

        self.wb = self.app.books.add()
        self.sht = self.wb.sheets[0]

        self.StuID = []
        self.cishu = 0

        self.scores = {}
        # ***************************************************************************************************

    def set_ui(self):
        # *****************************************************************************************
        # 分割线
        self.line = QtWidgets.QFrame(self)
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")

        self.line.setMinimumWidth(100)
        self.line.setMinimumHeight(640)

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
        # *****************************************************************************************

        self.__layout_main = QtWidgets.QHBoxLayout()  # 采用QHBoxLayout类，按照从左到右的顺序来添加控件

        # self.button_open_camera = QtWidgets.QPushButton(u'打开相机')

        # move()方法是移动窗口在屏幕上的位置到x = 500，y = 500的位置上
        self.move(500, 200)

        # 信息显示
        self.label_show_camera = QtWidgets.QLabel()
        self.label_show_camera.setMinimumWidth(480)
        self.label_show_camera.setMaximumWidth(480)
        self.label_show_camera.setScaledContents(True)

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
        # self.left_top_left.addWidget(self.button_open_camera)

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

        # 打开摄像头
        if self.timer_camera.isActive() == False:
            flag = self.cap.open(self.CAM_NUM)
            if flag == False:
                msg = QtWidgets.QMessageBox.Warning(self, u'Warning', u'请检测相机与电脑是否连接正确',
                                                    buttons=QtWidgets.QMessageBox.Ok,
                                                    defaultButton=QtWidgets.QMessageBox.Ok)
                # if msg==QtGui.QMessageBox.Cancel:
                #                     pass
                print(msg)
            else:
                self.timer_camera.start(30)

    def slot_init(self):  # 建立通信连接
        # self.button_open_camera.clicked.connect(self.button_open_camera_click)
        self.timer_camera.timeout.connect(self.show_camera)
        self.chooseAnswer.clicked.connect(self.chooseFile)
        self.createScores.clicked.connect(self.storedScores)
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
        frame = np.copy(self.image)
        self.image = np.rot90(self.image)
        show = cv2.cvtColor(self.image, cv2.COLOR_BGR2RGB)
        # print(show.shape)
        showImage = QtGui.QImage(show.data, show.shape[1], show.shape[0], QtGui.QImage.Format_RGB888)
        self.label_show_camera.setPixmap(QtGui.QPixmap.fromImage(showImage))
        score = cv2.Laplacian(frame, cv2.CV_64F).var()
        if score > 1000:
            # print('大于1000')
            totalCard = get_complete_card(frame)
            if totalCard is not None:
                tmp = getStuID(totalCard, self.idDigits)
                if self.StuID == tmp and tmp is not None:
                    self.cishu = self.cishu + 1
                else:
                    self.StuID = tmp
                    self.cishu = 0
                if self.cishu >= 3:
                    self.cishu = 0
                    sid = ''
                    for i in self.StuID:
                        sid = sid + str(i)

                    if sid in self.scores:
                        return
                    print('-' * 15, sid, '检测成功', '-' * 15)
                    # self.msg2.setText(sid+ '  检测成功')
                    self.msg2.append(sid + '  检测成功')
                    # print('optNumOfSelQList长度: ', optNumOfSelQList)
                    # 获取答案
                    answers = getAnswers(totalCard, self.optNumOfSelQList, self.cors)
                    # print(answers)
                    if answers is None:
                        return
                    # print(answers)
                    # for i in range(len(answers)):
                    #     print('第', i + 1, '题的答案是: ', answers[i])

                    # sht.range((1, 1), (1 ,1)).value = answers

                    pics = SubjectiveSegmentation(totalCard, self.divLines)
                    # print('函数之后')
                    if not os.path.exists(rf"../../record/subjective/{sid}"):
                        os.mkdir(rf"../../record/subjective/{sid}")
                    for i in range(len(pics)):
                        cv2.imwrite(f'../../record/subjective/{sid}/{sid}-{i}.jpg', pics[i])
                    answers_copy = []
                    # print(sid,'答案: ',answers)
                    for i in range(len(answers)):
                        tmp = answers[i]
                        st = ''
                        for j in tmp:
                            st = st + str(j)
                        answers_copy.append(st)
                    self.scores[sid] = answers_copy


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

    def storedScores(self):
        self.wb = self.app.books.add()
        self.sht = self.wb.sheets[0]
        self.sht.range((1, 1)).value = '考生号'
        self.sht.range((1, 2)).value = '分数'
        keys = list(self.scores.keys())
        # print(keys)
        # 存学号
        for i in range(len(keys)):
            self.sht.range((2 + i, 1)).value = str(keys[i])

        # 存题号
        for i in range(len(self.scores[keys[0]])):
            self.sht.range(1, i + 3).value = str(i + 1)
        # 学号分数字典
        # subScore = sub.total_score
        # print(subScore)
        # 存答案
        for j in range(len(keys)):
            # print('referenceAnswer',referenceAnswer)
            rightAnswers = 0
            for i in range(len(self.scores[keys[j]])):
                if self.scores[keys[j]][i] == self.referenceAnswer[i]:
                    # print('相等')
                    rightAnswers = rightAnswers + 1
            self.sht.range((j + 2, 2)).value = rightAnswers

            # print(scores[keys[j]])
            # 存选项
            self.sht.range((j + 2, 3), (j + 2, 3 + len(self.scores[keys[j]]))).value = self.scores[keys[j]]

        cv2.destroyAllWindows()
        self.wb.save(r'../../xlsx/scores.xlsx')
        self.wb.close()
        self.app.quit()
        # self.hide()
        self.close()
        # print(keys)
        # subApp = QApplication(sys.argv)
        # print(keys, len(self.divLines))
        self.sub = SubUI(keys, len(self.divLines))
        # sub = SubUI(["020200", "20000"], 2)
        self.sub.show()
        # sys.exit(subApp.exec_())

    def chooseFile(self, Filepath):
        # directory = QtWidgets.QFileDialog.getOpenFileName(self, "选取文件", "./", "All Files (*);;Text Files (*.txt)")
        directory = QtWidgets.QFileDialog.getOpenFileName(self, "选取文件", "./", "All Files (*);;Text Files (*.txt)")
        referenceAnswerPath = str(directory)[2:-19]

        # self.wb = self.app.books.open(r'../../xlsx/reference_answer.xlsx')  # 打开现有的工作簿
        if len(referenceAnswerPath) < 5:
            print('请选择excel文件格式')
            return
        if not referenceAnswerPath[-5:] == '.xlsx':
            print('请选择excel文件格式')
            return
        self.msg1.setText(referenceAnswerPath)
        self.wb = self.app.books.open(referenceAnswerPath)  # 打开现有的工作簿
        # print(self.wb)
        sht = self.wb.sheets[0]
        self.referenceAnswer = sht.range((2, 1), (2, len(self.optNumOfSelQList))).value

        self.wb.close()

        # print('referenceAnswer',self.referenceAnswer)
        for i in range(len(self.referenceAnswer)):
            if self.referenceAnswer[i] is not None:
                tmp = list(self.referenceAnswer[i])
            else:
                tmp = []
            strTmp = ''
            for j in range(len(tmp)):
                strTmp = strTmp + str(ord(tmp[j]) - 65 + 1)
            self.referenceAnswer[i] = strTmp
    # 当窗口非继承QtWidgets.QDialog时，self需替换成 None


App = QApplication(sys.argv)
win = Ui_MainWindow()
win.show()
sys.exit(App.exec_())
