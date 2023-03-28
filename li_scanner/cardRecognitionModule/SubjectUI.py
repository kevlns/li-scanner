# 无越界检测
import sys

import cv2
from PyQt5.QtGui import QPixmap, QDoubleValidator
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow, QLabel, QHBoxLayout, QVBoxLayout, QLineEdit, \
    QPushButton, QComboBox


class ZhuguantiUI(QMainWindow):

    def __init__(self, StuID: list, Numofquestion: int):
        """
        :param StuID: 学号列表
        :param Numofquestion: 分割出的图片个数
        """
        super(ZhuguantiUI, self).__init__()
        self.StuID = StuID  # 学号列表
        self.scores = {}  # 储存每人每道题得分
        self.index_ID = 0   # 列表索引
        self.path = "../../record/subjective/"    # 题目文件夹存放路径
        self.input_score = QLineEdit()
        self.label_pic = QLabel(self)
        self.total_num = Numofquestion
        self.num = 0  # 第几道大题
        self.cb_num = QComboBox(self)
        self.total_score = {}
        for id in self.StuID:
            self.scores[id] = []
            self.total_score[id] = 0
            for nq in range(0, Numofquestion):
                self.scores[id].append(0)
        self.initUI()


    def initUI(self):
        # 设置窗口大小
        self.resize(800, 500)
        # 设置窗口标题
        self.setWindowTitle("轻阅-主观题批改")

        for index in range(0, self.total_num):
            self.cb_num.addItem(str(index + 1))
        hbox_cb = QHBoxLayout()
        hbox_cb.addWidget(self.cb_num)
        hbox_cb.addStretch(1)
        self.cb_num.currentIndexChanged.connect(self.SetNum)

        # 加载图片
        # aaa = cv2.imread(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.num}.jpg")
        # print(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.num}.jpg")
        # cv2.imshow("aaa",aaa)
        pixmap = QPixmap(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.num}.jpg")
        self.label_pic.setPixmap(pixmap)
        self.label_pic.setScaledContents(True)

        # 创建图片水平布局
        hbox_pic = QHBoxLayout()
        hbox_pic.addStretch(1)
        hbox_pic.addWidget(self.label_pic)
        hbox_pic.addStretch(1)

        # 创建打分及切换其他人答案水平布局
        self.input_score.setFixedWidth(50)
        self.input_score.setFixedHeight(40)
        # 校验器，只能输入两位整数和一位小数
        doubleValidator = QDoubleValidator(self)
        doubleValidator.setRange(0, 99)
        doubleValidator.setNotation(QDoubleValidator.StandardNotation)
        doubleValidator.setDecimals(1)
        self.input_score.setValidator(doubleValidator)
        lable_sco = QLabel("得分：")
        last_but = QPushButton("上一张")
        next_but = QPushButton("下一张")
        last_but.clicked.connect(self.lastbutClicked)
        next_but.clicked.connect(self.nextbutClicked)
        hbox_sco = QHBoxLayout()
        hbox_sco.addWidget(lable_sco)
        hbox_sco.addWidget(self.input_score)
        hbox_sco.addStretch(1)
        hbox_sco.addWidget(last_but)
        hbox_sco.addWidget(next_but)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(hbox_cb)
        vbox.addLayout(hbox_pic)
        vbox.addLayout(hbox_sco)
        vbox.addStretch(1)

        mainFrame = QWidget()
        mainFrame.setLayout(vbox)
        self.setCentralWidget(mainFrame)

        # cv.waitKey(0)
        # cv.destroyAllWindows()

    # def changePicture(self):
    #     print(1)
    #     self.label_pic.setPixmap(f"cut{self.num}.jpg")

    def SetNum(self):
        self.num = self.cb_num.currentIndex()
        pixmap = QPixmap(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.num}.jpg")
        self.label_pic.setPixmap(pixmap)
        self.label_pic.setScaledContents(True)

    def lastbutClicked(self):
        self.index_ID -= 1
        self.total_score[self.StuID[self.index_ID]] -= self.scores[self.StuID[self.index_ID]][self.num]
        if self.index_ID < 0:
            self.index_ID = 0
        else:
            self.input_score.setText('')    # 清空输入框
        pixmap = QPixmap(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.num}.jpg")
        self.label_pic.setPixmap(pixmap)
        self.label_pic.setScaledContents(True)

    def nextbutClicked(self):
        # print("~"*10)
        # print(self.scores[self.StuID[self.index_ID]][self.num])
        self.scores[self.StuID[self.index_ID]][self.num] = float(self.input_score.text())
        # print(1)
        self.total_score[self.StuID[self.index_ID]] += float(self.input_score.text())
        # print(2)
        self.index_ID += 1
        # print(3, self.index_ID)
        if self.index_ID >= len(self.StuID):
            # print(4)
            self.index_ID = len(self.StuID) - 1
        else:
            # print(5)
            self.input_score.setText('')    # 清空输入框
        pixmap = QPixmap(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.num}.jpg")
        self.label_pic.setPixmap(pixmap)
        self.label_pic.setScaledContents(True)
        # print(111)


# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#
#     main = ZhuguantiUI([123])
#     main.show()
#
#     sys.exit(app.exec_())
