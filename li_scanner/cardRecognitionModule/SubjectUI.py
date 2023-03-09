import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QDoubleValidator
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow, QLabel, QHBoxLayout, QVBoxLayout, QLineEdit, \
    QPushButton, QComboBox
import xlwings


class ZhuguantiUI(QMainWindow):

    def __init__(self, StuID: list, Numofquestion: int):
        """
        :param StuID: 学号列表
        :param Numofquestion: 分割出的图片个数
        """
        super(ZhuguantiUI, self).__init__()
        self.StuID = StuID  # 学号列表
        self.scores = {}  # 储存每人每道题得分
        self.index_ID = 0  # 列表索引
        self.path = "../../record/subjective/"  # 题目文件夹存放路径
        self.input_score = QLineEdit()
        self.label_pic = QLabel(self)
        self.total_num = Numofquestion
        self.num = 0  # 第几道大题
        self.cb_num = QComboBox(self)  # 大题题号选择下拉框
        self.total_score = {}  # 每人大题总分
        for id in self.StuID:
            self.scores[id] = []
            self.total_score[id] = 0
            for nq in range(0, Numofquestion):
                self.scores[id].append(0)
        self.initUI()

    def initUI(self):
        # 设置窗口大小
        self.resize(800, 495)
        # 设置窗口标题
        self.setWindowTitle("轻阅-主观题批改")

        # 设置下拉框
        for index in range(0, self.total_num):
            self.cb_num.addItem("第" + str(index + 1) + "题")
        self.cb_num.setStyleSheet('font-size: 40px;')
        self.cb_num.currentIndexChanged.connect(self.SetNum)

        label_title = QLabel(self)
        label_title.setText("主观题批改")
        label_title.setStyleSheet('font-size: 50px; color: #24806A; font-weight: bold')

        button_export = QPushButton("导出成绩")

        button_start = QPushButton("进入")
        button_start.clicked.connect(self.correction)

        hbox_win1_top = QHBoxLayout()
        hbox_win1_top.addSpacing(25)
        hbox_win1_top.addWidget(label_title, 4)
        hbox_win1_top.addStretch(1)
        hbox_win1_top.addWidget(button_export, 1)
        hbox_win1_top.addSpacing(25)
        
        hbox_win1_middle = QHBoxLayout()
        hbox_win1_middle.addWidget(self.cb_num, 0, Qt.AlignJustify | Qt.AlignCenter)

        hbox_win1_bottom = QHBoxLayout()
        hbox_win1_bottom.addWidget(button_start, 0, Qt.AlignRight | Qt.AlignCenter)

        vbox_win1 = QVBoxLayout()
        vbox_win1.addSpacing(30)
        vbox_win1.addLayout(hbox_win1_top, 1)
        vbox_win1.addLayout(hbox_win1_middle, 3)
        vbox_win1.addLayout(hbox_win1_bottom, 1)
        vbox_win1.addSpacing(10)

        window_1 = QWidget()
        window_1.setLayout(vbox_win1)
        self.setCentralWidget(window_1)

    # def changePicture(self):
    #     print(1)
    #     self.label_pic.setPixmap(f"cut{self.num}.jpg")

    # 每道大题批改界面
    def correction(self):
        # 加载图片
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
        vbox.addLayout(hbox_pic)
        vbox.addLayout(hbox_sco)
        vbox.addStretch(1)

        mainFrame = QWidget()
        mainFrame.setLayout(vbox)
        self.setCentralWidget(mainFrame)

    # 读取下拉框索引，显示对应大题
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
            self.input_score.setText('')  # 清空输入框
        pixmap = QPixmap(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.num}.jpg")
        self.label_pic.setPixmap(pixmap)
        self.label_pic.setScaledContents(True)

    def nextbutClicked(self):
        # score[学号列表[学号索引]][大题索引] = 所输分数
        # if self.input_score.text() is None:
        #     print(123)
        # else:
        if self.input_score.text() == '':
            self.input_score.setText('0')
        self.scores[self.StuID[self.index_ID]][self.num] = float(self.input_score.text())
        # total_score[学号列表[学号索引]]
        self.total_score[self.StuID[self.index_ID]] += float(self.input_score.text())
        # 点击了下一个，则学号索引 + 1
        self.index_ID += 1
        # 若已经是最后一张，则无变化
        if self.index_ID >= len(self.StuID):
            self.index_ID = len(self.StuID) - 1
        else:
            self.input_score.clear()  # 清空输入框
        pixmap = QPixmap(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.num}.jpg")
        self.label_pic.setPixmap(pixmap)
        self.label_pic.setScaledContents(True)
        # print(111)


if __name__ == '__main__':
    app = QApplication(sys.argv)

    main = ZhuguantiUI([202240090, 202243050, 202243051], 2)
    main.show()

    sys.exit(app.exec_())
