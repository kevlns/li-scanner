import sys

from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt, QCoreApplication
from PyQt5.QtGui import QPixmap, QDoubleValidator
from PyQt5.QtWidgets import QApplication, QWidget, QComboBox, QLabel, QPushButton, QVBoxLayout, QHBoxLayout, QLineEdit, \
    QMessageBox


class SubUI(QWidget):

    def __init__(self, StuID: list, Numofquestion: int):
        """
            :param StuID: 学号列表
            :param Numofquestion: 分割出的图片个数
        """
        super(SubUI, self).__init__()

        self.StuID = StuID  # 学号列表
        self.total_num = Numofquestion  # 题目数量
        self.num = 0  # 第几道大题

        self.scores = {}  # 储存每人每道题得分
        self.total_score = {}  # 每人大题总分

        self.combobox_num = QComboBox(self)  # 题号选择下拉框

        for Stu_id in self.StuID:
            self.scores[Stu_id] = []
            self.total_score[Stu_id] = 0
            for nq in range(0, Numofquestion):
                self.scores[Stu_id].append(0)

        self.initUI()

    def initUI(self):
        # 设置窗口大小
        self.resize(800, 495)
        # 设置窗口标题
        self.setWindowTitle("轻阅-主观题批改")

        for index in range(0, self.total_num):
            self.combobox_num.addItem("第" + str(index + 1) + "题")
        self.combobox_num.setStyleSheet('font-size: 45px;')
        # self.combobox_num.currentIndexChanged.connect(self.SetNum)
        hbox_win1_combobox = QHBoxLayout()
        hbox_win1_combobox.addSpacing(60)
        hbox_win1_combobox.addWidget(self.combobox_num)
        hbox_win1_combobox.addSpacing(60)

        label_title = QLabel(self)
        label_title.setText("请选择批改题目：")
        label_title.setStyleSheet('font-size: 50px;font-weight: bold')

        button_yes = QPushButton("确认")
        button_yes.setStyleSheet('font-size: 30px;')
        button_yes.clicked.connect(self.correction)
        button_no = QPushButton("取消")
        button_no.setStyleSheet('font-size: 30px;')
        button_no.clicked.connect(QCoreApplication.instance().quit)

        hox_win1_button = QHBoxLayout()
        hox_win1_button.addWidget(button_yes)
        hox_win1_button.addStretch(1)
        hox_win1_button.addWidget(button_no)

        vbox_win1 = QVBoxLayout()
        vbox_win1.addStretch(1)
        vbox_win1.addWidget(label_title)
        vbox_win1.addSpacing(30)
        vbox_win1.addLayout(hbox_win1_combobox)
        vbox_win1.addSpacing(45)
        vbox_win1.addLayout(hox_win1_button)
        vbox_win1.addStretch(1)

        hbox_win1 = QHBoxLayout()
        hbox_win1.addSpacing(100)
        hbox_win1.addLayout(vbox_win1)
        hbox_win1.addSpacing(100)

        self.setLayout(hbox_win1)

    def correction(self):
        self.hide()
        self.num = self.combobox_num.currentIndex()
        self.main_cor = CorrectionUI(self.StuID, self.num, self.scores, self.total_score)
        # print(self.StuID, self.num)
        self.main_cor.show()

    # # 读取下拉框索引，显示对应大题
    # def SetNum(self):
    #     self.num = self.combobox_num.currentIndex()
    #     pixmap = QPixmap(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.num}.jpg")
    #     self.label_pic.setPixmap(pixmap)
    #     self.label_pic.setScaledContents(True)


class CorrectionUI(QWidget):

    def __init__(self, StuID: list, SubNum: int, Scores: dict, TotalScore: dict):
        super(CorrectionUI, self).__init__()

        self.StuID = StuID  # 学号列表
        self.sub_num = SubNum  # 题目序号
        self.scores = Scores
        self.total_score = TotalScore
        self.index_ID = 0  # 列表索引
        self.path = "../../record/subjective/"  # 题目文件夹存放路径
        # self.scores = {}  # 储存每人每道题得分
        # self.total_score = {}  # 每人大题总分

        self.label_pic = QLabel()  # 图片显示
        self.lineedit_score = QLineEdit()  # 分数输入框

        self.initUI()

    def initUI(self):
        # 设置窗口大小
        self.resize(1000, 618)
        # 设置窗口标题
        self.setWindowTitle(f"轻阅-第{self.sub_num + 1}题批改")

        # 显示题号
        label_title = QLabel(self)
        label_title.setText(f"第{self.sub_num + 1}题：")
        label_title.setStyleSheet('font-size: 50px;font-weight: bold')

        # 加载图片
        pixmap = QPixmap(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.sub_num}.jpg")
        self.label_pic.setPixmap(pixmap)
        self.label_pic.setScaledContents(True)

        # 得分输入
        label_sco = QLabel("得分：")
        label_sco.setStyleSheet('font-size: 25px;')

        # 创建打分及切换其他人答案水平布局
        self.lineedit_score.setFixedWidth(150)
        self.lineedit_score.setFixedHeight(40)
        # 校验器，只能输入两位整数和一位小数
        doubleValidator = QDoubleValidator(self)
        doubleValidator.setRange(0, 99)
        doubleValidator.setNotation(QDoubleValidator.StandardNotation)
        doubleValidator.setDecimals(1)
        self.lineedit_score.setValidator(doubleValidator)

        hbox_win2_sco = QHBoxLayout()
        hbox_win2_sco.addSpacing(20)
        hbox_win2_sco.addWidget(label_sco, Qt.AlignLeft | Qt.AlignCenter)
        hbox_win2_sco.addWidget(self.lineedit_score, Qt.AlignLeft | Qt.AlignCenter)

        button_next = QPushButton("下一张")
        button_last = QPushButton("上一张")
        button_next.clicked.connect(self.nextbutClicked)
        button_last.clicked.connect(self.lastbutClicked)

        hbox_win2_but = QHBoxLayout()
        hbox_win2_but.addStretch(1)
        hbox_win2_but.addWidget(button_last, 2)
        hbox_win2_but.addStretch(2)
        hbox_win2_but.addWidget(button_next, 2)
        hbox_win2_but.addStretch(1)

        vbox_win2 = QVBoxLayout()
        vbox_win2.addSpacing(50)
        vbox_win2.addWidget(label_title, Qt.AlignLeft | Qt.AlignCenter)
        vbox_win2.addWidget(self.label_pic, Qt.AlignJustify | Qt.AlignCenter)
        vbox_win2.addSpacing(15)
        vbox_win2.addLayout(hbox_win2_sco)
        vbox_win2.addSpacing(45)
        vbox_win2.addLayout(hbox_win2_but)
        vbox_win2.addSpacing(75)

        hbox_win2 = QHBoxLayout()
        hbox_win2.addSpacing(125)
        hbox_win2.addLayout(vbox_win2)
        hbox_win2.addSpacing(125)

        self.setLayout(hbox_win2)

    def lastbutClicked(self):
        if self.lineedit_score.text() == '':
            self.lineedit_score.setText('0')
        # score[学号列表[学号索引]][大题索引] = 所输分数
        self.scores[self.StuID[self.index_ID]][self.sub_num] = float(self.lineedit_score.text())

        self.index_ID -= 1
        self.total_score[self.StuID[self.index_ID]] -= self.scores[self.StuID[self.index_ID]][self.sub_num]
        if self.index_ID < 0:
            self.index_ID = 0
            messagebox_last = QMessageBox.information(QWidget(), "提示", "已是最后一份", QMessageBox.Ok)
        else:
            self.lineedit_score.setText(str(self.scores[self.StuID[self.index_ID]][self.sub_num]))  # 显示分数
        pixmap = QPixmap(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.sub_num}.jpg")
        self.label_pic.setPixmap(pixmap)
        self.label_pic.setScaledContents(True)

    def nextbutClicked(self):
        if self.lineedit_score.text() == '':
            self.lineedit_score.setText('0')
        # score[学号列表[学号索引]][大题索引] = 所输分数
        self.scores[self.StuID[self.index_ID]][self.sub_num] = float(self.lineedit_score.text())
        # total_score[学号列表[学号索引]]
        self.total_score[self.StuID[self.index_ID]] += float(self.lineedit_score.text())
        # 点击了下一个，则学号索引 + 1
        self.index_ID += 1
        # 若已经是最后一张，则无变化
        if self.index_ID >= len(self.StuID):
            self.index_ID = len(self.StuID) - 1
            messagebox_next = QMessageBox.information(QWidget(), "提示", "已是最后一份", QMessageBox.Ok)
        else:
            self.lineedit_score.clear()  # 清空输入框

        pixmap = QPixmap(self.path + f"{self.StuID[self.index_ID]}/{self.StuID[self.index_ID]}-{self.sub_num}.jpg")
        self.label_pic.setPixmap(pixmap)
        self.label_pic.setScaledContents(True)

        if self.scores[self.StuID[self.index_ID]][self.sub_num] != 0:  # 若成绩不为零，显示成绩
            self.lineedit_score.setText(str(self.scores[self.StuID[self.index_ID]][self.sub_num]))

    def closeEvent(self, event):
        reply = QMessageBox(QMessageBox.Question, self.tr("关闭确认"),
                            self.tr("确认退出？"), QtWidgets.QMessageBox.NoButton, self)
        button_close = reply.addButton(self.tr("确认"), QtWidgets.QMessageBox.YesRole)
        reply.addButton(self.tr("取消"), QtWidgets.QMessageBox.NoRole)
        reply.exec_()
        if reply.clickedButton() == button_close:
            event.accept()
            QtWidgets.qApp.quit()
        else:
            event.ignore()


if __name__ == '__main__':
    app = QApplication(sys.argv)

    main = SubUI([202240090, 202243050, 202243051], 2)
    # main = CorrectionUI([202240090, 202243050, 202243051], 0,
    #                     {202240090: [0, 0], 202243050: [0, 0], 202243051: [0, 0]},
    #                     {202240090: 0, 202243050: 0, 202243051: 0})
    main.show()

    sys.exit(app.exec_())
