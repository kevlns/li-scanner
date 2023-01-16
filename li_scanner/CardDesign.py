# -*- coding: utf-8 -*-
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QWidget, QHBoxLayout, QTableWidgetItem

from li_scanner.addQuestions import ObjectiveQuestions
from li_scanner.cardCreate import CardCreate


def insertQuestionsMessage(form: QtWidgets.QTableWidget):
    # 插入行
    row_count = form.rowCount()
    form.insertRow(row_count)


def delOneHeader(form: QtWidgets.QTableWidget):
    row_count = form.rowCount()
    form.removeRow(row_count - 1)


class CardDesign(QWidget):
    def __init__(self):
        super().__init__()
        self.resize(750, 712)
        # 布局
        # 左
        self.left = QtWidgets.QFrame(self)
        self.left.setGeometry(QtCore.QRect(10, 90, 381, 500))
        self.left.setFrameShadow(QtWidgets.QFrame.Raised)
        self.left.setObjectName("left")
        # 左上
        self.leftTop = QtWidgets.QFrame(self.left)
        self.leftTop.setGeometry(QtCore.QRect(10, 10, 371, 90))
        self.leftTop.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.leftTop.setFrameShadow(QtWidgets.QFrame.Raised)
        self.leftTop.setObjectName("leftTop")
        # 左下
        self.leftBottom = QtWidgets.QFrame(self.left)
        self.leftBottom.setGeometry(QtCore.QRect(10, 100, 371, 350))
        self.leftBottom.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.leftBottom.setFrameShadow(QtWidgets.QFrame.Raised)
        self.leftBottom.setObjectName("leftBottom")
        # 右
        self.right = QtWidgets.QFrame(self)
        self.right.setGeometry(QtCore.QRect(430, 10, 300, 800))
        self.right.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.right.setFrameShadow(QtWidgets.QFrame.Raised)
        self.right.setObjectName("right")

        hLayout = QHBoxLayout()
        hLayout.addWidget(self.left)
        hLayout.addWidget(self.right)

        # 客观题设计窗口
        self.objectiveDesignForm = ObjectiveQuestions()
        # 客观题表格
        self.preForm = QtWidgets.QTableWidget(self.right)
        # 控制按钮
        self.objEnter = QtWidgets.QPushButton(self.right)
        self.objDelete = QtWidgets.QPushButton(self.right)
        self.enterObjectiveDesign = QtWidgets.QPushButton(self.right)

        # 生成答题卡按钮
        self.createCard = QtWidgets.QPushButton(self.leftBottom)

        # 调用表格初始化函数
        self.initRight()

        # 纸张大小
        self.paperSize = QtWidgets.QComboBox(self.leftTop)
        self.paperSize.setGeometry(QtCore.QRect(130, 20, 87, 22))
        self.paperSize.setObjectName("paperSize")
        # 答题卡标题
        self.cardTitle = QtWidgets.QLineEdit(self.leftTop)
        self.cardTitle.setGeometry(QtCore.QRect(130, 60, 231, 21))
        self.cardTitle.setObjectName("cardTitle")
        # 提醒信息
        self.msg = QtWidgets.QTextEdit(self.leftBottom)
        self.msg.setGeometry(QtCore.QRect(10, 170, 341, 87))
        self.msg.setObjectName("msg")
        # 学号位数
        self.idDigits = QtWidgets.QSpinBox(self.leftBottom)
        self.idDigits.setGeometry(QtCore.QRect(100, 50, 91, 22))
        self.idDigits.setObjectName("idDigits")
        # 选择框
        self.haveIdentity = QtWidgets.QCheckBox(self.leftBottom)
        self.haveClass = QtWidgets.QCheckBox(self.leftBottom)
        self.haveName = QtWidgets.QCheckBox(self.leftBottom)
        # 标签
        self.label = QtWidgets.QLabel(self.leftTop)
        self.label.setGeometry(QtCore.QRect(10, 20, 72, 15))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.leftTop)
        self.label_2.setGeometry(QtCore.QRect(10, 60, 81, 16))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.leftBottom)
        self.label_3.setGeometry(QtCore.QRect(10, 10, 72, 15))
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.leftBottom)
        self.label_4.setGeometry(QtCore.QRect(10, 50, 72, 15))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.leftBottom)
        self.label_5.setGeometry(QtCore.QRect(10, 90, 72, 15))
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.leftBottom)
        self.label_6.setGeometry(QtCore.QRect(10, 130, 91, 16))
        self.label_6.setObjectName("label_6")
        # 分割线
        self.line = QtWidgets.QFrame(self.leftBottom)
        self.line.setGeometry(QtCore.QRect(10, -10, 341, 31))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")

        self.retranslateUi(self)
        QtCore.QMetaObject.connectSlotsByName(self)

    def initRight(self):
        _translate = QtCore.QCoreApplication.translate
        # 客观题按钮
        self.objEnter.setGeometry(QtCore.QRect(10, 20, 87, 28))
        self.objDelete.setGeometry(QtCore.QRect(110, 20, 87, 28))
        self.objEnter.setText(_translate("right", "添加"))
        self.objDelete.setText(_translate("right", "删除"))

        # 打开批量添加客观题
        self.enterObjectiveDesign.setGeometry(QtCore.QRect(10, 55, 120, 28))
        self.enterObjectiveDesign.setText(_translate("right", "批量添加"))

        # 连接按钮与函数
        self.objEnter.clicked.connect(lambda: insertQuestionsMessage(self.preForm))
        self.objDelete.clicked.connect(lambda: delOneHeader(self.preForm))
        self.enterObjectiveDesign.clicked.connect(self.createObjectiveDesignForm)

        self.preForm.setGeometry(QtCore.QRect(10, 90, 270, 600))
        self.preForm.setObjectName("objForm")
        # 设置题目表格
        self.preForm.setColumnCount(3)
        self.preForm.setRowCount(1)
        self.preForm.resizeRowsToContents()
        self.preForm.resizeColumnsToContents()
        # 设置表头
        horizontalHeader = ["题号", "类型", "小题个数"]
        self.preForm.setHorizontalHeaderLabels(horizontalHeader)
        # 设置列宽
        for index in range(self.preForm.columnCount()):
            self.preForm.setColumnWidth(index, 80)
        self.preForm.verticalHeader().hide()

    def createObjectiveDesignForm(self):
        self.objectiveDesignForm.save.clicked.connect(self.addObjectiveQuestion)
        self.objectiveDesignForm.show()

    def addObjectiveQuestion(self):
        rowCount = self.objectiveDesignForm.pre_form.rowCount()
        for i in range(rowCount - 1):
            start_qid = self.objectiveDesignForm.pre_form.item(i + 1, 0).text()
            end_qid = self.objectiveDesignForm.pre_form.item(i + 1, 1).text()
            question_type = self.objectiveDesignForm.pre_form.item(i + 1, 2).text()
            value = self.objectiveDesignForm.pre_form.item(i + 1, 3).text()
            for j in range(int(start_qid), int(end_qid) + 1):
                # 插入新行
                count = self.preForm.rowCount()
                self.preForm.insertRow(count)
                # 设值
                self.preForm.setItem(count, 0, QTableWidgetItem(str(j)))
                self.preForm.setItem(count, 1, QTableWidgetItem(question_type))
                self.preForm.setItem(count, 2, QTableWidgetItem(value))

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "答题卡设计"))
        # 生成答题卡按钮
        self.createCard.setGeometry(QtCore.QRect(10, 300, 100, 40))
        self.createCard.setText(_translate("leftBottom", "生成答题卡"))
        self.createCard.clicked.connect(self.createQuestionCard)

        # 是否有身份证号
        self.haveIdentity.setGeometry(QtCore.QRect(260, 90, 91, 19))
        self.haveIdentity.setObjectName("identity")
        self.haveIdentity.setText(_translate("Form", "身份证号"))
        # 是否有班级
        self.haveClass.setGeometry(QtCore.QRect(180, 90, 91, 19))
        self.haveClass.setObjectName("haveClass")
        self.haveClass.setText(_translate("Form", "班级"))
        # 是否有名字
        self.haveName.setGeometry(QtCore.QRect(100, 90, 91, 19))
        self.haveName.setObjectName("haveName")
        self.haveName.setText(_translate("Form", "姓名"))

        self.label.setText(_translate("Form", "纸张大小"))
        self.label_2.setText(_translate("Form", "答题卡标题"))
        self.label_4.setText(_translate("Form", "学号位数"))
        self.label_3.setText(_translate("Form", "信息区"))
        self.label_5.setText(_translate("Form", "可选项:"))
        self.label_6.setText(_translate("Form", "考试提醒信息"))

    def createQuestionCard(self):
        selNumberList = []
        optNumOfSelQList = []
        fillNumberList = []
        subNumberList = []
        subChNumberList = []
        title = self.cardTitle.text()
        warnMsg = self.msg.toPlainText()
        for i in range(self.preForm.rowCount() - 1):
            num = str(self.preForm.item(i + 1, 1).text())
            if num == '选择':
                selNumberList.append(int(self.preForm.item(i + 1, 0).text()))
                optNumOfSelQList.append(int(self.preForm.item(i + 1, 2).text()))
            if num == '填空':
                fillNumberList.append(int(self.preForm.item(i + 1, 0).text()))
            if num == '大题':
                subNumberList.append(int(self.preForm.item(i + 1, 0).text()))
                subChNumberList.append(int(self.preForm.item(i + 1, 2).text()))
        idDigits = self.idDigits.value()
        CardCreate(idDigits, selNumberList, optNumOfSelQList,fillNumberList, subNumberList, subChNumberList,title,warnMsg)
