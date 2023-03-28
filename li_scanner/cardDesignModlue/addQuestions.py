# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'cardDesign.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QWidget


class CustomizeTitleInformation(QWidget):
    def __init__(self):
        super().__init__()
        # Form = FatherWidget
        self.setObjectName("Form")
        self.resize(892, 500)
        self.label = QtWidgets.QLabel(self)
        self.label.setGeometry(QtCore.QRect(350, 30, 121, 16))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self)
        self.label_2.setGeometry(QtCore.QRect(30, 90, 72, 15))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self)
        self.label_3.setGeometry(QtCore.QRect(230, 90, 72, 15))
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self)
        self.label_4.setGeometry(QtCore.QRect(30, 130, 72, 15))
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self)
        self.label_5.setGeometry(QtCore.QRect(230, 130, 72, 15))
        self.label_5.setObjectName("label_5")

        # 单选多选
        self.question_type = QtWidgets.QComboBox(self)
        self.question_type.setGeometry(QtCore.QRect(110, 130, 87, 25))
        self.question_type.setObjectName("question_type")
        # 小题个数
        self.child_number = QtWidgets.QSpinBox(self)
        self.child_number.setGeometry(QtCore.QRect(320, 130, 87, 25))
        self.child_number.setObjectName("child_number")
        # 起始题号
        self.end_qid = QtWidgets.QSpinBox(self)
        self.end_qid.setGeometry(QtCore.QRect(320, 90, 87, 25))
        self.end_qid.setObjectName("end_qid")
        # 结束题号
        self.start_qid = QtWidgets.QSpinBox(self)
        self.start_qid.setGeometry(QtCore.QRect(110, 90, 87, 25))
        self.start_qid.setObjectName("start_qid")

        # 预览表格
        self.pre_form = QtWidgets.QTableWidget(self)
        self.pre_form.setGeometry(QtCore.QRect(450, 90, 339, 360))
        self.pre_form.setObjectName("tableWidget")

        # 确定按钮
        self.enter = QtWidgets.QPushButton(self)
        self.enter.setGeometry(QtCore.QRect(110, 170, 87, 28))
        self.enter.setObjectName("enter")

        # 删除按钮
        self.delete = QtWidgets.QPushButton(self)
        self.delete.setGeometry(QtCore.QRect(320, 170, 87, 28))
        self.delete.setObjectName("delete")

        # 保存按钮
        self.save = QtWidgets.QPushButton(self)
        self.save.setGeometry(QtCore.QRect(110, 210, 87, 28))
        self.save.setObjectName("save")
        self.retranslateUi(self)
        QtCore.QMetaObject.connectSlotsByName(self)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.label.setText(_translate("Form", "批量设置客观题"))
        self.label_2.setText(_translate("Form", "开始题号"))
        self.label_3.setText(_translate("Form", "结束题号"))
        self.label_4.setText(_translate("Form", "题类型"))
        self.label_5.setText(_translate("Form", "小题个数"))

        # 设置题目类型多选框
        self.question_type.addItems(["选择", "填空", "大题"])
        # 确定键
        self.enter.setText(_translate("Form", "确定"))
        self.enter.clicked.connect(self.insertQuestionsMessage)
        # 删除键
        self.delete.setText(_translate("Form", "删除"))
        self.delete.clicked.connect(self.delOneHeader)
        # 保存键
        self.save.setText(_translate("Form", "保存"))

        # 设置预览表格
        self.pre_form.setColumnCount(4)
        self.pre_form.setRowCount(1)
        self.pre_form.resizeRowsToContents()
        self.pre_form.resizeColumnsToContents()
        # 设置表头
        horizontalHeader = ["开始题号", "结束题号", "类型", "小题个数"]
        self.pre_form.setHorizontalHeaderLabels(horizontalHeader)
        # 设置列宽
        for index in range(self.pre_form.colorCount()):
            self.pre_form.setColumnWidth(index, 80)
        # 设置不可编辑
        self.pre_form.setEditTriggers(QTableWidget.NoEditTriggers)

    def insertQuestionsMessage(self):
        start_qid = self.start_qid.value()
        end_qid = self.end_qid.value()
        question_type = self.question_type.currentText()
        question_value = self.child_number.value()
        # 插入行
        row_count = self.pre_form.rowCount()
        self.pre_form.insertRow(row_count)

        self.pre_form.setItem(row_count, 0, QTableWidgetItem(str(start_qid)))
        self.pre_form.setItem(row_count, 1, QTableWidgetItem(str(end_qid)))
        self.pre_form.setItem(row_count, 2, QTableWidgetItem(question_type))
        self.pre_form.setItem(row_count, 3, QTableWidgetItem(str(question_value)))

    def delOneHeader(self):
        row_count = self.pre_form.rowCount()
        self.pre_form.removeRow(row_count - 1)
