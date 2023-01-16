# -*- coding: utf-8 -*-
import xlwings as xw

# 一页包含的行数和列数
columnNum = 26
rowNum = 62
# 每一页的开始位置
pageBeginNum = [1, 63]


class CardCreate:
    def __init__(self, idDigits: int, selNumberList: list, optNumOfSelQList: list,
                 fillNumberList: list, subNumberList: list, subChNumberList: list, cardTitle: str, warnMsg: str):
        # 选择题和填空题的分割位置
        self.partition1 = 0
        # 填空题和大题分割位置
        self.partition2 = 0
        # 学号位数
        self.idDigits = idDigits
        # 选择题题号
        self.selNumberList = selNumberList
        # 选择题选项个数
        self.optNumOfSelQList = optNumOfSelQList
        # 填空题题号
        self.fillNumberList = fillNumberList
        # 大题题号
        self.subNumberList = subNumberList
        # 每个大题的小题个数
        self.subChNumberList = subChNumberList
        # 答题卡标题
        self.cardTitle = cardTitle
        # 考试需知
        self.warnMsg = warnMsg
        app = xw.App(visible=True, add_book=False)
        app.display_alerts = False  # 关闭一些提示信息，可以加快运行速度。 默认为 True。
        app.screen_updating = True  # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
        # 工作溥
        wb = app.books.add()
        self.sht = wb.sheets.active
        # 获取工作表
        # 设置整个表格列宽
        self.sht.range((1, 1), (1, columnNum)).column_width = 2.78
        # 画大的定位点
        self.drawBigDot()
        # 画信息区小的定位点
        self.drawSmallDot()
        # 填充信息区
        self.fillMsgArea()
        # 填充选择题
        self.fillSelQuestion()
        # 填充填空题
        self.fillFillQuestion()
        # 填充大题
        self.fillSubQuestion()
        wb.save(r'../card/card_template.xlsx')
        wb.close()
        app.quit()

    # 画大定位点函数
    def drawBigDot(self):
        # 如果扩展成A3纸的话，把这个函数写成循环形式
        # 第一页
        # 合并单元格
        self.sht.range((1, 1), (2, 2)).api.Merge()
        self.sht.range((1, columnNum - 1), (2, columnNum)).api.Merge()
        self.sht.range((rowNum - 1, 1), (rowNum, 2)).api.Merge()
        self.sht.range((rowNum - 1, columnNum - 1), (rowNum, columnNum)).api.Merge()
        # 填充内容
        self.sht.range(1, 1).value = '█'
        self.sht.range(1, 1).api.Font.Size = 48
        self.sht.range(1, 1).api.HorizontalAlignment = -4108  # -4108 水平居中。 -4131 靠左，-4152 靠右
        self.sht.range(1, columnNum - 1).value = '█'
        self.sht.range(1, columnNum - 1).api.Font.Size = 48
        self.sht.range(1, columnNum - 1).api.HorizontalAlignment = -4108  # -4108 水平居中。 -4131 靠左，-4152 靠右
        self.sht.range(rowNum - 1, 1).value = '█'
        self.sht.range(rowNum - 1, 1).api.Font.Size = 48
        self.sht.range(rowNum - 1, 1).api.HorizontalAlignment = -4108  # -4108 水平居中。 -4131 靠左，-4152 靠右
        self.sht.range(rowNum - 1, columnNum - 1).value = '█'
        self.sht.range(rowNum - 1, columnNum - 1).api.Font.Size = 48
        self.sht.range(rowNum - 1, columnNum - 1).api.HorizontalAlignment = -4108  # -4108 水平居中。 -4131 靠左，-4152 靠右

        # 第二页
        # 合并单元格
        self.sht.range((pageBeginNum[1], 1), (pageBeginNum[1] + 1, 2)).api.Merge()
        self.sht.range((pageBeginNum[1], columnNum - 1), (pageBeginNum[1] + 1, columnNum)).api.Merge()
        self.sht.range((pageBeginNum[1] + rowNum - 2, 1), (pageBeginNum[1] + rowNum - 1, 2)).api.Merge()
        self.sht.range((pageBeginNum[1] + rowNum - 2, columnNum - 1),
                       (pageBeginNum[1] + rowNum - 1, columnNum)).api.Merge()
        # 填充内容
        self.sht.range(pageBeginNum[1], 1).value = '█'
        self.sht.range(pageBeginNum[1], 1).api.Font.Size = 48
        self.sht.range(pageBeginNum[1], 1).api.HorizontalAlignment = -4108  # -4108 水平居中。 -4131 靠左，-4152 靠右
        self.sht.range(pageBeginNum[1], columnNum - 1).value = '█'
        self.sht.range(pageBeginNum[1], columnNum - 1).api.Font.Size = 48
        self.sht.range(pageBeginNum[1], columnNum - 1).api.HorizontalAlignment = -4108  # -4108 水平居中。 -4131 靠左，-4152 靠右
        self.sht.range(pageBeginNum[1] + rowNum - 2, 1).value = '█'
        self.sht.range(pageBeginNum[1] + rowNum - 2, 1).api.Font.Size = 48
        self.sht.range(pageBeginNum[1] + rowNum - 2, 1).api.HorizontalAlignment = -4108  # -4108 水平居中。 -4131 靠左，-4152 靠右
        self.sht.range(pageBeginNum[1] + rowNum - 2, columnNum - 1).value = '█'
        self.sht.range(pageBeginNum[1] + rowNum - 2, columnNum - 1).api.Font.Size = 48
        self.sht.range(pageBeginNum[1] + rowNum - 2, columnNum - 1).api.HorizontalAlignment = -4108  # -4108 水平居中。

    # 画小定位点函数
    def drawSmallDot(self):
        # 水平点
        self.sht.range((2, 3), (2, columnNum - 2)).value = '▃'
        self.sht.range((2, 3), (2, columnNum - 2)).api.Font.Size = 14
        self.sht.range((2, 3), (2, columnNum - 2)).api.HorizontalAlignment = -4108
        # 垂直点
        self.sht.range((3, 2), (13, 2)).value = '▍'
        self.sht.range((3, 2), (13, 2)).api.Font.Size = 10
        self.sht.range((3, 2), (13, 2)).api.HorizontalAlignment = -4108

    # 填充选择题函数
    def fillSelQuestion(self):
        # 选择题开始行数
        selBeginRow = 16
        # 选择题开始列数
        selBeginCol = 3
        # 设置字体和居右
        self.sht.range((selBeginRow, 3), (rowNum - 2, columnNum - 2)).api.Font.Size = 7
        self.sht.range((selBeginRow, 3), (rowNum - 2, columnNum - 2)).api.HorizontalAlignment = -4152
        # 行指针
        rowPtr = selBeginRow
        # 列指针
        colPtr = selBeginCol
        # 水平偏移
        colOffset = 0
        # 分割线集合
        VerDivLine = set()
        # 遍历选择题号
        for i in range(len(self.selNumberList)):
            if rowPtr < selBeginRow + 5:
                # 记录第一行分割线
                VerDivLine.add(colPtr)
            else:
                if not (colPtr in VerDivLine):
                    flag = 1
                    # 找到比现在colPtr大的第一个分割线实现对齐
                    for m in VerDivLine:
                        if colPtr <= m and flag == 1:
                            colPtr = m
                            flag = 0
                    # 如果没有比colPtr大的分割线，直接跳到下一行
                    if flag == 1:
                        colPtr = selBeginCol
                        rowPtr = rowPtr + 6
            # 填充题号
            # print('rowPtr: ', rowPtr, 'colPtr: ', colPtr, 'i: ', i)
            self.sht.range(rowPtr, colPtr).value = self.selNumberList[i]
            # 垂直点
            self.sht.range(rowPtr, 2).value = '▍'
            self.sht.range(rowPtr, 2).api.Font.Size = 10
            self.sht.range(rowPtr, 2).api.HorizontalAlignment = -4108
            # print('row: ', rowPtr, 'col: ', colPtr)
            # 填充选项
            for j in range(self.optNumOfSelQList[i]):
                self.sht.range(rowPtr, colPtr + j + 1).value = '[ ' + str(chr(ord('A') + j)) + ' ]'
            rowPtr = rowPtr + 1
            if (i + 1) % 5 == 0 and i < len(self.selNumberList) - 5:
                preColOffset = 0
                for k in range(5):
                    # 之后五个题的最长长度
                    if self.optNumOfSelQList[i + k] > colOffset:
                        colOffset = self.optNumOfSelQList[i + k]
                    # 之前五个题的最长长度
                    if self.optNumOfSelQList[i - k] > preColOffset:
                        preColOffset = self.optNumOfSelQList[i - k]
                # 如果超过截止范围,换下一行
                if colPtr + colOffset + preColOffset > columnNum - 3:
                    colPtr = selBeginCol
                    rowPtr = rowPtr + 1
                # 没超过就继续在水平延申
                else:
                    rowPtr = rowPtr - 5
                    colPtr = colPtr + preColOffset + 1
                # print('colOffset: ', colOffset)
                colOffset = 0
        self.partition1 = rowPtr + 1

    # 填充填空题
    def fillFillQuestion(self):
        print('partition1 :', self.partition1)
        # 文字
        self.sht.range((self.partition1, 4), (self.partition1, 5)).api.Merge()
        self.sht.range((self.partition1, 4)).value = '填空题'
        self.sht.range((self.partition1, 4)).api.HorizontalAlignment = -4108
        # 填空题画线框
        # self.sht.range((subline1, 3), (rowNum - 2, columnNum - 2)).api.merge()
        self.sht.range((self.partition1, 4), (rowNum - 3, columnNum - 3)).api.Borders(7).LineStyle = 1
        self.sht.range((self.partition1, 4), (rowNum - 3, columnNum - 3)).api.Borders(8).LineStyle = 1
        self.sht.range((self.partition1, 4), (rowNum - 3, columnNum - 3)).api.Borders(9).LineStyle = 1
        self.sht.range((self.partition1, 4), (rowNum - 3, columnNum - 3)).api.Borders(10).LineStyle = 1

        self.sht.range((pageBeginNum[1] + 3, 4), (pageBeginNum[1] + rowNum - 4, columnNum - 3)).api.Borders(
            7).LineStyle = 1
        self.sht.range((pageBeginNum[1] + 3, 4), (pageBeginNum[1] + rowNum - 4, columnNum - 3)).api.Borders(
            8).LineStyle = 1
        self.sht.range((pageBeginNum[1] + 3, 4), (pageBeginNum[1] + rowNum - 4, columnNum - 3)).api.Borders(
            9).LineStyle = 1
        self.sht.range((pageBeginNum[1] + 3, 4), (pageBeginNum[1] + rowNum - 4, columnNum - 3)).api.Borders(
            10).LineStyle = 1
        self.sht.range((pageBeginNum[1] + 3, 4), (pageBeginNum[1] + rowNum - 4, columnNum - 3)).api.Font.Size = 7

        fillContent = ''
        for i in range(int(len(self.fillNumberList) / 5)):
            self.sht.range((self.partition1 + 2 + i * 2, 4), (self.partition1 + 2 + i * 2, columnNum - 3)).api.Merge()
            for j in range(4):
                fillContent = fillContent + str(self.fillNumberList[i * 5 + j]) + '______________             '
            fillContent = fillContent + str(self.fillNumberList[i * 5 + 4]) + '______________'
            self.sht.range((self.partition1 + 2 + i * 2, 4)).api.HorizontalAlignment = -4108
            self.sht.range((self.partition1 + 2 + i * 2, 4)).value = fillContent
            fillContent = ''

        self.sht.range((self.partition1 + 2 + len(self.fillNumberList) / 5 * 2, 4),
                       (self.partition1 + 2 + len(self.fillNumberList) / 5 * 2, columnNum - 3)).api.Merge()
        for i in range(int(len(self.fillNumberList) % 5)):
            fillContent = fillContent + str(
                self.fillNumberList[i + int(len(self.fillNumberList) / 5) * 5]) + '____________            '
            self.sht.range((self.partition1 + 2 + len(self.fillNumberList) / 5 * 2, 4)).value = fillContent
            self.sht.range((self.partition1 + 2 + len(self.fillNumberList) / 5 * 2, 4)).api.HorizontalAlignment = -4131

    # 填充信息区函数
    def fillMsgArea(self):
        # 答题卡标题部分
        self.sht.range((1, 3), (1, columnNum - 2)).api.Merge()
        self.sht.range((1, 3)).value = self.cardTitle
        self.sht.range(1, 3).api.HorizontalAlignment = -4108

        # 班级姓名部分
        # 合并单元格
        self.sht.range((4, 4), (5, 4)).api.Merge()
        self.sht.range((6, 4), (7, 4)).api.Merge()
        self.sht.range((4, 5), (5, 6)).api.Merge()
        self.sht.range((6, 5), (7, 6)).api.Merge()
        # 画线框
        self.sht.range((4, 4), (7, 6)).api.Borders(7).LineStyle = 1
        self.sht.range((4, 4), (7, 6)).api.Borders(8).LineStyle = 1
        self.sht.range((4, 4), (7, 6)).api.Borders(9).LineStyle = 1
        self.sht.range((4, 4), (7, 6)).api.Borders(10).LineStyle = 1
        self.sht.range((4, 4), (4 + 3, 4)).api.Borders(10).LineStyle = 1
        self.sht.range((4, 4), (4 + 1, 4 + 2)).api.Borders(10).LineStyle = 1
        # 填内容
        self.sht.range((4, 4)).value = '班\n级'
        self.sht.range((6, 4)).value = '姓\n名'

        # 准考证号部分
        cardBeginRow = 4
        self.sht.range((cardBeginRow, 8), (cardBeginRow + 10, 8)).api.Merge()
        self.sht.range((cardBeginRow, 8), (cardBeginRow + 10, 8)).api.Borders(7).LineStyle = 1
        self.sht.range((cardBeginRow, 8), (cardBeginRow + 10, 8)).api.Borders(8).LineStyle = 1
        self.sht.range((cardBeginRow, 8), (cardBeginRow + 10, 8)).api.Borders(9).LineStyle = 1
        self.sht.range((cardBeginRow, 8), (cardBeginRow + 10, 8)).api.Borders(10).LineStyle = 1
        self.sht.range((cardBeginRow, 8), (cardBeginRow + 10, 8)).value = '准\n考\n证\n号'

        for i in range(self.idDigits):
            # 数字框
            self.sht.range((cardBeginRow, 9 + i), (cardBeginRow, 9 + i)).api.Borders(8).LineStyle = 1
            self.sht.range((cardBeginRow, 9 + i), (cardBeginRow, 9 + i)).api.Borders(9).LineStyle = 1
            self.sht.range((cardBeginRow, 9 + i), (cardBeginRow, 9 + i)).api.Borders(10).LineStyle = 1
            # 填写框
            # self.sht.range((4, 9 + i), (13, 9 + i)).api.merge()
            self.sht.range((cardBeginRow + 1, 9 + i), (cardBeginRow + 10, 9 + i)).api.Borders(10).LineStyle = 1
            self.sht.range((cardBeginRow + 1, 9 + i), (cardBeginRow + 10, 9 + i)).api.Borders(9).LineStyle = 1
            # self.sht.range((4, 9 + i), (13, 9 + i)).value = '[0]\n[1]\n[2]\n[3]\n[4]\n[5]\n[6]\n[7]\n[8]\n[9]'
            self.sht.range((cardBeginRow + 1, 9 + i), (cardBeginRow + 10, 9 + i)).value = [['[ 0 ]'], ['[ 1 ]'],
                                                                                           ['[ 2 ]'], ['[ 3 ]'],
                                                                                           ['[ 4 ]'],
                                                                                           ['[ 5 ]'], ['[ 6 ]'],
                                                                                           ['[ 7 ]'], ['[ 8 ]'],
                                                                                           ['[ 9 ]']]
            self.sht.range((cardBeginRow + 1, 9 + i), (cardBeginRow + 10, 9 + i)).api.HorizontalAlignment = -4108
            self.sht.range((cardBeginRow + 1, 9 + i), (cardBeginRow + 10, 9 + i)).api.Font.Size = 7

        # 填涂样例部分
        # 合并单元格
        self.sht.range((9, 4), (9, 6)).api.Merge()
        self.sht.range((10, 4), (11, 4)).api.Merge()
        self.sht.range((12, 4), (14, 4)).api.Merge()
        # 画线框
        self.sht.range((9, 4), (14, 6)).api.Borders(10).LineStyle = 1

        self.sht.range((9, 4), (9, 6)).api.Borders(7).LineStyle = 1
        self.sht.range((9, 4), (9, 6)).api.Borders(8).LineStyle = 1
        self.sht.range((9, 4), (9, 6)).api.Borders(9).LineStyle = 1
        self.sht.range((9, 4), (9, 6)).api.Borders(10).LineStyle = 1

        self.sht.range((10, 4), (11, 6)).api.Borders(7).LineStyle = 1
        self.sht.range((10, 4), (11, 6)).api.Borders(9).LineStyle = 1
        self.sht.range((10, 4), (11, 6)).api.Borders(10).LineStyle = 1

        self.sht.range((12, 4), (14, 6)).api.Borders(7).LineStyle = 1
        self.sht.range((12, 4), (14, 6)).api.Borders(9).LineStyle = 1
        self.sht.range((12, 4), (14, 6)).api.Borders(10).LineStyle = 1

        # 填内容
        self.sht.range((9, 4)).value = '填涂样例'
        self.sht.range((10, 4)).value = '正\n确'
        self.sht.range((12, 4)).value = '错\n误'

        self.sht.range((10, 5), (11, 6)).api.Merge()
        self.sht.range((10, 5)).value = '▃'
        self.sht.range(10, 5).api.Font.Size = 11
        self.sht.range(10, 5).api.HorizontalAlignment = -4108

        self.sht.range((12, 5), (14, 6)).api.Merge()
        self.sht.range((12, 5)).value = '[√ ]\n[ × ]'
        self.sht.range(12, 5).api.Font.Size = 8
        self.sht.range(12, 5).api.HorizontalAlignment = -4108

        # 考试需知部分
        self.sht.range((4, 10 + self.idDigits), (4, columnNum - 3)).api.Merge()
        self.sht.range((4, 10 + self.idDigits), (4, columnNum - 3)).api.Borders(7).LineStyle = 1
        self.sht.range((4, 10 + self.idDigits), (4, columnNum - 3)).api.Borders(8).LineStyle = 1
        self.sht.range((4, 10 + self.idDigits), (4, columnNum - 3)).api.Borders(9).LineStyle = 1
        self.sht.range((4, 10 + self.idDigits), (4, columnNum - 3)).api.Borders(10).LineStyle = 1

        self.sht.range((5, 10 + self.idDigits), (5 + 9, columnNum - 3)).api.Merge()
        self.sht.range((5, 10 + self.idDigits), (5 + 9, columnNum - 3)).api.Borders(7).LineStyle = 1
        self.sht.range((5, 10 + self.idDigits), (5 + 9, columnNum - 3)).api.Borders(9).LineStyle = 1
        self.sht.range((5, 10 + self.idDigits), (5 + 9, columnNum - 3)).api.Borders(10).LineStyle = 1
        self.sht.range((4, 10 + self.idDigits), (4, columnNum - 3)).value = '考试需知'
        self.sht.range(4, 10 + self.idDigits).api.HorizontalAlignment = -4108
        self.sht.range(5, 10 + self.idDigits).value = self.warnMsg
        self.sht.range(5, 10 + self.idDigits).api.Font.Size = 10
        self.sht.range(5, 10 + self.idDigits).api.HorizontalAlignment = -4108

    # 填充大题函数
    def fillSubQuestion(self):
        offset = 0
        self.partition2 = int(self.partition1 + (int(len(self.fillNumberList)) / 5) * 2 + 4)
        self.sht.range((self.partition2 - 1, 4)).value = '大题'
        self.sht.range((self.partition2 - 1, 4), (self.partition2 - 1, 5)).api.Merge()
        self.sht.range((self.partition2 - 1, 4)).api.HorizontalAlignment = -4108
        # print(self.partition2)
        fillContent = []
        for i in range(int(len(self.subNumberList))):
            if self.subChNumberList[i] == 1:
                fillContent.append(str(self.subNumberList[i]))
            else:
                for j in range(1, self.subChNumberList[i] + 1):
                    fillContent.append(str(str(self.subNumberList[i]) + '(' + str(j) + ')'))
        # 第一页
        curIdx = 0
        idx = 0
        for i in range(int(len(self.subNumberList))):
            # 不打印此题直接延伸到下一页无偏移
            if (self.partition2 + i * 6 == rowNum - 3) or (self.partition2 + i * 6 == rowNum - 4):
                offset = 0
                curIdx = i
                break
            # 打印完此题再到下一页
            else:
                self.sht.range((self.partition2 + i * 6, 4)).value = fillContent[idx]
                idx = idx + 1
                self.sht.range((self.partition2 + i * 6, 4),
                               (self.partition2 + i * 6, columnNum - 3)).api.Borders(8).LineStyle = 1
                if (self.partition2 + i * 6 == rowNum - 5) or (self.partition2 + i * 6 == rowNum - 6):
                    offset = rowNum - (self.partition2 + i * 6) - 2
                    curIdx = i + 1
                    break
                elif (self.partition2 + i * 6 == rowNum - 7) or (self.partition2 + i * 6 == rowNum - 8):
                    offset = 0
                    curIdx = i + 1
                    break

        # 第二页 根据偏移往下延申
        for i in range(0, len(fillContent) - curIdx):
            self.sht.range((i * 6 + pageBeginNum[1] + offset + 3, 4)).value = fillContent[idx]
            idx = idx + 1
            self.sht.range((i * 6 + pageBeginNum[1] + offset + 3, 4),
                           (i * 6 + pageBeginNum[1] + offset + 3, columnNum - 3)).api.Borders(
                8).LineStyle = 1
