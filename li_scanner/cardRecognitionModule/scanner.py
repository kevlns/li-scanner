import os
import cv2
import numpy as np
import xlwings as xw
from li_scanner.cardRecognitionModule.utils import getStuID, get_complete_card, getAnswers, SubjectiveSegmentation

# 每题选项个数
divLines = []
f = open("../../doc/divLines", "r")
lines = f.readlines()
for i in range(len(lines)):
    tmp = int(lines[i].strip('\n'))
    if tmp < 66:
        divLines.append(tmp)

optNumOfSelQList = []
f = open("../../doc/optNumOfSelQList.txt", "r")
lines = f.readlines()
for i in range(len(lines)):
    optNumOfSelQList.append(int(lines[i].strip('\n')))

# 获取每个题的坐标
cors = []
f1 = open("../../doc/cors", "r")
lines = f1.readlines()
for i in range(len(lines)):
    tmp = lines[i].strip('\n').split(' ')
    cors.append([int(tmp[0]), int(tmp[1])])  # 删除\n

app = xw.App(visible=True, add_book=False)
app.display_alerts = False  # 关闭一些提示信息，可以加快运行速度。 默认为 True。
app.screen_updating = True  # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度

wb = app.books.open(r'../../xlsx/reference_answer.xlsx')  # 打开现有的工作簿
sht = wb.sheets[0]
referenceAnswer = sht.range((2, 1), (2, len(optNumOfSelQList))).value
wb.close()

wb = app.books.add()
sht = wb.sheets[0]

# print('referenceAnswer',referenceAnswer)
for i in range(len(referenceAnswer)):
    tmp = list(referenceAnswer[i])
    strTmp = ''
    for j in range(len(tmp)):
        strTmp = strTmp + str(ord(tmp[j]) - 65 + 1)
    referenceAnswer[i] = strTmp

capture = cv2.VideoCapture(0)
StuID = []
cishu = 0

scores = {}
# detail = []
while capture.isOpened():
    ret, frame = capture.read()
    if ret:
        frame_copy = np.copy(frame)
        frame_copy = np.rot90(frame_copy)
        cv2.imshow("show", frame_copy)

        score = cv2.Laplacian(frame, cv2.CV_64F).var()
        if score > 1000:
            totalCard = get_complete_card(frame)
            if totalCard is not None:
                tmp = getStuID(totalCard, 9)
                if StuID == tmp and tmp is not None:
                    cishu = cishu + 1
                    # print('-' * 10, '第', cishu, '次的学号是', StuID, '-' * 10)
                else:
                    StuID = tmp
                    cishu = 0
                if cishu >= 5:
                    cishu = 0
                    sid = ''
                    for i in StuID:
                        sid = sid + str(i)
                    if sid in scores:
                        continue
                    print('-' * 15, sid, '检测成功', '-' * 15)
                    # print('optNumOfSelQList长度: ', optNumOfSelQList)
                    # 获取答案
                    answers = getAnswers(totalCard, optNumOfSelQList, cors)
                    if answers is None:
                        continue
                    # print('函数之前')
                    # print(answers)
                    # for i in range(len(answers)):
                    #     print('第', i + 1, '题的答案是: ', answers[i])


                    # sht.range((1, 1), (1 ,1)).value = answers

                    pics = SubjectiveSegmentation(totalCard, divLines)
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
                    scores[sid] = answers_copy

        # 按下 q 键可退出程序执行
        if cv2.waitKey(1) & 0xFF == ord('q'):
            sht.range((1, 1)).value = '考生号'
            sht.range((1, 2)).value = '分数'
            keys = list(scores.keys())
            # print(keys)
            # 存学号
            for i in range(len(keys)):
                sht.range((2 + i, 1)).value = str(keys[i])

            # 存题号
            for i in range(len(scores[keys[0]])):
                sht.range(1, i + 3).value = str(i + 1)
            # 存答案
            for j in range(len(keys)):
                # print('referenceAnswer',referenceAnswer)
                rightAnswers = 0
                for i in range(len(scores[keys[j]])):
                    if scores[keys[j]][i] == referenceAnswer[i]:
                        # print('相等')
                        rightAnswers = rightAnswers + 1
                sht.range((j + 2, 2)).value = rightAnswers

                # print(scores[keys[j]])
                sht.range((j + 2, 3), (j + 2, 3 + len(scores[keys[j]]))).value = scores[keys[j]]

            wb.save(r'../../xlsx/scores.xlsx')
            wb.close()
            app.quit()
            break

capture.release()
cv2.destroyAllWindows()
