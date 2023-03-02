import cv2
import xlwings as xw
from li_scanner.cardRecognitionModule.utils import getStuID, get_complete_card, getAnswers

# 每题选项个数
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
app.display_alerts = False    # 关闭一些提示信息，可以加快运行速度。 默认为 True。
app.screen_updating = True    # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度

wb = app.books.open(r'../../xlsx/reference_answer.xlsx')      # 打开现有的工作簿
sht = wb.sheets[0]
referenceAnswer = sht.range((2,1),(2,len(optNumOfSelQList))).value
wb.close()

for i in range(len(referenceAnswer)):
    tmp = list(referenceAnswer[i])
    for j in range(len(tmp)):
        tmp[j] = ord(tmp[j])-65+1
    referenceAnswer[i] = tmp

capture = cv2.VideoCapture(0)
StuID = []
cishu = 0

scores = {}

while capture.isOpened():
    ret, frame = capture.read()
    if ret:
        cv2.imshow("show", frame)

        score = cv2.Laplacian(frame, cv2.CV_64F).var()
        if score > 1000:
            totalCard = get_complete_card(frame)
            if totalCard is not None:
                tmp = getStuID(totalCard, 8)
                if StuID == tmp and tmp is not None:
                    cishu = cishu + 1
                    print('-' * 10, '第', cishu, '次的学号是', StuID, '-' * 10)
                else:
                    StuID = tmp
                    cishu = 0
                if cishu >= 3:
                    cishu = 0
                    sid = ''
                    for i in StuID:
                        sid = sid + str(i)

                    # print('optNumOfSelQList长度: ', optNumOfSelQList)
                    # 获取答案
                    answers = getAnswers(totalCard, optNumOfSelQList, cors)
                    if answers is None:
                        continue
                    # print(answers)
                    # for i in range(len(answers)):
                    #     print('第', i + 1, '题的答案是: ', answers[i])
                    rightAnswers = 0
                    for i in range(len(answers)):
                        if answers[i] == referenceAnswer[i]:
                            rightAnswers = rightAnswers + 1

                    scores[sid] = rightAnswers
        # 按下 q 键可退出程序执行
        if cv2.waitKey(1) & 0xFF == ord('q'):
            print(scores)
            wb1 = app.books.add()
            sht1 = wb1.sheets[0]
            sht1.range((1,1)).value = '考生号'
            sht1.range((1,2)).value = '分数'
            sht1.range((1, 2),(1,len(scores)+1)).value = list(scores.keys())
            sht1.range((2, 2),(2,len(scores)+1)).value = list(scores.values())
            wb1.save(r'../../xlsx/scores.xlsx')
            wb1.close()
            app.quit()
            break

capture.release()
cv2.destroyAllWindows()
