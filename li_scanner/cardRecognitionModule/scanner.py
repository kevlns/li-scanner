import sys

import cv2

from li_scanner.cardRecognitionModule.utils import getStuID, get_complete_card, getAnswers

capture = cv2.VideoCapture(0)
StuID = []
cishu = 0
while capture.isOpened():
    ret, frame = capture.read()
    if ret:
        cv2.imshow("show", frame)

        score = cv2.Laplacian(frame, cv2.CV_64F).var()
        if score > 1000:
            totalCard = get_complete_card(frame)
            if totalCard is not None:
                # 获取水平垂直定位点坐标
                # cv.imshow('totalCard',totalCard)
                tmp = getStuID(totalCard, 8)
                if StuID == tmp and tmp is not None:
                    cishu = cishu + 1
                    print('-'*10,'第',cishu ,'次的学号是',StuID,'-'*10)
                else:
                    StuID = tmp
                    cishu = 0
                if cishu>=2:
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

                    # 获取答案
                    answers = getAnswers(totalCard, optNumOfSelQList, cors)

                    for i in range(len(answers)):
                        print('第', i + 1, '题的答案是: ', answers[i])
                    # print(len(optNumOfSelQList))
                    sys.exit()

        # 按下 q 键可退出程序执行
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break
        # if cv2.waitKey(1) & 0xFF == ord('a'):
        #     cv2.imwrite('../../card/card_with_QRcode/test14.jpg',frame)

capture.release()
cv2.destroyAllWindows()
