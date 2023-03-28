import cv2 as cv
from matplotlib import pyplot as plt

from utils import get_complete_card, getStuID, getAnswers


def getMsg(img):
    # 获取去除边缘之后的答题卡
    totalCard = get_complete_card(img)

    # 获取水平垂直定位点坐标
    cv.imshow('totalCard', totalCard)
    # stuID = getStuID(totalCard, 8)
    #
    # print('学号: ', stuID)
    # # 获取每个题的选项个数
    # optNumOfSelQList = []
    # f = open("../doc/optNumOfSelQList.txt", "r")
    # lines = f.readlines()
    # for i in range(len(lines)):
    #     optNumOfSelQList.append(int(lines[i].strip('\n')))
    #
    # # 获取每个题的坐标
    # cors = []
    # f1 = open("../doc/cors", "r")
    # lines = f1.readlines()
    # for i in range(len(lines)):
    #     tmp = lines[i].strip('\n').split(' ')
    #     cors.append([int(tmp[0]), int(tmp[1])])  # 删除\n
    #
    # # 获取答案
    # answers = getAnswers(totalCard, optNumOfSelQList, cors)
    #
    # for i in range(len(answers)):
    #     print('第', i + 1, '题的答案是: ', answers[i])
    # print(len(optNumOfSelQList))
    # totalPlt = 1
    # plt.subplot(1, totalPlt, 1), plt.imshow(totalCard, cmap=plt.cm.gray), plt.title('totalCard')
    # plt.xticks([]), plt.yticks([])
    #
    # plt.show()
    cv.waitKey(0)
    cv.destroyAllWindows()


if __name__ == '__main__':
    img = cv.imread('../card/card_with_QRcode/test12.jpg')
    getMsg(img)
