import cv2 as cv
from matplotlib import pyplot as plt

# 用普通模式加载图像
from utils import get_complete_card, get_small_dots, getStuID, getAnswers

img = cv.imread('../card/card_with_QRcode/test3.jpg')
# 获取去除边缘之后的答题卡
totalCard = get_complete_card(img)
# 获取水平垂直定位点坐标

stuID = getStuID(totalCard, 8)

print('学号: ', stuID)

optNumOfSelQList = []
f = open("../doc/optNumOfSelQList", "r")
lines = f.readlines()
cors = []
for i in range(len(lines)):
    optNumOfSelQList.append(int(lines[i].strip('\n')))


answers = getAnswers(totalCard, optNumOfSelQList)
# print(answers)

for i in range(len(answers)):
    print('第', i + 1, '题的答案是: ', answers[i])
# print(len(optNumOfSelQList))
totalPlt = 1
plt.subplot(1, totalPlt, 1), plt.imshow(totalCard, cmap=plt.cm.gray), plt.title('totalCard')
plt.xticks([]), plt.yticks([])

plt.show()
cv.waitKey(0)
cv.destroyAllWindows()
