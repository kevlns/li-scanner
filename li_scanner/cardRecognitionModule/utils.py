import cv2 as cv
import numpy as np
import pandas as pd
from imutils.perspective import four_point_transform
from scipy.stats import t


def findRectangles(img, contours_sort_with_Area, threshold=0.1):
    # 所有矩形
    recs = []
    # 矩形的角点
    recsCornerPoints = []
    # 矩形轮廓在所有轮廓中的索引
    recsIdx = []
    i = 0
    # 遍历所有轮廓
    for c in contours_sort_with_Area:
        # 获取周长
        peri = threshold * cv.arcLength(c, True)
        # 获取近似矩形
        approx = cv.approxPolyDP(c, peri, True)
        # 如果是矩形(有四个点)
        if len(approx) == 4:
            # 透视变换
            recCnt = four_point_transform(img, approx.reshape(4, 2))
            # print('type:',type(approx.reshape(4, 2)))
            recs.append(recCnt)
            recsCornerPoints.append(approx.reshape(4, 2))
            recsIdx.append(i)
        i = i + 1
    # 返回矩形,矩形的四个角点,矩形在所有轮廓中的索引
    return recs, recsCornerPoints, recsIdx


def get_complete_card(img):
    """
    获取根据大定位点切割后图像
    :param img: 输入RGB图像
    :return: 根据大定位点切割后的灰度图像
    """
    # 旋转90度
    # img = np.rot90(img, k=-1)
    img = np.rot90(img)
    # 转成灰度图
    gray = cv.cvtColor(img, cv.COLOR_BGR2GRAY)
    # blur0 = cv.blur(gray, (5, 5))

    blur0 = cv.medianBlur(gray, 3)
    # 二值化
    ret, th1 = cv.threshold(blur0, 100, 255, cv.THRESH_BINARY)
    # th1 = cv.adaptiveThreshold(blur0,255,cv.ADAPTIVE_THRESH_GAUSSIAN_C,cv.THRESH_BINARY_INV,11,2)
    # 形态学腐蚀
    blur1 = cv.blur(th1, (5, 5))
    # 再次二值化
    ret, th2 = cv.threshold(blur1, 150, 255, cv.THRESH_BINARY_INV)
    # th2 = cv.adaptiveThreshold(blur1, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY_INV, 11, 2)
    # blur2 = cv.blur(th2, (5, 5))
    # ret, th3 = cv.threshold(blur2, 100, 255, cv.THRESH_BINARY_INV)
    # cv.imshow('th3', th1)
    # 返回所有轮廓
    th2, contours, hierarchy = cv.findContours(th2, cv.RETR_TREE, cv.CHAIN_APPROX_SIMPLE)
    contours_sort_with_Area = sorted(contours, key=cv.contourArea, reverse=True)
    # 获取矩形
    recs, recsPoints, recsIdx = findRectangles(th2, contours_sort_with_Area)
    # print('共有矩形: ', len(recsPoints))

    rgb = cv.cvtColor(th2, cv.COLOR_GRAY2RGB)

    # 记录大定位点的坐标
    pointers = []
    # 将所有矩形的坐标转换成[左上角，右上角，右下角，左下角]的形式
    for i in range(len(recsPoints)):
        recsPoints_copy = []
        # 左上角,横纵坐标相加最小
        t = sorted(recsPoints[i], key=lambda x: x[0] + x[1])
        recsPoints_copy.append(t[0])
        # 右上角,横纵坐标相加位于中间两个，通过比较X值判断
        recsPoints_copy.append(t[1] if t[1][0] > t[2][0] else t[2])
        # 右下角,横纵坐标相加最小
        recsPoints_copy.append(t[3])
        # 左下角,横纵坐标相加位于中间两个，通过比较X值判断
        recsPoints_copy.append(t[1] if t[1][0] < t[2][0] else t[2])
        w1 = recsPoints_copy[1][0] - recsPoints_copy[0][0]
        w2 = recsPoints_copy[2][0] - recsPoints_copy[3][0]
        h1 = recsPoints_copy[3][1] - recsPoints_copy[0][1]
        h2 = recsPoints_copy[2][1] - recsPoints_copy[1][1]

        if i == 0 or i == 1 or i == 2 or i == 3:
            # print(i, '大定位点的长宽是: ', w1, w2, h1, h2, (w1 + w2) / (h1 + h2), (w1 + w2) * (h1 + h2) / 4)
            pointers.append(recsPoints_copy)
        # print(i, '大定位点的长宽是: ', w1, w2, h1, h2, (w1 + w2) / (h1 + h2), (w1 + w2) * (h1 + h2) / 4)
        #
        #         # 用长宽比和最大最小面积筛选大定位点
        # if h1 + h2 != 0 and 1.8 > (w1 + w2) / (h1 + h2) > 1.3 and 7000 > (w1 + w2) * (h1 + h2) / 4 > 3500:

    for i in pointers:
        w1 = i[1][0] - i[0][0]
        w2 = i[2][0] - i[3][0]
        h1 = i[3][1] - i[0][1]
        h2 = i[2][1] - i[1][1]
        # or not (800 > (w1 + w2) * (h1 + h2) / 4 > 300)
        if (h1 + h2) == 0:
            h1  = 1
        if not (2.0 > (w1 + w2) / (h1 + h2) > 1.3):
            # print('大定位点不完整，请调整角度!!!')
            return None

    idd = 0
    cv.rectangle(rgb, (pointers[idd][0][0], pointers[idd][0][1]), (pointers[idd][2][0], pointers[idd][2][1]),
                 (0, 255, 0), -1)
    cv.rectangle(rgb, (pointers[1][0][0], pointers[1][0][1]), (pointers[1][2][0], pointers[1][2][1]), (0, 255, 0), -1)
    cv.rectangle(rgb, (pointers[2][0][0], pointers[2][0][1]), (pointers[2][2][0], pointers[2][2][1]), (0, 255, 0), -1)
    cv.rectangle(rgb, (pointers[3][0][0], pointers[3][0][1]), (pointers[3][2][0], pointers[3][2][1]), (0, 255, 0), -1)

    # cv.imshow('rgb', rgb)

    # 绘制矩形
    # idx = 4
    # print(recsPoints[1], cv.contourArea(contours_sort_with_Area[recsIdx[1]]))
    # print(pointers[0])

    # 找到整个答题卡区域的角点
    # 左上角横纵坐标相加最小
    top_left = sorted(pointers, key=lambda x: x[0][0] + x[0][1])[0][0]
    # 右下角横纵坐标相加最大
    bottom_right = sorted(pointers, key=lambda x: x[2][0] + x[2][1], reverse=True)[0][2]
    # 右上角,优先判断y坐标
    tmp1 = sorted([pointers[0][1], pointers[1][1], pointers[2][1], pointers[3][1]], key=lambda x: x[1])
    top_right = tmp1[0] if tmp1[0][0] > tmp1[1][0] else tmp1[1]
    # 左下角, 优先断y坐标
    tmp2 = sorted([pointers[0][3], pointers[1][3], pointers[2][3], pointers[3][3]], key=lambda x: x[1], reverse=True)
    bottom_left = tmp2[0] if tmp2[0][0] < tmp2[1][0] else tmp2[1]

    # print('top_left', top_left)
    # print('top_right', top_right)
    # print('bottom_right', bottom_right)
    # print('bottom_left', bottom_left)
    # 将角点封装成数组
    bigCornerPointers = np.array([top_left, top_right, bottom_right, bottom_left])

    # 去除大定位点外围之后的答题卡,并且映射成规则矩形
    totalCard = four_point_transform(gray, bigCornerPointers)

    return totalCard


def grubbs(x, alpha=0.95):
    """
    去除离群值函数
    :param x: 待筛选的列表
    :param alpha: 该参数越小，筛选效果越明显
    :return: 删除离群值之后的列表，被删除元素在原列表中的索引
    """
    if isinstance(x, pd.Series) or isinstance(x, pd.DataFrame):
        x = x.astype('float').values
    elif isinstance(x, list):
        x = np.array(x)
    delIdx = []
    # 样本个数
    p = len(x)
    beta = 1 - alpha
    while True:
        # 格拉布斯法算出离群值
        if p > 2:
            # 求均值和方差
            mean = np.mean(x)
            std = np.std(x)
            if std == 0:
                std = 1
            g_arr = np.abs(x - mean) / std
            # 最有可能是离群值的样本的 index
            g_index = g_arr.argmax()
            # 求出 Grubbs test 的能力统计量 G
            G = g_arr[g_index]
            t_crital = t.ppf(beta / (2 * p), p - 2)
            # 求检验统计量在显著性水平中的临界值
            criteria = (p - 1) / np.sqrt(p) * np.sqrt(t_crital ** 2 / (p - 2 + t_crital ** 2))
            if G > criteria:
                # 若样本中有离群值, 删除离群值
                x = np.delete(x, g_index)
                delIdx.append(g_index)
                # 重新求取参加者个数
                p = len(x)
            else:
                # 若样本中没有离群值，则返回
                return x, delIdx
        else:
            return x, delIdx


def remove_outliers(pointers):
    """
    筛选定位点
    :param pointers: 待筛选的定位点列表
    :return: 筛选之后的定位点
    """
    # print('筛选之前:', len(pointers))
    # 根据面积去除异常值
    # area = []
    # for i in pointers:
    #     w1 = i[1][0] - i[0][0]
    #     w2 = i[2][0] - i[3][0]
    #     h1 = i[3][1] - i[0][1]
    #     h2 = i[2][1] - i[1][1]
    #     area.append(((w1 + w2) * (h1 + h2) / 4))
    # area, delIdx = grubbs(area, 0.95)
    # for i in delIdx:
    #     pointers.pop(i)
    # print('面积筛选过后', len(pointers))
    # 根据周长去除异常值
    peri = []
    for i in pointers:
        w1 = i[1][0] - i[0][0]
        w2 = i[2][0] - i[3][0]
        h1 = i[3][1] - i[0][1]
        h2 = i[2][1] - i[1][1]
        peri.append(((w1 + w2) / 2 + (h1 + h2) / 2))
    peri, delIdx = grubbs(peri, 0.97)
    for i in delIdx:
        pointers.pop(i)
    # print('周长筛选过后', len(pointers))

    return pointers


def get_small_dots(img):
    """
    获取小的定位点
    :param img: 根据大定位点切割后的灰度图
    :return: 水平定位点，垂直定位点
    """
    # 二值化
    ret, th1 = cv.threshold(img, 100, 255, cv.THRESH_BINARY)
    # 形态学腐蚀
    blur1 = cv.blur(th1, (3, 3))
    # 再次二值化
    ret, th2 = cv.threshold(blur1, 100, 255, cv.THRESH_BINARY)
    blur2 = cv.blur(th2, (5, 5))

    # 返回所有轮廓
    blur2, contours, hierarchy = cv.findContours(blur1, cv.RETR_TREE, cv.CHAIN_APPROX_SIMPLE)
    # cv.imshow('small', blur1)
    # 按面积大小排序
    contours_sort_with_Area = sorted(contours, key=cv.contourArea, reverse=True)
    recsPoints = []
    for cnt in contours_sort_with_Area:
        x, y, w, h = cv.boundingRect(cnt)
        recsPoints.append([[x, y], [x + w, y], [x + w, y + h], [x, y + h]])

    # 获取矩形
    # recs, recsPoints, recsIdx = findRectangles(th2, contours_sort_with_Area, 0.05)
    # print('共有矩形: ', len(recsPoints))

    rgb = cv.cvtColor(blur2, cv.COLOR_GRAY2RGB)

    # 记录小定位点的坐标
    pointers = []
    # 将所有矩形的坐标转换成[左上角，右上角，右下角，左下角]的形式
    for i in range(len(recsPoints)):
        recsPoints_copy = []
        # 左上角,横纵坐标相加最小
        t = sorted(recsPoints[i], key=lambda x: x[0] + x[1])
        recsPoints_copy.append(t[0])
        # 右上角,横纵坐标相加位于中间两个，通过比较X值判断
        recsPoints_copy.append(t[1] if t[1][0] > t[2][0] else t[2])
        # 右下角,横纵坐标相加最小
        recsPoints_copy.append(t[3])
        # 左下角,横纵坐标相加位于中间两个，通过比较X值判断
        recsPoints_copy.append(t[1] if t[1][0] < t[2][0] else t[2])
        pointers.append(recsPoints_copy)

    for i in range(1, len(pointers)):
        cv.rectangle(rgb, (pointers[i][0][0], pointers[i][0][1]), (pointers[i][2][0], pointers[i][2][1]),
                     (0, 0, 255), -1)
    # cv.imshow('rgb', rgb)
    # 水平和垂直阈值
    Threshold_x = 0
    Threshold_y = 0
    for i in range(blur1.shape[1]):
        if blur1[5, i] > 100:
            Threshold_x = i
            break
    for i in range(blur1.shape[0]):
        if blur1[i, 5] > 100:
            Threshold_y = i
            break
    # print('Threshold_x: ',Threshold_x,'Threshold_y: ',Threshold_y)
    # 水平小定位点
    horizontalDots = []
    # 垂直小定位点
    verticalDots = []
    # 用垂直坐标筛选水平定位点
    for i in pointers:
        if i[2][1] < Threshold_y + 5:
            horizontalDots.append(i)

    # 用水平坐标筛选垂直定位点
    for i in pointers:
        if i[2][0] < Threshold_x + 5:
            verticalDots.append(i)
    # print('horizontalDots:',len(horizontalDots),'verticalDots',len(verticalDots))
    # 去除异常值
    horizontalDots = remove_outliers(horizontalDots)
    # 根据左上角排序
    horizontalDots = sorted(horizontalDots, key=lambda x: x[0][0])
    # 去除异常值
    verticalDots = remove_outliers(verticalDots)
    # 根据右上角排序
    verticalDots = sorted(verticalDots, key=lambda x: x[1][1])

    horizontalPointers = []
    verticalPointers = []
    for i in horizontalDots:
        horizontalPointers.append((i[3][0], i[2][0]))
    # print('水平定位点个数: ', len(horizontalPointers))

    for i in verticalDots:
        verticalPointers.append((i[1][1], i[2][1]))
    # print('垂直定位点个数: ', len(verticalPointers), ' 水平定位点个数: ', len(horizontalPointers))
    if len(verticalPointers) != 16 or len(horizontalPointers) != 20:
        # print('小定位点不完整，请调整角度!!!')
        return [], []
    return horizontalPointers, verticalPointers


def getStuID(img, idDigits):
    """
    获取学号
    :param img: 根据大定位点切割后的灰度图
    :param idDigits: 学号位数
    :return: 学号列表
    """
    horizontalPointers, verticalPointers = get_small_dots(img)
    if len(horizontalPointers) == 0 or len(verticalPointers) == 0:
        return None
    # if len(horizontalPointers) != 20 or len(verticalPointers) != 25:
    #     return
    # 二值化
    ret, th1 = cv.threshold(img, 0, 255, cv.THRESH_BINARY + cv.THRESH_OTSU)
    # 形态学腐蚀
    blur1 = cv.blur(th1, (5, 5))
    # 再次二值化
    ret, th2 = cv.threshold(blur1, 100, 255, cv.THRESH_BINARY_INV)
    blur2 = cv.blur(th2, (5, 5))

    stuID = []
    for i in range(5, 5 + idDigits):
        white_dots = 0
        idx = 0
        for j in range(0 + 1, 10 + 1):
            pic = blur2[verticalPointers[j][0]:verticalPointers[j][1],
                  horizontalPointers[i][0]:horizontalPointers[i][1]]
            tmp = cv.countNonZero(pic)
            if white_dots < tmp:
                white_dots = tmp
                idx = j - 1
        stuID.append(idx)
    return stuID


def getAnswers(img, optNumOfSelQList, originalCors):
    """
    获取答案
    :param img: 根据大定位点切割后的灰度图
    :param optNumOfSelQList: 每道题的选项个数
    :param originalCors: 选择题题的坐标
    :return: 返回答案列表
    """
    cors = np.copy(originalCors)
    horizontalPointers, verticalPointers = get_small_dots(img)
    if len(horizontalPointers) == 0 or len(verticalPointers) == 0:
        return None
    y = set()
    for i in cors:
        y.add(i[0])
    # print('长度',len(y),len(verticalPointers))
    if len(y) + 10 + 1 != len(verticalPointers):
        # print('请调整角度!!!')
        return

        # 二值化
    ret, th1 = cv.threshold(img, 0, 255, cv.THRESH_BINARY + cv.THRESH_OTSU)
    # 形态学腐蚀
    blur1 = cv.blur(th1, (5, 5))
    # 再次二值化
    ret, th2 = cv.threshold(blur1, 100, 255, cv.THRESH_BINARY_INV)
    blur2 = cv.blur(th2, (5, 5))
    # 调整定位点坐标偏移量
    for i in range(len(cors)):
        cors[i][0] = cors[i][0] - 5
        cors[i][1] = cors[i][1] - 3

    offset = 0
    # 答案
    answers = []

    for i in range(len(cors)):
        if cors[i][1] == 0 and i > 5 and i % 5 == 0:
            offset = offset - 1
        idx = []
        # print('cors: ',cors)
        for j in range(optNumOfSelQList[i]):
            # print('偏移',cors[i][0] + offset,j+cors[i][1])
            pic = blur2[verticalPointers[cors[i][0] + offset][0]:verticalPointers[cors[i][0] + offset][1],
                  horizontalPointers[j + cors[i][1]][0]:horizontalPointers[j + cors[i][1]][1]]
            tmp = cv.countNonZero(pic)
            if tmp / (pic.shape[0] * pic.shape[1]) > 0.4:
                idx.append(j + 1)

        answers.append(idx)
    for i in answers:
        if len(i) == 0:
            i.append(0)

    # cv.waitKey(0)
    # cv.destroyAllWindows()
    return answers


def SubjectiveSegmentation(img, length):
    """
    :param img: 需要处理的图片
    :param length: (list),每道大题行坐标
    :return: void
    """

    # 转灰度图
    # gray = cv.cvtColor(img, cv.COLOR_BGR2GRAY)
    # cv.imshow("img",img)

    # 局部直方图均衡化
    clahe_1 = cv.createCLAHE(clipLimit=2.0, tileGridSize=(15, 15))
    equal = clahe_1.apply(img)

    # 伽马矫正
    Cimg = equal / 255
    gamma = 1.5
    out_1 = np.power(Cimg, gamma)
    out_1 = out_1 * 255
    out_1 = out_1.astype(np.uint8)
    # cv.imshow("out1",out_1)

    # 直方图正规化
    Maximg = np.max(equal)
    Minimg = np.min(equal)
    Omin, Omax = 0, 255
    a = float(Omax - Omin) / (Maximg - Minimg)
    b = Omin - a * Minimg
    enhanced = a * equal + b
    enhanced = enhanced.astype(np.uint8)
    gamma_1 = out_1

    # 提取大题边缘
    Gauss = cv.GaussianBlur(gamma_1, (5, 5), 0)
    edged = cv.Canny(~Gauss, 100, 255)

    # 获取矩形框
    contours = cv.findContours(edged, cv.RETR_LIST, cv.CHAIN_APPROX_SIMPLE)[1]
    rect = []
    rect_area = []
    for c in contours:
        x, y, w, h = cv.boundingRect(c)
        area = w * h
        rect_area.append([area, x, x + w, y, y + h])
        # print("123")
        # if area > 1000.0:
        # temp_rect = [x, x + w, y, y + h]
        # if temp_rect not in rect:
    # print(rect_area)
    # 筛选最外面矩形
    rect_area.sort()
    # print(rect_area)
    Max_rect = [rect_area[-1][1], rect_area[-1][2], rect_area[-1][3], rect_area[-1][4]]

    # print(Max_rect)
    # 切出大题区域
    img_1 = cv.drawContours(img, contours, -1, (0, 255, 0), 3)
    # cv.imshow("img1", img_1)
    img_cut_0 = enhanced[Max_rect[2]:Max_rect[3], Max_rect[0]:Max_rect[1]]

    # 增强图片对比度及亮度
    enhanced_cut = np.uint8(np.clip((1.2 * img_cut_0 + 20), 0, 255))
    Cimg = enhanced_cut / 255
    gamma = 2
    out_2 = np.power(Cimg, gamma)
    out_2 = out_2 * 255
    out_2 = out_2.astype(np.uint8)
    enhanced_cut = cv.addWeighted(enhanced_cut, 0.65, out_2, 0.45, 0)

    # 将图片放大两倍
    cut_resize = cv.resize(enhanced_cut, (enhanced_cut.shape[1] * 2,
                                          enhanced_cut.shape[0] * 2))

    # 计算每道题分割高度
    total_length = 62 * (length[0] // 62 + 1) - length[0] - 1
    Max_length = 6 * len(length) + 2
    temp_length = np.copy(length)
    temp_length = list(temp_length)
    if total_length > Max_length:
        temp_length.append(length[-1] + 6)
    ever_len = []  # 每道题目长度
    trans_length = []
    for row_index in range(0, len(temp_length) - 1):
        ever_len.append(temp_length[row_index + 1] - temp_length[row_index])
    ever_len.append(62 * (temp_length[0] // 62 + 1) - temp_length[-1])
    cr_high, cr_wide = cut_resize.shape
    for row_length in ever_len:
        trans_length.append(int(row_length / total_length * cr_high + 0.99))

    # 分割题目
    present_y = 0
    pics = []
    c_para = int(0.025 * cr_high)  # corrected parameter
    for num_question in range(0, len(length)):
        q_y1 = 0 if present_y - c_para < 0 else present_y - c_para
        q_y2 = cr_high if present_y + trans_length[num_question] + c_para > cr_high else \
            present_y + trans_length[num_question] + c_para
        # print(q_y1, present_y - c_para, q_y2, present_y + c_para)
        img_cut_ever = cut_resize[q_y1:q_y2, 0:cr_wide]

        pics.append(img_cut_ever)

        # cv.imshow(f"cut_{num_question}", img_cut_ever)

        present_y += trans_length[num_question]

        # cv.waitKey(0)
        # cv.destroyAllWindows()
    # if total_length > actual_length:
        # print (total_length, 6 * len(length) + 2, len(length))
        # return pics[:len(length)]

    return pics