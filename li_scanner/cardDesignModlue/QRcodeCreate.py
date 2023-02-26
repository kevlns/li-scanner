# -*- coding: utf-8 -*-
import qrcode
import os
import xlwings as xw

QRMsg = 'U202142008'
# 生成二维码
img = qrcode.make(QRMsg)
picName = '../card/QRcode_pic/'+QRMsg + ".png"
img.save(picName)

wb = xw.Book('../card/card_template.xlsx')  # 打开模板答题卡
sht = wb.sheets['Sheet1']
fileName = os.path.join(os.getcwd(), picName)
width, height = 35, 35  # 指定图片大小

# 正面
rng = sht.range(6, 25)  # 目标单元格
left = rng.left
top = rng.top
sht.pictures.add(fileName, left=left, top=top, width=width, height=height)
# 反面
rng = sht.range(68, 25)  # 目标单元格
left = rng.left
top = rng.top
sht.pictures.add(fileName, left=left, top=top, width=width, height=height)
wb.save()
wb.close()
