# -*- coding: UTF8 -*-
import os
from PIL import Image
from docx import Document
from docx.shared import Pt

width_i = 1440
height_i = 3120
row_max = 2
line_max = 2
all_path = list()
num = 0
pic_max = line_max * row_max
dir_name = r"C:\Users\www\Desktop\out"

data_start = 8
data_end = 31

doc_path = "C:\\Users\\www\\Desktop\\out\\"

doc = Document()
# doc = Document(doc_path + "111.docx")

# root文件夹的路径  dirs 路径下的文件夹列表  files路径下的文件列表
for root, dirs, files in os.walk(dir_name):
    for file in files:
        if "jpg" in file:  # 子串在母串里面不
            all_path.append(os.path.join(root,file))

# all_path获取每张图片的绝对路径

toImage = Image.new('RGBA',(width_i*line_max,height_i*row_max))


for i in range(row_max):
    for j in range(line_max):
        # 每次打开图片绝对路路径列表的第一张图片
        pic_fole_head = Image.open(all_path[num])
        # 获取图片的尺寸
        wihth,height = pic_fole_head.size
        # 按照指定的尺寸，给图片重新赋值，<PIL.Image.Image image mode=RGB size=200x200 at 0x127B7978>
        tmppic = pic_fole_head.resize((width_i, height_i))
        # 计算每个图片的左上角的坐标点(0, 0)，(0, 1440)，(0, 6240)，(1440, 0)，(200, 200)。。。。(400, 400)
        loc = (int(i % line_max * width_i), int(j % line_max * height_i))
        print("第{}张图的存放位置".format(num),loc)
        toImage.paste(tmppic, loc)
        num = num + 1

        if num >= len(all_path):
            print("breadk")
            break
    if num >= pic_max:
        break

print(toImage.size)
toImage.save(r'C:\Users\www\Desktop\out\merged11.png')



doc.add_picture(doc_path + 'merged.jpg')
doc.save(doc_path + "111.docx")