# -*- coding: utf-8 -*-
"""新背景图处理：去水印 -> 裁 16:9 -> 封面版 + 内容页柔白蒙版版"""
import os
from PIL import Image, ImageDraw, ImageFilter

# 跨电脑兼容：基于脚本自身位置自动定位，不写任何绝对路径
HERE = os.path.dirname(os.path.abspath(__file__))   # work/ 目录
SRC = os.path.join(HERE, "new_bg.jpg")             # 把新背景原图放到 work/new_bg.jpg 再运行
BGS = os.path.join(HERE, "bgs")

im = Image.open(SRC).convert('RGB')
w, h = im.size
print("原图:", w, h)

# 1) 去水印：右下角白字 "AI生成 WORKBUDDY"，用左侧红绸 patch 覆盖（左缘羽化）
wx0, wy0, wx1, wy1 = 1378, 938, 1536, 1024
pw = wx1 - wx0
patch = im.crop((wx0 - pw, wy0, wx0, wy1))
mask = Image.new('L', patch.size, 255)
d = ImageDraw.Draw(mask)
for i in range(30):
    d.line([(i, 0), (i, patch.size[1])], fill=int(255 * i / 30))
im.paste(patch, (wx0, wy0), mask)
print("水印已去除:", (wx0, wy0, wx1, wy1))

# 2) 裁剪 16:9：检测华表金色最高点，顶部少切保狮子，其余切底部
px = im.load()
gold_min_y = h
for y in range(0, 300, 2):
    hit = False
    for x in range(860, 1250, 3):
        r, g, b = px[x, y]
        if r > 170 and g > 115 and b < 150 and r - b > 55:
            hit = True
            break
    if hit:
        gold_min_y = y
        break
top = max(0, gold_min_y - 14)
if top + 864 > h:
    top = h - 864
im2 = im.crop((0, top, w, top + 864)).resize((960, 540), Image.LANCZOS)
im2.save(BGS + "/bg_cover.png")
print("封面版完成: gold_min_y=%d top=%d" % (gold_min_y, top))

# 3) 内容页：柔边白蒙版盖"标题+内容区"(85,46)-(952,460)，四边保留氛围
white = Image.new('RGB', im2.size, (255, 255, 255))
m = Image.new('L', im2.size, 0)
d = ImageDraw.Draw(m)
d.rectangle((85, 46, 952, 460), fill=215)
m = m.filter(ImageFilter.GaussianBlur(36))
im3 = Image.composite(white, im2, m)
im3.save(BGS + "/bg_content.png")
print("内容页版完成")

# 预览图（输出到脚本同目录，跨平台可用）
im2.save(os.path.join(HERE, "new_bg_cover_preview.png"))
im3.save(os.path.join(HERE, "new_bg_content_preview.png"))
