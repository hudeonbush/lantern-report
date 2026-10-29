# -*- coding: utf-8 -*-
"""思政说课 1:1 复刻 — 母版背景合成 + 照片按框中心裁剪 + 红色图表生成"""
import os, math
from PIL import Image, ImageDraw, ImageFilter

# 跨电脑兼容：基于脚本自身位置自动定位，不写任何绝对路径
HERE = os.path.dirname(os.path.abspath(__file__))   # 本脚本所在 work/ 目录
FIG  = os.path.join(HERE, "figs")
OUT  = os.path.join(HERE, "bgs")
os.makedirs(OUT, exist_ok=True)
CW, CH = 1920, 1080          # 2x of 960x540

def resize_cover(im, w=CW, h=CH):
    iw, ih = im.size
    s = max(w/iw, h/ih)
    im2 = im.resize((int(iw*s)+1, int(ih*s)+1), Image.LANCZOS)
    x = (im2.width - w)//2; y = (im2.height - h)//2
    return im2.crop((x, y, x+w, y+h))

# ---------- P1 封面背景 ----------
cover = Image.open(f"{FIG}/cover_bg.png").convert("RGB")
resize_cover(cover).save(f"{OUT}/bg_cover.png")

# ---------- 内容页母版：白底 + 顶部红渐变导航带 + 底部红波浪 ----------
bg = Image.new("RGB", (CW, CH), (255, 252, 250))
d = ImageDraw.Draw(bg)

# 内容区淡淡的暖白光晕
glow = Image.new("L", (CW, CH), 0)
dg = ImageDraw.Draw(glow)
dg.ellipse([200, 160, 1720, 900], fill=14)
glow = glow.filter(ImageFilter.GaussianBlur(80))
bg = Image.composite(Image.new("RGB", (CW, CH), (253, 238, 232)), bg, glow)
d = ImageDraw.Draw(bg)

# 顶部红渐变带 y 0..112px (56pt)，底缘微波
band = Image.new("RGB", (CW, 118), 0)
bd = ImageDraw.Draw(band)
for x in range(CW):
    t = x / CW
    r = int(146 + 40*t); g = int(22 + 26*t); b = int(26 + 12*t)   # 92161A -> B24B32? 保持红
    r = min(r, 190)
    bd.line([(x,0),(x,118)], fill=(r, g, b))
bg.paste(band, (0, 0))
# 带底缘弧线(浅金细线)
for x in range(CW):
    y = 112 + int(4*math.sin(x/CW*math.pi*2))
    d.point((x, y), fill=(232, 184, 75))
    d.point((x, y+1), fill=(232, 184, 75))

# 底部红波浪 y>=1030px (515pt)
for x in range(CW):
    y0 = 1032 + int(16*math.sin(x/CW*3.0) + 8*math.sin(x/CW*7.1+1.3))
    t = x/CW
    col = (int(150+40*t), int(24+18*t), int(28+8*t))
    d.line([(x, y0), (x, CH)], fill=col)

bg.save(f"{OUT}/bg_content.png")
print("masters done")

# ---------- 照片中心裁剪到目标纵横比 ----------
def crop_to(src, aspect, out, w=800):
    im = Image.open(src).convert("RGB")
    iw, ih = im.size
    cur = iw/ih
    if cur > aspect:      # 太宽，裁左右
        nw = int(ih*aspect); x = (iw-nw)//2
        im = im.crop((x, 0, x+nw, ih))
    else:                 # 太高，裁上下(偏上保留主体)
        nh = int(iw/aspect); y = max(0, int((ih-nh)*0.35))
        im = im.crop((0, y, iw, y+nh))
    im = im.resize((w, int(w/aspect)), Image.LANCZOS)
    im.save(out, quality=92)

jobs = [
    # P2 三张小横照 (2.86)
    (f"{FIG}/class1.png", 2.86, f"{FIG}/c1_w.png"),
    (f"{FIG}/class2.png", 2.86, f"{FIG}/c2_w.png"),
    (f"{FIG}/class3.png", 2.86, f"{FIG}/c3_w.png"),
    # P2 教材 (0.60)
    (f"{FIG}/textbook.png", 0.60, f"{FIG}/tb.png"),
    # P3 两张调研照 (1.40)
    (f"{FIG}/class1.png", 1.40, f"{FIG}/c1_m.png"),
    (f"{FIG}/class2.png", 1.40, f"{FIG}/c2_m.png"),
    # P4 圆形照 (1.0)
    (f"{FIG}/class1.png", 1.0, f"{FIG}/c1_sq.png"),
    (f"{FIG}/class2.png", 1.0, f"{FIG}/c2_sq.png"),
    (f"{FIG}/class3.png", 1.0, f"{FIG}/c3_sq.png"),
    # P4 小照 (1.50) 6张
    (f"{FIG}/class2.png", 1.50, f"{FIG}/c2_s.png"),
    (f"{FIG}/class3.png", 1.50, f"{FIG}/c3_s.png"),
    (f"{FIG}/class4.png", 1.50, f"{FIG}/c4_s.png"),
    # P5 三张 (1.74)
    (f"{FIG}/class1.png", 1.74, f"{FIG}/c1_p5.png"),
    (f"{FIG}/class2.png", 1.74, f"{FIG}/c2_p5.png"),
    (f"{FIG}/class4.png", 1.74, f"{FIG}/c4_p5.png"),
    # P5 华表 (0.33)
    (f"{FIG}/huabiao.png", 0.33, f"{FIG}/hb.png"),
    # P6 左照 (2.08) 右照 (2.15)
    (f"{FIG}/class2.png", 2.08, f"{FIG}/c2_p6.png"),
    (f"{FIG}/class3.png", 2.15, f"{FIG}/c3_p6.png"),
    # P6 中间 2x2 (1.41)
    (f"{FIG}/class1.png", 1.41, f"{FIG}/c1_g.png"),
    (f"{FIG}/class2.png", 1.41, f"{FIG}/c2_g.png"),
    (f"{FIG}/class3.png", 1.41, f"{FIG}/c3_g.png"),
    (f"{FIG}/class4.png", 1.41, f"{FIG}/c4_g.png"),
    # P7 (2.44 / 2.0 / 2.0)
    (f"{FIG}/class3.png", 2.44, f"{FIG}/c3_p7.png"),
    (f"{FIG}/class1.png", 2.0,  f"{FIG}/c1_p7.png"),
    (f"{FIG}/class2.png", 2.0,  f"{FIG}/c2_p7.png"),
    # P8 (1.44) x2
    (f"{FIG}/class1.png", 1.44, f"{FIG}/c1_p8.png"),
    (f"{FIG}/class2.png", 1.44, f"{FIG}/c2_p8.png"),
]
for src, asp, out in jobs:
    crop_to(src, asp, out)
print("crops done:", len(jobs))

# ---------- P3 红色统计图 (matplotlib) ----------
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei"]
plt.rcParams["axes.unicode_minus"] = False

RD = "#B01F24"; RD2 = "#E08A80"; GY = "#666666"

def style(ax):
    for sp in ["top", "right"]: ax.spines[sp].set_visible(False)
    for sp in ["left", "bottom"]: ax.spines[sp].set_color("#CC8878")
    ax.tick_params(colors=GY, labelsize=8)
    ax.set_facecolor("white")

# 图1: 思想政治基础水平分布
fig, ax = plt.subplots(figsize=(5.8, 2.0), dpi=200)
cats = ["优秀", "良好", "中等", "待提升"]
v1 = [18, 34, 30, 18]
bars = ax.bar(cats, v1, width=0.52, color=[RD, "#C74A44", RD2, "#EDC0BA"], zorder=3)
for b, v in zip(bars, v1):
    ax.text(b.get_x()+b.get_width()/2, v+1, f"{v}%", ha="center", fontsize=9, color=RD, fontweight="bold")
ax.set_ylim(0, 42); ax.set_yticks([])
style(ax)
fig.patch.set_alpha(0)
fig.tight_layout()
fig.savefig(f"{FIG}/chart1.png", transparent=True)
plt.close(fig)

# 图2: 素养变化趋势 (柱+折线)
fig, ax = plt.subplots(figsize=(6.0, 2.0), dpi=200)
terms = ["开学初", "期中", "期末", "结课"]
v2 = [42, 55, 68, 81]
bars = ax.bar(terms, v2, width=0.5, color=[RD2, "#D96A62", RD, "#8E1414"], zorder=3)
ax.plot(terms, v2, color="#E8B84B", marker="o", lw=2.2, ms=5, zorder=4)
for x, v in zip(range(4), v2):
    ax.text(x, v+3.5, f"{v}%", ha="center", fontsize=9, color=RD, fontweight="bold")
ax.set_ylim(0, 100); ax.set_yticks([])
style(ax)
fig.patch.set_alpha(0)
fig.tight_layout()
fig.savefig(f"{FIG}/chart2.png", transparent=True)
plt.close(fig)
print("charts done")
