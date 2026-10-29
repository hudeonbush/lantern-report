# -*- coding: utf-8 -*-
"""思政说课 PPT 1:1 复刻 — WPS COM 可见模式构建 8 页"""
import os, time, math
import win32com.client
import pythoncom

# 跨电脑兼容：基于脚本自身位置自动定位根目录，不写任何绝对路径
HERE = os.path.dirname(os.path.abspath(__file__))   # 本文件所在目录 work/
BASE = os.path.dirname(HERE)                         # 项目根目录（含本 md/txt 与 work/）
FIG  = os.path.join(BASE, "work", "figs")
BGS  = os.path.join(BASE, "work", "bgs")
OUT_PPTX = os.path.join(BASE, "思政说课_1比1复刻.pptx")
OUT_PDF  = os.path.join(BASE, "思政说课_1比1复刻.pdf")

FT = 'Microsoft YaHei'      # 正文
FX = '华文行楷'              # 封面书法
FK = 'Microsoft YaHei'      # 导航

# 配色（红金党建风）
RED   = (176, 31, 36)    # 主红 B01F24
REDD  = (142, 20, 20)    # 深红 8E1414
REDT  = (166, 27, 27)    # 标题深红 A61B1B
GOLD  = (232, 184, 75)   # 金 E8B84B
GOLDD = (216, 164, 65)   # 深金
ORG   = (232, 131, 58)   # 橙箭头
INK   = (51, 51, 51)     # 正文黑
G     = (120, 120, 120)  # 灰
WHITE = (255, 255, 255)
PINKL = (240, 201, 196)  # 浅红边框
CREAM = (253, 243, 240)  # 浅底

def h2b(c):
    return c[0] + c[1]*256 + c[2]*65536   # RGB -> BGR

TASKKILL = 'taskkill /F /IM wps.exe /IM wpscenter.exe /IM wpp.exe /IM et.exe /IM etx.exe >NUL 2>&1'
os.system(TASKKILL)
time.sleep(2.5)

def connect_wps(retries=4):
    for i in range(retries):
        try:
            app = win32com.client.Dispatch('KWPP.Application')
            app.Visible = True
            _ = app.Presentations.Count
            return app
        except Exception:
            pythoncom.CoUninitialize()
            time.sleep(2.5)
    raise RuntimeError("WPS COM 连接失败")

app = connect_wps()
pres = app.Presentations.Add()
pres.PageSetup.SlideWidth = 960
pres.PageSetup.SlideHeight = 540

def new_slide(idx, bg):
    s = pres.Slides.Add(idx, 12)
    try:
        s.Background.Fill.UserPicture(bg)
    except Exception:
        s.FollowMasterBackground = False
        s.Background.Fill.UserPicture(bg)
    try:
        app.ActiveWindow.View.GotoSlide(idx)
    except Exception:
        pass
    time.sleep(0.25)
    return s

def txt(s, x, y, w, h, text, fs=13, color=INK, bold=False, align=1, font=FT, italic=False):
    tb = s.Shapes.AddTextbox(1, x, y, w, h)
    tf = tb.TextFrame
    tf.WordWrap = True
    try:
        tf.MarginLeft = 0; tf.MarginRight = 0; tf.MarginTop = 0; tf.MarginBottom = 0
        tf.AutoSize = 0
        tf.VerticalAnchor = 3
    except Exception:
        pass
    tf.TextRange.Text = text
    f = tf.TextRange.Font
    f.Name = font; f.Size = fs
    f.Color.RGB = h2b(color)
    if bold: f.Bold = True
    if italic: f.Italic = True
    try:
        tf.TextRange.ParagraphFormat.Alignment = {1: 1, 2: 2, 3: 3}[align]
    except Exception:
        pass
    return tb

def shape(s, x, y, w, h, fill, line=None, lw=1.0, st=1):
    sp = s.Shapes.AddShape(st, x, y, w, h)
    sp.Fill.Solid()
    sp.Fill.ForeColor.RGB = h2b(fill)
    if line is None:
        sp.Line.Visible = False
    else:
        sp.Line.ForeColor.RGB = h2b(line)
        sp.Line.Weight = lw
    return sp

def shape_text(sp, text, fs, color, bold=True, font=FT):
    tf = sp.TextFrame
    try:
        tf.MarginLeft = 0; tf.MarginRight = 0; tf.MarginTop = 0; tf.MarginBottom = 0
        tf.AutoSize = 0
        tf.VerticalAnchor = 3
    except Exception:
        pass
    tf.WordWrap = True
    tf.TextRange.Text = text
    f = tf.TextRange.Font
    f.Name = font; f.Size = fs
    f.Color.RGB = h2b(color)
    if bold: f.Bold = True
    try:
        tf.TextRange.ParagraphFormat.Alignment = 2
    except Exception:
        pass

def pill(s, x, y, w, h, text, fs, fg, bg, bold=True, line=None, lw=1.0, font=FT):
    sp = shape(s, x, y, w, h, bg, line=line, lw=lw, st=5)
    shape_text(sp, text, fs, fg, bold=bold, font=font)
    return sp

def pic(s, path, x, y, w, h):
    return s.Shapes.AddPicture(os.path.normpath(path), 0, 1, x, y, w, h)

SECTIONS = ["课程介绍", "学情分析", "课前准备", "能力提升", "价值认同", "教学反思", "创新方向"]
TABS = ["教学整体设计", "教学实施过程", "学生学习效果", "反思改进措施"]

def chrome(s, sec_name, active_tab=None):
    """内容页公共部件：五角星 + 页标题 + 4 标签 + 左侧竖导航"""
    # 五角星（金星）
    try:
        star = shape(s, 26, 9, 46, 42, GOLD, st=92)
    except Exception:
        star = pill(s, 26, 9, 46, 42, "★", 24, WHITE, GOLD)
    # 页标题（深红胶囊 + 白字）
    pill(s, 84, 10, 196, 40, sec_name, 21, WHITE, REDD)
    # 顶部 4 标签
    tx0, tw, tgap, ty, th = 318, 140, 12, 14, 32
    for i, lab in enumerate(TABS):
        x = tx0 + i*(tw+tgap)
        if i == active_tab:
            pill(s, x, ty, tw, th, lab, 13, RED, WHITE, line=RED, lw=1.5)
        else:
            pill(s, x, ty, tw, th, lab, 13, INK, WHITE)
    # 左侧竖导航
    y0, nh, ngap = 190, 24, 8
    for i, lab in enumerate(SECTIONS):
        y = y0 + i*(nh+ngap)
        if lab == sec_name:
            pill(s, 8, y, 100, nh, lab, 12.5, WHITE, RED)
        else:
            pill(s, 8, y, 100, nh, lab, 11.5, (170,170,170), WHITE, bold=False, line=(235,225,222), lw=0.75)

def card(s, x, y, w, h, line=PINKL, fill=WHITE, lw=1.2):
    return shape(s, x, y, w, h, fill, line=line, lw=lw, st=5)

# ================= P1 封面 =================
s = new_slide(1, f"{BGS}/bg_cover.png")
txt(s, 160, 120, 640, 130, "思政说课", fs=92, color=REDT, font=FX, align=2)
# 分隔装饰
shape(s, 300, 296, 130, 2, GOLDD)
shape(s, 530, 296, 130, 2, GOLDD)
txt(s, 436, 278, 88, 40, "★", fs=22, color=GOLDD, align=2)
# 红色胶囊横幅
pill(s, 250, 352, 460, 50, "—— 深耕思政课堂  落实立德树人 ——", 19, GOLD, RED)
# 汇报人 / 时间
pill(s, 298, 428, 190, 38, "汇报人：懒洋洋", 14, INK, WHITE)
pill(s, 508, 428, 210, 38, "汇报时间：2026.08", 14, INK, WHITE)
print('P1 封面 完成')

# ================= P2 课程介绍 =================
s = new_slide(2, f"{BGS}/bg_content.png")
chrome(s, "课程介绍", active_tab=0)
# 主标题
txt(s, 380, 86, 60, 36, "★☆", fs=17, color=GOLDD, align=2)
txt(s, 170, 82, 620, 44, "素养提升三路径筑牢育人根基", fs=26, color=REDT, bold=True, align=2)
txt(s, 520, 86, 60, 36, "☆★", fs=17, color=GOLDD, align=2)
# 左侧教材
pic(s, f"{FIG}/tb.png", 118, 150, 188, 313)
# 卡1
card(s, 322, 140, 606, 158)
pill(s, 338, 150, 190, 28, "深化思政课堂教学", 14, WHITE, RED)
txt(s, 538, 150, 374, 28, "打造“理论+实践”融合式育人模式", fs=15, color=INK, bold=True, align=1)
txt(s, 340, 186, 580, 26, "•  课程严格依据新课标《思想政治》教材及课程标准", fs=13, color=INK, align=1, bold=False)
txt(s, 340, 214, 580, 26, "•  引导学生系统掌握思政理论的基本观点、价值导向", fs=13, color=INK, align=1, bold=False)
pill(s, 338, 246, 186, 30, "课程理论扎实学", 13, WHITE, RED)
pill(s, 534, 246, 186, 30, "思想政治理论素养", 13, WHITE, RED)
pill(s, 730, 246, 186, 30, "核心素养培育提升", 13, WHITE, RED)
# 卡2
card(s, 322, 310, 606, 68)
txt(s, 336, 326, 150, 36, "问题导向优化", fs=17, color=INK, bold=True, align=1)
pill(s, 496, 326, 146, 36, "聚焦教学难点", 14, INK, WHITE, line=PINKL, lw=1.2)
shape(s, 652, 332, 52, 24, RED, st=33)
txt(s, 714, 326, 196, 36, "调整教学方法", fs=14, color=INK, bold=True, align=1)
# 卡3
card(s, 322, 390, 606, 120)
pill(s, 338, 398, 574, 26, "网络课堂讲授  案例研讨  实践体验等多元教学法", 13, WHITE, RED)
for i, (img, cap) in enumerate([("c1_w.png", "理论夯实"), ("c2_w.png", "实践强化"), ("c3_w.png", "素养提升")]):
    px = 340 + i*192
    pic(s, f"{FIG}/{img}", px, 432, 182, 62)
    pill(s, px, 478, 74, 20, cap, 11, WHITE, RED)
print('P2 课程介绍 完成')

# ================= P3 学情分析 =================
s = new_slide(3, f"{BGS}/bg_content.png")
chrome(s, "学情分析", active_tab=0)
txt(s, 170, 82, 620, 44, "思政课堂学情画像与学习特点分析", fs=25, color=REDT, bold=True, align=2)
# 左卡：调研结论
card(s, 122, 138, 540, 158)
shape(s, 138, 152, 6, 22, RED)
txt(s, 152, 148, 500, 28, "基于课前调研数据，精准分析学情特征", fs=15, color=INK, bold=True, align=1)
txt(s, 152, 184, 500, 26, "•  学生对思政课学习兴趣较为积极，关注国家政策与社会热点", fs=12.5, color=INK, align=1, bold=False)
txt(s, 152, 212, 500, 26, "•  理论联系实际的能力有待提升，对理论的理解不足", fs=12.5, color=INK, align=1, bold=False)
txt(s, 152, 240, 500, 26, "•  通过问卷与访谈调研，精准掌握学生学习基础与需求", fs=12.5, color=INK, align=1, bold=False)
# 右侧调研照片
pic(s, f"{FIG}/c1_m.png", 670, 140, 130, 96)
pic(s, f"{FIG}/c2_m.png", 808, 140, 130, 96)
pill(s, 670, 244, 268, 24, "课前问卷调研与焦点访谈现场", 11.5, WHITE, RED)
# 两个图表卡
card(s, 122, 312, 380, 168)
txt(s, 122, 318, 380, 26, "学生思想政治基础水平分布", fs=14, color=REDT, bold=True, align=2)
pic(s, f"{FIG}/chart1.png", 130, 348, 364, 126)
card(s, 512, 312, 428, 168)
txt(s, 512, 318, 428, 26, "学生思政课堂素养变化趋势", fs=14, color=REDT, bold=True, align=2)
pic(s, f"{FIG}/chart2.png", 522, 348, 408, 126)
print('P3 学情分析 完成')

# ================= P4 课前准备 =================
s = new_slide(4, f"{BGS}/bg_content.png")
chrome(s, "课前准备", active_tab=1)
cols = [(118, "c1_sq.png", "学情分析与目标确立", ["课前问卷调研学生对思政课的兴趣与认知基础", "根据课标和教材梳理教学目标"],
         ["c2_s.png", "c3_s.png"], "精准定位教学起点与要求"),
        (398, "c2_sq.png", "教学方案与情境准备", ["收集时事热点案例、红色故事素材", "设计课堂讨论等教学活动方案"],
         ["c3_s.png", "c4_s.png"], "理论与实践结合的场景"),
        (678, "c3_sq.png", "预习任务与课堂衔接", ["引导学生自主学习核心概念", "布置小组探究问题，明确分工"],
         ["c4_s.png", "c2_s.png"], "课前预习与教学无缝衔接")]
for cx, circ, ttl, bl, imgs, foot in cols:
    card(s, cx, 158, 260, 348)
    # 圆形照片（压卡片顶部）
    try:
        csp = s.Shapes.AddShape(9, cx+82, 72, 96, 96)
        csp.Fill.UserPicture(os.path.normpath(f"{FIG}/{circ}"))
        csp.Line.ForeColor.RGB = h2b(WHITE); csp.Line.Weight = 2.5
    except Exception:
        pic(s, f"{FIG}/{circ}", cx+82, 72, 96, 96)
    pill(s, cx+12, 186, 236, 32, ttl, 15, WHITE, RED)
    txt(s, cx+16, 228, 228, 30, "✓  " + bl[0], fs=12, color=INK, align=1, bold=False)
    txt(s, cx+16, 262, 228, 30, "✓  " + bl[1], fs=12, color=INK, align=1, bold=False)
    pic(s, f"{FIG}/{imgs[0]}", cx+14, 300, 112, 76)
    pic(s, f"{FIG}/{imgs[1]}", cx+134, 300, 112, 76)
    txt(s, cx+12, 386, 236, 44, foot, fs=14, color=REDT, bold=True, align=2)
# 圆之间的衔接标签
for lx, l1, l2 in [(333, "学情调研", "目标梳理"), (613, "资源整合", "情境创设")]:
    txt(s, lx, 84, 110, 22, l1, fs=11.5, color=INK, align=2, bold=True)
    shape(s, lx+48, 110, 14, 14, RED, st=33)
    pill(s, lx+8, 128, 94, 22, l2, 11.5, WHITE, RED)
print('P4 课前准备 完成')

# ================= P5 能力提升 =================
s = new_slide(5, f"{BGS}/bg_content.png")
chrome(s, "能力提升", active_tab=2)
pic(s, f"{FIG}/hb.png", 806, 66, 142, 432)
rows = [(70, "c1_p5.png", "理论认知深化", "筑牢根基", ["进一步加深学生对思政理论的理解与内化", "为树立正确世界观、人生观奠定坚实基础"], None),
        (215, "c2_p5.png", "课堂互动优化", "激发热情", ["推动思政课教学从单向讲授转向双向互动"], ["课堂讨论参与", "小组合作探究"]),
        (360, "c4_p5.png", "实践能力强化", "知行合一", None, ["红色实践研学", "志愿服务参与"])]
for ry, img, ttl, tag, bl, pills in rows:
    card(s, 122, ry, 650, 130)
    pic(s, f"{FIG}/{img}", 134, ry+14, 180, 102)
    pill(s, 134, ry+96, 112, 20, ttl, 11, WHITE, RED)
    txt(s, 322, ry+16, 180, 30, ttl, fs=17, color=REDT, bold=True, align=1)
    txt(s, 502, ry+18, 190, 28, "»  " + tag, fs=15, color=INK, bold=True, align=1)
    if bl:
        for i, b in enumerate(bl):
            txt(s, 320, ry+52+i*26, 440, 24, "•  " + b, fs=12.5, color=INK, align=1, bold=False)
    if pills:
        for i, p in enumerate(pills):
            if ry < 300:
                pill(s, 320+i*164, ry+86, 152, 28, p, 12.5, WHITE, RED)
            else:
                shape(s, 320+i*170, ry+78, 24, 24, RED, st=9)
                pill(s, 352+i*170, ry+76, 150, 28, p, 12.5, REDT, WHITE, line=RED, lw=1.2)
print('P5 能力提升 完成')

# ================= P6 价值认同 =================
s = new_slide(6, f"{BGS}/bg_content.png")
chrome(s, "价值认同", active_tab=2)
txt(s, 100, 78, 760, 32, "思政教学实施是一个系统性、多层次、多维度的育人过程", fs=19, color=REDT, bold=True, align=2)
txt(s, 100, 112, 760, 32, "有效实现价值引领目标，引导学生树立正确三观", fs=19, color=REDT, bold=True, align=2)
# 左卡
card(s, 122, 158, 260, 344)
txt(s, 122, 164, 260, 30, "理论认知提升", fs=16, color=REDT, bold=True, align=2)
txt(s, 122, 196, 260, 26, "思政理论的理解与运用", fs=13.5, color=INK, bold=True, align=2)
pill(s, 134, 226, 110, 24, "理论深度不够", 10.5, REDT, WHITE, line=RED, lw=1.0)
pill(s, 254, 226, 110, 24, "课程学习经验", 10.5, REDT, WHITE, line=RED, lw=1.0)
txt(s, 132, 258, 240, 66, "围绕思政课核心理论，结合社会现实议题，立足热点案例进行分层解读，夯实理论根基", fs=11.5, color=INK, align=1, bold=False)
pic(s, f"{FIG}/c2_p6.png", 132, 334, 240, 115)
pill(s, 132, 434, 84, 22, "理论探究", 11.5, WHITE, RED)
# 中卡
card(s, 398, 158, 260, 344)
pill(s, 418, 164, 220, 30, "创新思维与能力提升", 14.5, WHITE, RED)
grid = [("c1_g.png", "理论探究"), ("c2_g.png", "资源拓展"), ("c3_g.png", "实践反思"), ("c4_g.png", "育人培育")]
for i, (img, cap) in enumerate(grid):
    gx = 410 + (i % 2)*122
    gy = 204 + (i // 2)*112
    pic(s, f"{FIG}/{img}", gx, gy, 114, 84)
    pill(s, gx+2, gy+86, 84, 20, cap, 10.5, WHITE, RED)
# 右卡
card(s, 674, 158, 266, 344)
txt(s, 674, 164, 266, 30, "价值认同提升", fs=16, color=REDT, bold=True, align=2)
txt(s, 674, 196, 266, 26, "思政素养的内化与践行", fs=13.5, color=INK, bold=True, align=2)
for i, p in enumerate(["理论认知", "情感认同", "行为践行"]):
    pill(s, 684+i*84, 226, 78, 24, p, 10.5, REDT, WHITE, line=RED, lw=1.0)
txt(s, 684, 258, 246, 66, "引导学生将价值理念内化于心、外化于行，把课堂所学转化为社会责任与行动自觉", fs=11.5, color=INK, align=1, bold=False)
pic(s, f"{FIG}/c3_p6.png", 684, 334, 246, 108)
pill(s, 684, 434, 84, 22, "知行合一", 11.5, WHITE, RED)
print('P6 价值认同 完成')

# ================= P7 教学反思 =================
s = new_slide(7, f"{BGS}/bg_content.png")
chrome(s, "教学反思", active_tab=3)
# 实践 / 学情 两行
pill(s, 122, 76, 66, 46, "实践", 18, WHITE, RED)
txt(s, 200, 76, 560, 46, "以“理论精讲·案例剖析·价值引领·实践拓展”为主线，\n推动思政育人高质量发展", fs=12.5, color=INK, align=1, bold=False)
pill(s, 122, 130, 66, 46, "学情", 18, WHITE, RED)
txt(s, 200, 130, 560, 46, "结合课前调研数据与课堂表现，精准把握学生理论基础、\n价值认知与实践需求", fs=12.5, color=INK, align=1, bold=False)
# 阶梯三卡
def step_card(x, y, w, h, num, ttl):
    card(s, x, y, w, h)
    txt(s, x+14, y+6, 64, 44, num, fs=27, color=GOLDD, bold=True, align=1)
    txt(s, x+80, y+10, w-90, 38, ttl, fs=19, color=REDT, bold=True, align=1)

step_card(122, 330, 245, 178, "01", "夯实基础")
txt(s, 122, 380, 245, 26, "课堂讲授 ＋ 小组研讨", fs=13.5, color=INK, bold=True, align=2)
txt(s, 122, 406, 245, 26, "理论学习效率提升 52%", fs=13, color=REDT, bold=True, align=2)
pic(s, f"{FIG}/c3_p7.png", 134, 434, 220, 64)

step_card(392, 268, 245, 240, "02", "价值塑造")
txt(s, 392, 320, 245, 26, "学生价值认同度提升", fs=13.5, color=INK, bold=True, align=2)
txt(s, 392, 346, 245, 26, "家国情怀三维培育路径", fs=13, color=REDT, bold=True, align=2)
pic(s, f"{FIG}/c1_p7.png", 404, 378, 220, 110)

step_card(662, 206, 265, 302, "03", "能力转化")
txt(s, 662, 258, 265, 26, "理论认知成为时代担当者", fs=13.5, color=INK, bold=True, align=2)
pill(s, 734, 288, 120, 26, "素养推进", 13, WHITE, RED)
for i, b in enumerate(["培育", "伦理", "科学", "公共"]):
    pill(s, 676+i*60, 324, 52, 40, b, 12, REDT, WHITE, line=RED, lw=1.0)
pic(s, f"{FIG}/c2_p7.png", 676, 378, 236, 110)
# 橙色箭头
for (ax, ay) in [(370, 330), (640, 268)]:
    ar = shape(s, ax, ay, 22, 26, ORG, st=33)
    try: ar.Rotation = -22
    except Exception: pass
print('P7 教学反思 完成')

# ================= P8 创新方向 =================
s = new_slide(8, f"{BGS}/bg_content.png")
chrome(s, "创新方向", active_tab=3)
tops = [(122, 262, "理论与实践脱节", "教学仍偏重理论讲授"),
        (398, 262, "互动形式较为单一", "未充分调动学生主动性"),
        (674, 256, "内容更新不及时", "与时代发展节奏有差距")]
for x, w, lab, sub in tops:
    card(s, x, 74, w, 68, fill=CREAM)
    txt(s, x+14, 80, w-60, 26, lab, fs=13.5, color=REDT, bold=True, align=1)
    txt(s, x+14, 108, w-60, 24, sub, fs=11.5, color=INK, align=1, bold=False)
    shape(s, x+w-38, 92, 26, 26, RED, st=9)
# 左卡：教学模式创新
card(s, 122, 165, 360, 340)
pill(s, 140, 172, 200, 30, "教学模式创新", 15, WHITE, RED)
pic(s, f"{FIG}/c1_p8.png", 140, 212, 158, 100)
pic(s, f"{FIG}/c2_p8.png", 306, 212, 158, 100)
txt(s, 122, 326, 360, 28, "实践育人成果", fs=14.5, color=REDT, bold=True, align=2)
pill(s, 140, 358, 324, 44, "开发思政主题案例教学资源\n助力学生价值认同度提升", 12.5, REDT, WHITE, line=RED, lw=1.2)
pill(s, 140, 412, 324, 44, "搭建线上思政学习互动平台\n形成线上线下混合式学习模式", 12.5, REDT, WHITE, line=RED, lw=1.2)
# 右卡：课程思政融合
card(s, 498, 165, 432, 340)
pill(s, 606, 172, 216, 30, "课程思政融合", 15, WHITE, RED)
# 中央金色圆环
try:
    dn = s.Shapes.AddShape(18, 634, 240, 150, 150)   # donut
    dn.Fill.ForeColor.RGB = h2b(GOLDD)
    dn.Line.Visible = False
    shape_text(dn, "育人\n架构", 19, REDT)
except Exception:
    pass
for (ex, ey, lab) in [(508, 258, "实践"), (508, 330, "探究"), (838, 258, "理论"), (838, 330, "引领")]:
    sp = s.Shapes.AddShape(9, ex, ey, 76, 40)
    sp.Fill.Solid(); sp.Fill.ForeColor.RGB = h2b(WHITE)
    sp.Line.ForeColor.RGB = h2b(INK); sp.Line.Weight = 1.2
    shape_text(sp, lab, 14, INK)
txt(s, 588, 300, 42, 36, "＋", fs=24, color=REDT, bold=True, align=2)
txt(s, 790, 300, 42, 36, "＋", fs=24, color=REDT, bold=True, align=2)
pill(s, 514, 438, 190, 40, "个性化学习方案", 14, REDT, WHITE, line=RED, lw=1.4)
pill(s, 724, 438, 190, 40, "线上互动课堂", 14, REDT, WHITE, line=RED, lw=1.4)
print('P8 创新方向 完成')

# ================= 导出 =================
if os.path.exists(OUT_PDF):
    try: os.remove(OUT_PDF)
    except Exception: pass
pres.SaveAs(os.path.normpath(OUT_PPTX), 26)   # pptx
pres.SaveAs(os.path.normpath(OUT_PDF), 32)    # pdf
print("pptx size:", os.path.getsize(OUT_PPTX) if os.path.exists(OUT_PPTX) else "MISSING")
print("pdf  size:", os.path.getsize(OUT_PDF) if os.path.exists(OUT_PDF) else "MISSING")
print("ALL DONE — WPS 窗口保留")
