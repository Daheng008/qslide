#!/usr/bin/env python3
"""QSlide 按钮图标生成器。
在 160x160 画布上绘制语义化图标（PPT 原生同概念的按钮用 PPT 同语义图形），
缩放导出 80/32/16 三档 PNG 至 public/assets/icons/，另生成应用图标 public/assets/icon-*.png。
运行：python3 scripts/gen_icons.py
"""
from PIL import Image, ImageDraw, ImageFont
import os

OUT = os.path.join(os.path.dirname(__file__), "..", "src", "public", "assets", "icons")
APP = os.path.join(os.path.dirname(__file__), "..", "src", "public", "assets")
os.makedirs(OUT, exist_ok=True)

BLUE = (46, 117, 182, 255)      # 品牌蓝（描边主色）
NAVY = (31, 78, 121, 255)       # 深蓝（箭头/强调）
LIGHT = (222, 235, 247, 255)    # 浅蓝填充
DARK = (89, 89, 89, 255)        # 中性深灰
RED = (192, 0, 0, 255)          # 移除类标记
ORANGE = (237, 125, 49, 255)
GRAY = (165, 165, 165, 255)
MIDGRAY = (110, 110, 110, 255)

S = 160  # 画布尺寸
W = 11   # 主线宽


def canvas():
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    return img, ImageDraw.Draw(img)


def font(sz):
    return ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", sz)


def arrow(d, p1, p2, color=NAVY, w=W, head=None):
    """单头箭头 p1 -> p2。"""
    head = head or w * 2.6
    d.line([p1, p2], fill=color, width=w)
    _head(d, p2, p1, color, w, head)


def darrow(d, p1, p2, color=NAVY, w=W, head=None):
    """双头箭头。"""
    head = head or w * 2.6
    d.line([p1, p2], fill=color, width=w)
    _head(d, p2, p1, color, w, head)
    _head(d, p1, p2, color, w, head)


def _head(d, tip, tail, color, w, head):
    import math
    ang = math.atan2(tail[1] - tip[1], tail[0] - tip[0])
    a1 = (tip[0] + head * math.cos(ang - 0.44), tip[1] + head * math.sin(ang - 0.44))
    a2 = (tip[0] + head * math.cos(ang + 0.44), tip[1] + head * math.sin(ang + 0.44))
    d.polygon([tip, a1, a2], fill=color)


def rect(d, box, fill=LIGHT, outline=BLUE, w=W, r=0):
    if r:
        d.rounded_rectangle(box, radius=r, fill=fill, outline=outline, width=w)
    else:
        d.rectangle(box, fill=fill, outline=outline, width=w)
def xmark(d, box, color=RED, w=None):
    """红叉：box 区域内的对角交叉。"""
    w = w or W
    x0, y0, x1, y1 = box
    d.line([(x0, y0), (x1, y1)], fill=color, width=w)
    d.line([(x0, y1), (x1, y0)], fill=color, width=w)


def letter(d, ch, xy, sz=96, color=NAVY):
    d.text(xy, ch, font=font(sz), fill=color, anchor="mm")


# ---------------- 各图标绘制函数（160 画布） ----------------

def i_pane(d):
    d.rectangle((98, 32, 138, 128), fill=LIGHT)
    rect(d, (22, 32, 138, 128), fill=None, r=10)
    for y in (58, 82, 106):
        d.line([(42, y), (80, y)], fill=BLUE, width=8)


def i_SameWidth(d):
    rect(d, (26, 26, 62, 118))            # 窄高
    rect(d, (86, 56, 134, 118))           # 宽矮
    darrow(d, (30, 138), (130, 138))      # ↔ 下方


def i_SameHeight(d):
    rect(d, (26, 40, 74, 122))            # 窄
    rect(d, (96, 66, 136, 122))           # 矮
    darrow(d, (80, 24), (80, 140))        # ↕


def i_TouchLeft(d):
    arrow(d, (96, 26), (128, 26), w=9)    # →
    arrow(d, (64, 26), (32, 26), w=9)     # ←
    rect(d, (22, 58, 60, 122))
    rect(d, (71, 58, 105, 122))
    rect(d, (116, 58, 140, 122))


def i_TouchUp(d):
    arrow(d, (26, 88), (26, 120), w=9)    # ↓
    arrow(d, (134, 88), (134, 120), w=9)  # ↓
    rect(d, (52, 40, 108, 66))
    rect(d, (52, 76, 108, 110))
    rect(d, (52, 120, 108, 142))


def i_RemoveText(d):
    letter(d, "T", (70, 74), 110)
    xmark(d, (100, 96, 142, 138), w=12)


def i_DistributeH(d):
    rect(d, (18, 46, 46, 118))
    rect(d, (68, 46, 92, 118), fill=(255, 255, 255, 255))
    rect(d, (114, 46, 142, 118))
    darrow(d, (48, 82), (66, 82), w=6, head=13)
    darrow(d, (94, 82), (112, 82), w=6, head=13)


def i_DistributeV(d):
    rect(d, (46, 18, 118, 46))
    rect(d, (46, 68, 118, 92), fill=(255, 255, 255, 255))
    rect(d, (46, 114, 118, 142))
    darrow(d, (82, 48), (82, 66), w=6, head=13)
    darrow(d, (82, 94), (82, 112), w=6, head=13)


def i_ActionTitle(d):
    rect(d, (22, 30, 138, 62), fill=NAVY, outline=NAVY)
    for y in (88, 112):
        d.line([(22, y), (138, y)], fill=BLUE, width=8)
    d.line([(22, 136), (96, 136)], fill=BLUE, width=8)


def i_Unify(d):
    d.ellipse((28, 28, 132, 132), fill=LIGHT, outline=BLUE, width=W)
    d.ellipse((78, 44, 96, 62), fill=NAVY)
    d.ellipse((46, 76, 64, 94), fill=ORANGE)
    d.ellipse((80, 96, 98, 114), fill=GRAY)
    d.ellipse((104, 72, 120, 88), fill=BLUE)


def i_EdgeUp(d):
    d.rectangle((14, 20, 146, 34), fill=NAVY)
    rect(d, (48, 66, 112, 130))
    arrow(d, (80, 122), (80, 48), w=9)


def i_EdgeDown(d):
    d.rectangle((14, 126, 146, 140), fill=NAVY)
    rect(d, (48, 30, 112, 94))
    arrow(d, (80, 38), (80, 112), w=9)


def i_EdgeLeft(d):
    d.rectangle((20, 14, 34, 146), fill=NAVY)
    rect(d, (66, 48, 130, 112))
    arrow(d, (122, 80), (48, 80), w=9)


def i_EdgeRight(d):
    d.rectangle((126, 14, 140, 146), fill=NAVY)
    rect(d, (30, 48, 94, 112))
    arrow(d, (38, 80), (112, 80), w=9)


def i_SubjectTitle(d):
    letter(d, "T", (80, 72), 118, MIDGRAY)
    d.line([(44, 134), (116, 134)], fill=MIDGRAY, width=7)


def i_Footnote(d):
    rect(d, (34, 22, 126, 138), fill=(255, 255, 255, 255))
    for y in (48, 68):
        d.line([(50, y), (110, y)], fill=GRAY, width=7)
    d.line([(50, 104), (94, 104)], fill=BLUE, width=7)
    letter(d, "*", (114, 112), 56, ORANGE)


def i_MergeText(d):
    rect(d, (16, 30, 62, 74))
    rect(d, (34, 66, 78, 110), fill=(255, 255, 255, 255))
    arrow(d, (88, 80), (104, 80), w=9)
    rect(d, (110, 46, 146, 114))
    letter(d, "T", (128, 80), 44)


def i_CollapseSpaces(d):
    arrow(d, (44, 32), (70, 32), w=9)     # →
    arrow(d, (116, 32), (90, 32), w=9)    # ←
    letter(d, "a", (52, 92), 74)
    letter(d, "b", (104, 96), 62, BLUE)
    d.line([(78, 96), (82, 96)], fill=RED, width=8)


def i_RoundUp(d):
    rect(d, (26, 40, 134, 132), r=14)
    darc(d, (98, 20, 142, 64), 180, 270, NAVY)
    arrow(d, (120, 42), (142, 20), w=9)


def darc(d, box, a0, a1, color):
    d.arc(box, a0, a1, fill=color, width=9)


def i_RoundDown(d):
    rect(d, (26, 40, 134, 132), r=44)
    arrow(d, (140, 22), (114, 48), w=9)


def i_Straighten(d):
    rect(d, (16, 44, 66, 120), r=18, fill=None, outline=GRAY, w=8)
    arrow(d, (74, 82), (90, 82), w=9)
    rect(d, (96, 44, 146, 120), fill=None)


def i_RemoveOutline(d):
    for seg in dashed_rect((30, 34, 130, 126)):
        d.line(seg, fill=BLUE, width=8)
    xmark(d, (104, 96, 146, 138), w=11)


def dashed_rect(box, dash=14, gap=10):
    x0, y0, x1, y1 = box
    segs = []
    def hline(a, b, y):
        x = a
        while x < b:
            segs.append(((x, y), (min(x + dash, b), y)))
            x += dash + gap
    def vline(a, b, x):
        y = a
        while y < b:
            segs.append(((x, y), (x, min(y + dash, b))))
            y += dash + gap
    hline(x0, x1, y0); hline(x0, x1, y1); vline(y0, y1, x0); vline(y0, y1, x1)
    return segs


def i_UnifyFont(d):
    letter(d, "A", (52, 78), 104)
    letter(d, "a", (112, 106), 62, BLUE)
    arrow(d, (96, 52), (120, 32), w=8)


def i_InsertTextbox(d):
    rect(d, (24, 40, 122, 120), fill=(255, 255, 255, 255))
    letter(d, "T", (66, 80), 56)
    d.line([(136, 48), (136, 112)], fill=ORANGE, width=8)
    d.line([(128, 48), (144, 48)], fill=ORANGE, width=7)
    d.line([(128, 112), (144, 112)], fill=ORANGE, width=7)


def i_CopyToAll(d):
    rect(d, (20, 22, 78, 66))
    for i, (ox, oy) in enumerate(((72, 62), (60, 88), (48, 114))):
        rect(d, (ox, oy, ox + 58, oy + 40), fill=(255, 255, 255, 255) if i else LIGHT, w=8)
    arrow(d, (88, 44), (110, 60), w=8)


def i_RemoveNotes(d):
    rect(d, (30, 22, 118, 138), fill=(255, 255, 255, 255))
    for y in (52, 76, 100):
        d.line([(48, y), (100, y)], fill=GRAY, width=7)
    xmark(d, (102, 92, 144, 134), w=11)


def i_RemoveAnimations(d):
    star(d, (18, 24, 118, 124))
    xmark(d, (104, 98, 146, 140), w=11)


def star(d, box, fill=LIGHT, outline=BLUE):
    import math
    cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
    R = (box[2] - box[0]) / 2
    pts = []
    for k in range(10):
        r = R if k % 2 == 0 else R * 0.45
        a = -math.pi / 2 + k * math.pi / 5
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    d.polygon(pts, fill=fill, outline=outline)


def i_DuplicateSlide(d):
    rect(d, (46, 24, 140, 92), fill=(255, 255, 255, 255), w=9)
    rect(d, (20, 52, 114, 136))
    d.line([(36, 72), (98, 72)], fill=NAVY, width=7)


def i_Harvey(d):
    d.pieslice((28, 28, 132, 132), 90, 270, fill=BLUE, outline=BLUE, width=4)
    d.ellipse((28, 28, 132, 132), fill=None, outline=NAVY, width=W)


def i_Agenda(d):
    for k, y in enumerate((42, 80, 118)):
        d.rectangle((24, y - 14, 52, y + 14), fill=NAVY if k == 0 else BLUE)
        d.line([(66, y), (136, y)], fill=BLUE, width=8)


def i_ValueChain(d):
    def chevron(x, y0, y1, tip, color):
        d.polygon([(x, y0), (x + 34, y0), (x + 34 + tip, (y0 + y1) / 2),
                   (x + 34, y1), (x, y1), (x + tip, (y0 + y1) / 2)], fill=color)
    chevron(20, 34, 78, 14, BLUE)
    chevron(20, 88, 132, 14, NAVY)


def i_FontUp(d):
    letter(d, "A", (58, 84), 116)
    d.polygon([(112, 88), (142, 88), (127, 56)], fill=NAVY)


def i_FontDown(d):
    letter(d, "A", (58, 84), 116)
    d.polygon([(112, 60), (142, 60), (127, 92)], fill=NAVY)


def i_LineUp(d):
    for y in (44, 112, 140):
        d.line([(24, y), (104, y)], fill=BLUE, width=8)
    arrow(d, (134, 132), (134, 56), w=9)


def i_LineDown(d):
    for y in (26, 54, 122):
        d.line([(24, y), (104, y)], fill=BLUE, width=8)
    arrow(d, (134, 34), (134, 110), w=9)


def i_RemoveEffects(d):
    d.rectangle((44, 54, 138, 132), fill=GRAY)          # 阴影（偏移灰）
    rect(d, (26, 36, 120, 114))
    xmark(d, (100, 20, 142, 62), w=11)


def brush(d, fill, outline):
    """格式刷（PPT 格式刷同语义）：斜杆刷柄 + 梯形刷头 + 涂抹线。"""
    d.line([(108, 20), (76, 58)], fill=outline, width=13)         # 柄
    d.polygon([(52, 62), (94, 62), (84, 96), (62, 96)], fill=fill, outline=outline)
    for y in (112, 132):
        d.line([(30, y), (86, y)], fill=BLUE, width=8)


def i_PickFormat(d):
    brush(d, (255, 255, 255, 255), NAVY)


def i_ApplyFormat(d):
    brush(d, ORANGE, NAVY)


def i_ZeroMargins(d):
    rect(d, (22, 30, 138, 130), fill=(255, 255, 255, 255), w=9)
    for seg in dashed_rect((50, 56, 110, 104), dash=10, gap=7):
        d.line(seg, fill=RED, width=7)
    arrow(d, (62, 80), (34, 80), w=7, head=16)
    arrow(d, (98, 80), (126, 80), w=7, head=16)


def i_ZeroMarginsAll(d):
    rect(d, (56, 18, 144, 84), fill=(255, 255, 255, 255), w=8)
    rect(d, (16, 48, 112, 142), fill=LIGHT, w=9)
    for seg in dashed_rect((26, 60, 102, 132), dash=9, gap=6):
        d.line(seg, fill=RED, width=6)


def i_Table(d):
    d.rectangle((18, 34, 142, 128), fill=(255, 255, 255, 255), outline=BLUE, width=9)
    d.rectangle((18, 34, 142, 64), fill=BLUE)
    for x in (59, 100):
        d.line([(x, 64), (x, 128)], fill=BLUE, width=7)
    d.line([(18, 96), (142, 96)], fill=BLUE, width=7)


def i_Grayscale(d):
    d.pieslice((28, 28, 132, 132), 90, 270, fill=MIDGRAY)
    d.ellipse((28, 28, 132, 132), fill=None, outline=DARK, width=W)
    d.line([(80, 28), (80, 132)], fill=DARK, width=7)




def i_CenterH(d):
    d.rectangle((16, 44, 144, 116), outline=MIDGRAY, width=5)
    d.line([(80, 44), (80, 116)], fill=NAVY, width=3)
    rect(d, (54, 64, 106, 96), fill=LIGHT, outline=BLUE, w=9)


def i_CenterV(d):
    d.rectangle((16, 44, 144, 116), outline=MIDGRAY, width=5)
    d.line([(16, 80), (144, 80)], fill=NAVY, width=3)
    rect(d, (54, 64, 106, 96), fill=LIGHT, outline=BLUE, w=9)


def i_CenterBoth(d):
    d.rectangle((16, 44, 144, 116), outline=MIDGRAY, width=5)
    d.line([(80, 44), (80, 116)], fill=NAVY, width=3)
    d.line([(16, 80), (144, 80)], fill=NAVY, width=3)
    rect(d, (54, 64, 106, 96), fill=LIGHT, outline=BLUE, w=9)


def i_RotationZero(d):
    rect(d, (40, 60, 120, 100), fill=LIGHT, outline=BLUE, w=9)
    d.arc((56, 28, 104, 76), 205, 335, fill=ORANGE, width=7)
    d.polygon([(80, 28), (72, 42), (88, 42)], fill=ORANGE)


def i_ZFront(d):
    rect(d, (44, 52, 118, 116), fill=(210, 210, 210, 255), outline=MIDGRAY, w=7)
    rect(d, (30, 40, 104, 104), fill=LIGHT, outline=BLUE, w=9)
    arrow(d, (60, 18), (60, 38), w=6, head=13)


def i_ZBack(d):
    rect(d, (30, 40, 104, 104), fill=LIGHT, outline=BLUE, w=9)
    rect(d, (44, 52, 118, 116), fill=(210, 210, 210, 255), outline=MIDGRAY, w=7)
    arrow(d, (82, 138), (82, 120), w=6, head=13)


def i_ZForward(d):
    rect(d, (46, 52, 116, 114), fill=(210, 210, 210, 255), outline=MIDGRAY, w=7)
    rect(d, (34, 42, 104, 104), fill=LIGHT, outline=BLUE, w=9)
    arrow(d, (56, 22), (56, 34), w=5, head=11)


def i_ZBackward(d):
    rect(d, (34, 42, 104, 104), fill=LIGHT, outline=BLUE, w=9)
    rect(d, (46, 52, 116, 114), fill=(210, 210, 210, 255), outline=MIDGRAY, w=7)
    arrow(d, (84, 134), (84, 122), w=5, head=11)


# ---------------- 插入组：形状 / 图片 ----------------

def poly_shape(d, box, n, rot=-90, fill=LIGHT, outline=BLUE):
    """正 n 边形。rot=-90 表示一个顶点朝上。"""
    import math
    x0, y0, x1, y1 = box
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    R = min(x1 - x0, y1 - y0) / 2
    pts = [(cx + R * math.cos(math.radians(rot + k * 360 / n)),
            cy + R * math.sin(math.radians(rot + k * 360 / n))) for k in range(n)]
    d.polygon(pts, fill=fill, outline=outline, width=7)


def i_ShapeMenu(d):
    """PPT「形状」下拉同语义：矩形 + 椭圆 + 五角星组合。"""
    rect(d, (20, 26, 74, 72))
    d.ellipse((88, 26, 142, 72), fill=None, outline=BLUE, width=8)
    star(d, (26, 92, 82, 148), fill=LIGHT, outline=BLUE)
    d.line([(98, 120), (134, 120)], fill=NAVY, width=9)


def i_ShapeRect(d):
    rect(d, (30, 44, 130, 116))


def i_ShapeRoundRect(d):
    rect(d, (30, 44, 130, 116), r=24)


def i_ShapeEllipse(d):
    d.ellipse((26, 40, 134, 120), fill=LIGHT, outline=BLUE, width=W)


def i_ShapeTriangle(d):
    d.polygon([(80, 30), (134, 126), (26, 126)], fill=LIGHT, outline=BLUE, width=7)


def i_ShapeDiamond(d):
    d.polygon([(80, 28), (136, 80), (80, 132), (24, 80)], fill=LIGHT, outline=BLUE, width=7)


def i_ShapePentagon(d):
    poly_shape(d, (26, 30, 134, 130), 5)


def i_ShapeHexagon(d):
    poly_shape(d, (26, 36, 134, 124), 6, rot=0)


def i_ShapeArrowRight(d):
    """PPT 右箭头形状（块状箭头）。"""
    d.polygon([(20, 58), (88, 58), (88, 30), (142, 80), (88, 130), (88, 102), (20, 102)],
              fill=LIGHT, outline=BLUE, width=7)


def i_ShapeStar5(d):
    star(d, (22, 24, 138, 140))


def i_ShapeHeart(d):
    d.ellipse((24, 30, 82, 88), fill=LIGHT, outline=BLUE, width=7)
    d.ellipse((78, 30, 136, 88), fill=LIGHT, outline=BLUE, width=7)
    d.polygon([(30, 72), (130, 72), (80, 140)], fill=LIGHT, outline=BLUE, width=7)
    d.line([(30, 74), (130, 74)], fill=BLUE, width=4)


def i_ShapeCloud(d):
    for box in ((22, 80, 68, 122), (48, 44, 110, 98), (94, 68, 140, 120)):
        d.ellipse(box, fill=LIGHT, outline=BLUE, width=6)
    d.rectangle((32, 96, 130, 118), fill=LIGHT)


def i_ShapeLine(d):
    d.line([(26, 130), (134, 30)], fill=NAVY, width=W)


def i_Picture(d):
    """PPT「图片」同语义：相框 + 太阳 + 山。"""
    rect(d, (22, 30, 138, 130), fill=(255, 255, 255, 255), w=9)
    d.ellipse((38, 44, 64, 70), fill=ORANGE)
    d.polygon([(32, 118), (74, 72), (106, 118)], fill=BLUE)
    d.polygon([(84, 118), (112, 86), (136, 118)], fill=NAVY)

ICONS = {
    "pane": i_pane, "SameWidth": i_SameWidth, "SameHeight": i_SameHeight,
    "TouchLeft": i_TouchLeft, "TouchUp": i_TouchUp, "RemoveText": i_RemoveText,
    "DistributeH": i_DistributeH, "DistributeV": i_DistributeV, "ActionTitle": i_ActionTitle, "Unify": i_Unify,
    "EdgeUp": i_EdgeUp, "EdgeDown": i_EdgeDown, "EdgeLeft": i_EdgeLeft, "EdgeRight": i_EdgeRight,
    "SubjectTitle": i_SubjectTitle, "Footnote": i_Footnote, "MergeText": i_MergeText,
    "CollapseSpaces": i_CollapseSpaces, "RoundUp": i_RoundUp, "RoundDown": i_RoundDown,
    "Straighten": i_Straighten, "RemoveOutline": i_RemoveOutline, "UnifyFont": i_UnifyFont,
    "InsertTextbox": i_InsertTextbox, "CopyToAll": i_CopyToAll, "RemoveNotes": i_RemoveNotes,
    "RemoveAnimations": i_RemoveAnimations, "DuplicateSlide": i_DuplicateSlide,
    "Harvey": i_Harvey, "Agenda": i_Agenda, "ValueChain": i_ValueChain,
    "FontUp": i_FontUp, "FontDown": i_FontDown, "LineUp": i_LineUp, "LineDown": i_LineDown,
    "RemoveEffects": i_RemoveEffects, "PickFormat": i_PickFormat, "ApplyFormat": i_ApplyFormat,
    "ZeroMargins": i_ZeroMargins, "ZeroMarginsAll": i_ZeroMarginsAll,
    "Table": i_Table, "Grayscale": i_Grayscale,
    "CenterH": i_CenterH, "CenterV": i_CenterV, "CenterBoth": i_CenterBoth,
    "RotationZero": i_RotationZero, "ZFront": i_ZFront, "ZBack": i_ZBack,
    "ZForward": i_ZForward, "ZBackward": i_ZBackward,
    "ShapeMenu": i_ShapeMenu,
    "ShapeRect": i_ShapeRect, "ShapeRoundRect": i_ShapeRoundRect, "ShapeEllipse": i_ShapeEllipse,
    "ShapeTriangle": i_ShapeTriangle, "ShapeDiamond": i_ShapeDiamond, "ShapePentagon": i_ShapePentagon,
    "ShapeHexagon": i_ShapeHexagon, "ShapeArrowRight": i_ShapeArrowRight, "ShapeStar5": i_ShapeStar5,
    "ShapeHeart": i_ShapeHeart, "ShapeCloud": i_ShapeCloud, "ShapeLine": i_ShapeLine,
    "Picture": i_Picture,
}


def export(img, base):
    for size in (80, 32, 16):
        img.resize((size, size), Image.LANCZOS).save(os.path.join(OUT, f"{base}-{size}.png"))


def main():
    for name, fn in ICONS.items():
        img, d = canvas()
        fn(d)
        export(img, name)
        print("icon:", name)
    # 应用图标：蓝底圆角方块 + 白色 Q
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((8, 8, 152, 152), radius=32, fill=BLUE)
    d.text((80, 74), "Q", font=font(96), fill=(255, 255, 255, 255), anchor="mm")
    d.ellipse((96, 96, 126, 126), fill=(255, 255, 255, 255))
    for size in (80, 32, 16):
        img.resize((size, size), Image.LANCZOS).save(os.path.join(APP, f"icon-{size}.png"))
    print("app icon done;", len(ICONS), "button icons")


if __name__ == "__main__":
    main()
