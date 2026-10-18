#!/usr/bin/env python3
"""把 public/assets/icons/ 下按按钮名生成的图标逐个接线到 manifest.xml：
- 每个 Control 的 <Icon> 替换为专属 resid（Icon.<Name>.16x16/32x32/80x80）
- <bt:Images> 追加对应资源条目（URL 指向 dev server）
- Group 头像保留共享 Icon.16x16 等资源
运行：python3 scripts/apply_icon_manifest.py
"""
import re
import os

ROOT = os.path.join(os.path.dirname(__file__), "..")
MANIFEST = os.path.join(ROOT, "manifest.xml")

# Control id 后缀（manifest 中 id="QSlide.<Name>"，无 Btn 前缀）-> 图标名
MAP = {
    "BtnTaskpane": "pane", "SameWidth": "SameWidth", "SameHeight": "SameHeight",
    "TouchLeft": "TouchLeft", "TouchUp": "TouchUp", "RemoveText": "RemoveText",
    "DistributeH": "DistributeH", "ActionTitle": "ActionTitle", "Unify": "Unify",
    "EdgeUp": "EdgeUp", "EdgeDown": "EdgeDown", "EdgeLeft": "EdgeLeft",
    "EdgeRight": "EdgeRight", "SubjectTitle": "SubjectTitle", "Footnote": "Footnote",
    "MergeText": "MergeText", "CollapseSpaces": "CollapseSpaces",
    "RoundUp": "RoundUp", "RoundDown": "RoundDown", "Straighten": "Straighten",
    "RemoveOutline": "RemoveOutline", "UnifyFont": "UnifyFont",
    "InsertTextbox": "InsertTextbox", "CopyToAll": "CopyToAll",
    "RemoveNotes": "RemoveNotes", "RemoveAnimations": "RemoveAnimations",
    "DuplicateSlide": "DuplicateSlide", "Harvey": "Harvey", "Agenda": "Agenda",
    "ValueChain": "ValueChain", "FontUp": "FontUp", "FontDown": "FontDown",
    "LineUp": "LineUp", "LineDown": "LineDown", "RemoveEffects": "RemoveEffects",
    "PickFormat": "PickFormat", "ApplyFormat": "ApplyFormat",
    "ZeroMargins": "ZeroMargins", "ZeroMarginsAll": "ZeroMarginsAll",
    "Table": "Table", "Grayscale": "Grayscale",
}

s = open(MANIFEST).read()

# 1) 逐 Control 替换 Icon 块
replaced = []
def sub_icon(m):
    ctl, mid = m.group(1), m.group(2)
    name = MAP.get(ctl)
    if not name:
        return m.group(0)
    replaced.append(name)
    return (f'<Control xsi:type="Button" id="QSlide.{ctl}">{mid}'
            f'<Icon><bt:Image size="16" resid="Icon.{name}.16x16"/>'
            f'<bt:Image size="32" resid="Icon.{name}.32x32"/>'
            f'<bt:Image size="80" resid="Icon.{name}.80x80"/></Icon>')

s2 = re.sub(
    r'<Control xsi:type="Button" id="QSlide\.(\w+)">(.*?)<Icon>.*?</Icon>',
    sub_icon, s, flags=re.S)
assert len(replaced) == len(MAP), f"只替换了 {len(replaced)}/{len(MAP)} 个按钮图标"

# 2) 追加 bt:Image 资源条目
entries = []
for ctl, name in MAP.items():
    for size in (16, 32, 80):
        entries.append(
            f'        <bt:Image id="Icon.{name}.{size}x{size}" '
            f'DefaultValue="https://localhost:3000/assets/icons/{name}-{size}.png"/>')
s2 = s2.replace("      </bt:Images>", "\n".join(entries) + "\n      </bt:Images>")

open(MANIFEST, "w").write(s2)
print(f"manifest 接线完成：{len(replaced)} 个按钮图标 + {len(entries)} 条资源")
