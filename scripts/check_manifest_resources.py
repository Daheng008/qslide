#!/usr/bin/env python3
"""QSlide manifest 资源自检：Tip 节位 / resid 悬空 / 图标文件存在 / 重复定义 / Menu 控件。
用法: python3 scripts/check_manifest_resources.py  （在仓库根目录执行）
"""
import re, os, sys

root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
s = open(os.path.join(root, "manifest.xml"), encoding="utf-8").read()
errors = []

short = re.search(r"<bt:ShortStrings>.*?</bt:ShortStrings>", s, re.S).group(0)
long_ = re.search(r"<bt:LongStrings>.*?</bt:LongStrings>", s, re.S).group(0)
long_ids = set(re.findall(r'id="([^"]+)"', long_))

# 坑① Supertip Description 必须引用 LongStrings
descs = set(re.findall(r'<Description resid="([^"]+)"/>', s))
bad_tips = descs - long_ids
if bad_tips:
    errors.append(f"Tip 错节（必须在 LongStrings）: {sorted(bad_tips)}")
tips_in_short = re.findall(r'<bt:String id="[^"]*\.Tip"', short)
if tips_in_short:
    errors.append(f"ShortStrings 残留 Tip {len(tips_in_short)} 条（应全部在 LongStrings）")

# 坑② resid 引用必须闭环
defs = set(re.findall(r'<bt:(?:String|Url|Image)\b[^>]*? id="([^"]+)"', s))
refs = set(re.findall(r'resid="([^"]+)"', s))
dangling = refs - defs
if dangling:
    errors.append(f"resid 悬空（引用但未定义）: {sorted(dangling)}")

# 坑③ 重复资源定义
all_defs = re.findall(r'<bt:(?:String|Url|Image)\b[^>]*? id="([^"]+)"', s)
dups = sorted(i for i in set(all_defs) if all_defs.count(i) > 1)
if dups:
    errors.append(f"重复资源定义: {dups}")

# 图标文件存在性（相对 src/public）
bad_icons = []
for i in (refs & defs):
    m = re.match(r"Icon\.(\w+)\.(\d+)x\2$", i)
    if m and not os.path.exists(os.path.join(root, "src/public/assets/icons", f"{m.group(1)}-{m.group(2)}.png")):
        bad_icons.append(i)
if bad_icons:
    errors.append(f"图标 PNG 缺失: {sorted(bad_icons)}")

# Menu 控件（PowerPoint 不支持）
if 'xsi:type="Menu"' in s:
    errors.append('存在 xsi:type="Menu" 控件：Mac PowerPoint 不支持，会整体拒绝 manifest')

if errors:
    print("❌ manifest 自检失败：")
    for e in errors:
        print("  -", e)
    sys.exit(1)
print("✅ manifest 资源自检通过（Tip 节位 / resid 闭环 / 无重复 / 图标齐全 / 无 Menu）")
