#!/bin/bash
# 侧载对照测试：清空 Wef → 只放入指定 manifest → 清缓存 → 重启 PowerPoint
# 用法: bash scripts/sideload-test.sh <manifest文件路径>
M="$1"
[ -f "$M" ] || { echo "❌ manifest 不存在: $M"; exit 1; }
xmllint --noout "$M" || { echo "❌ manifest XML 不合法"; exit 1; }

echo "→ 请确保演示文稿已保存，正在退出 PowerPoint..."
osascript -e 'quit app "Microsoft PowerPoint"' 2>/dev/null; sleep 3
pkill -x "Microsoft PowerPoint" 2>/dev/null && sleep 2   # 兜底强杀（已先尝试正常退出）

WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/Wef"
mkdir -p "$WEF"
rm -f "$WEF"/*.xml 2>/dev/null          # 清空旧 manifest，保证测试唯一性
cp "$M" "$WEF/QSlide.manifest.xml" || { echo "❌ 拷贝失败（TCC 权限）"; exit 1; }

rm -rf "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/OfficeFileCache" \
       "$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/OfficeFileCache_15" 2>/dev/null

echo "--- Wef 当前内容 ---"; ls -la "$WEF"
open -a "Microsoft PowerPoint"
echo "✅ 已重启 PowerPoint。检查「插入 → 加载项」下拉里是否出现新加载项（约 10 秒后）。"
