#!/bin/bash
# QSlide macOS sideload 一键安装
# 1) 拷 manifest 到 PowerPoint 沙箱 Wef 目录
# 2) 提取 vite dev server 自签证书并装入钥匙串信任（Office/WKWebView 要求）
# 前置：本终端需有「完全磁盘访问」权限；dev server 已在跑（npm run dev）
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"

# --- 0. 确保 dev server 在跑（从本终端启动的进程常驻）---
if ! curl -sk -o /dev/null https://localhost:3000/taskpane.html; then
  echo "▶️ dev server 未运行，启动中..."
  nohup bash "$DIR/scripts/dev-server.sh" > /tmp/qslide-dev.log 2>&1 &
  for i in $(seq 1 20); do
    sleep 1
    if curl -sk -o /dev/null https://localhost:3000/taskpane.html; then break; fi
  done
fi
if curl -sk -o /dev/null https://localhost:3000/taskpane.html; then
  echo "✅ dev server 在线: https://localhost:3000"
else
  echo "❌ dev server 启动失败，查看 /tmp/qslide-dev.log"
fi

# --- 1. manifest sideload ---
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/Wef"
mkdir -p "$WEF"
cp "$DIR/manifest.xml" "$WEF/QSlide.manifest.xml"
echo "✅ manifest 已拷贝到: $WEF/QSlide.manifest.xml"
# 清理 Wef 里的旧 manifest 文件（避免同 id 冲突导致 Office 加载旧版本）
for stale in "$WEF/manifest.xml" "$WEF"/QSlide_*.manifest.xml "$WEF"/*.manifest.xml; do
  [ -f "$stale" ] || continue
  base="$(basename "$stale")"
  [ "$base" = "QSlide.manifest.xml" ] && continue
  rm -f "$stale" && echo "🧹 已清理旧文件: $base"
done

# --- 2. 信任 dev server 自签证书（幂等；失败不中断，证书已信任过可跳过）---
if curl -sk -o /dev/null https://localhost:3000/taskpane.html; then
  TMP_PEM="${TMPDIR:-/tmp}/qslide-cert-$$.pem"
  if openssl s_client -connect localhost:3000 -showcerts </dev/null 2>/dev/null \
    | openssl x509 -outform PEM > "$TMP_PEM" && [ -s "$TMP_PEM" ]; then
    security add-trusted-cert -d -r trustRoot \
      -k "$HOME/Library/Keychains/login.keychain-db" "$TMP_PEM" 2>/dev/null \
      && echo "✅ localhost:3000 证书已装入钥匙串并信任" \
      || echo "ℹ️ 证书信任步骤跳过（可能已信任过，不影响使用）"
  else
    echo "ℹ️ 证书提取失败，跳过（不影响 manifest 侧载）"
  fi
  rm -f "$TMP_PEM"
  echo "   （basic-ssl 证书有效期 30 天；过期或清了 node_modules/.vite 缓存后重跑本脚本）"
else
  echo "⚠️ dev server 未运行，跳过证书信任。先 npm run dev 再重跑本脚本。"
fi

echo "→ 完全退出并重新打开 PowerPoint（Cmd+Q），即可在「开始」标签页看到 QSlide 分组。"
echo "→ （可选）开机自启 dev server：launchctl load -w ~/Library/LaunchAgents/com.qslide.vite.plist"
