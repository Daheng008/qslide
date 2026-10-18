#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成「部署版」manifest：把开发态的 https://localhost:3000 全部替换为线上域名，
供 Vercel / GitHub Pages / 任意 HTTPS 静态托管使用。

用法：
  python3 scripts/gen_manifest_prod.py https://qslide.vercel.app
  python3 scripts/gen_manifest_prod.py https://<user>.github.io/qslide   # GitHub Pages 子路径

输出：manifest.prod.xml （不覆盖开发版 manifest.xml）

注意：
  - 末尾不要带斜杠；脚本会自动拼 /assets/... 等路径。
  - 部署前请先 `npm run build` 生成 dist/，再把 dist/ 上传到目标托管。
"""
import sys, re, pathlib

DEV = "https://localhost:3000"
HERE = pathlib.Path(__file__).resolve().parent.parent
SRC = HERE / "manifest.xml"
OUT = HERE / "manifest.prod.xml"

def main():
    if len(sys.argv) < 2:
        print("用法: python3 scripts/gen_manifest_prod.py <BASE_URL>")
        print("示例: python3 scripts/gen_manifest_prod.py https://qslide.vercel.app")
        sys.exit(1)
    base = sys.argv[1].rstrip("/")
    if not re.match(r"^https://", base):
        print("错误：BASE_URL 必须以 https:// 开头（Office 要求可信 HTTPS）。")
        sys.exit(1)
    if "localhost" in base or "127.0.0.1" in base:
        print("错误：部署版不能用 localhost。")
        sys.exit(1)

    text = SRC.read_text(encoding="utf-8")
    n = text.count(DEV)
    text = text.replace(DEV, base)
    OUT.write_text(text, encoding="utf-8")
    print(f"已生成 {OUT.name}：替换 {n} 处 {DEV} → {base}")
    print("下一步：")
    print("  1) npm run build  →  dist/")
    print(f"  2) 把 dist/ 部署到 {base} 根目录（保持 taskpane.html / commands.html / assets/ 相对位置）")
    print(f"  3) 在 PowerPoint 里 sideload / 集中部署 {OUT.name}")

if __name__ == "__main__":
    main()
