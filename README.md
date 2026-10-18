# QSlide v1.0.0.1

**QSlide = QPT 的交互基因 × Office.js 现代 API** —— 面向咨询级 PPT 排版的 PowerPoint 加载项（Mac/Windows/Web 一套代码）。当前状态：**9 个功能分组 50 个按钮 + 右侧任务面板，插入图形/图片/文本框全链路可用，已在 Mac PowerPoint 16.113 真机验证**。

## 目录结构

```
QSlide/
├── manifest.xml            # Office 加载项清单：CustomTab 9 组 50 按钮（v1.0.0.2）
├── install-sideload.command# 一键侧载：启动 dev server + 拷 manifest + 信任证书
├── src/
│   ├── core/ops.ts         # 核心操作库（~60 个 op，全部带能力检测与防御式报错）
│   ├── commands/commands.ts# Ribbon ExecuteFunction 入口（62 handler，统一 withEvent 包装）
│   ├── taskpane/           # 任务面板：反馈日志 + 排版体检 + 取色器 + 思维模型
│   ├── public/             # 图标（189 PNG）+ 图形画廊 shape-picker.html + 图片选择 picker.html
│   └── types/              # 类型声明
├── scripts/
│   ├── check_manifest_resources.py  # manifest 资源自检（改清单后必跑）
│   ├── gen_icons.py                 # PIL 批量生成图标（160 画布 → 80/32/16）
│   ├── apply_icon_manifest.py       # 图标 ↔ manifest resid 批量接线
│   ├── dev-server.sh                # dev server 包装脚本
│   └── sideload-test.sh             # 侧载对照测试（清缓存 + 换 manifest + 重启 PPT）
├── docs/                   # 项目文档 / 产品手册 / MVP方案 / 交互基因清单
└── 临时文件夹/              # 历史诊断产物归档（可删除）
```

## 快速开始（macOS）

```bash
npm install
bash install-sideload.command          # 启动服务 + 侧载 + 信任证书
osascript -e 'quit app "Microsoft PowerPoint"'; open -a "Microsoft PowerPoint"
```

重启 PowerPoint 后功能区出现 **QSlide 标签页**；面板版本号应与 manifest `<Version>` 一致。

## 开发约定（踩坑总结，详见 docs/项目文档.md）

- 改 `manifest.xml` 后：跑 `python3 scripts/check_manifest_resources.py`，再重跑 `install-sideload.command` + 重启 PPT（Office 会清空 Wef 目录且缓存失败状态）；
- 只改 JS/HTML 时：Cmd+Q 重启 PowerPoint 即可，无需重装；
- Ribbon 反馈一律写入 localStorage 由面板展示，**成功静默、永不弹窗**；
- 每个 handler 必须调用 `event.completed()`（已由 `withEvent` 统一兜底）。

## 文档

| 文档 | 面向 |
|---|---|
| [docs/项目文档.md](docs/项目文档.md) | 开发者：架构、数据流、清单三坑、排障手册 |
| [docs/产品手册.md](docs/产品手册.md) | 用户：安装、各功能用法、FAQ |
| [docs/交互基因清单.md](docs/交互基因清单.md) | 设计基因（QPT 对标） |
| [docs/MVP方案.md](docs/MVP方案.md) | 早期方案存档 |
