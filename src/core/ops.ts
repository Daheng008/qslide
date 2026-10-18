/**
 * QSlide 核心操作（Office.js 实现）
 *
 * 设计原则（源自 QPT 交互基因）：
 *  - 短路径：一次点击直接到位，无对话框；
 *  - 防御式：PowerPoint JS API 各版本能力差异大，每个操作都做能力检测，
 *    失败时给出「为什么 + 怎么办」的明确提示，绝不静默失败。
 */

export interface OpResult {
  ok: boolean;
  op: string;
  detail: string;
  /** 命中的 PowerPointApi 能力集（诊断用） */
  api?: string;
}

export function apiSupported(setName: string, min: string): boolean {
  try {
    return !!Office?.context?.requirements?.isSetSupported?.(setName, min);
  } catch {
    return false;
  }
}

/** 选取形状集合。返回 null 表示当前主机不支持选区 API。
 *  ⚠️ 必须在同一个 PowerPoint.run 内使用返回的 items——run 结束后代理对象即脱钩，
 *  跨 run 修改脱钩代理会静默失效（sync 无效果且不报错）。 */
async function getSelectedShapesIn(ctx: any): Promise<any[] | null> {
  if (!apiSupported("PowerPointApi", "1.7")) return null;
  const sel = ctx.presentation.getSelectedShapes();
  // ⚠️ autoShapeType 必须显式 load，否则读到 undefined，shapeKey 会退化成
  // 粗粒度的 type 匹配（GeometricShape 大类），把三角形也算成矩形的"同类"。
  sel.load("items/shapeId,type,autoShapeType,left,top,width,height,name");
  await ctx.sync();
  return sel.items;
}

function fmtPt(v: number): string {
  return Math.round(v) + "pt";
}

/* ─────────── 同类模式（宿主不支持 setSelectedShapes 的变通方案） ───────────
 * Mac PowerPoint JS API 能读选区但不能写选区，故「选择相似形状」无法真正改选区。
 * 变通：开启同类模式后，填充/删除文本等批量样式操作自动作用于
 * 当前页所有与首个选中形状同类型的形状，效果等价于"已选中它们"。 */

let similarMode = false;

export function setSimilarMode(v: boolean): void {
  similarMode = v;
}
export function getSimilarMode(): boolean {
  return similarMode;
}

/** 相似性判据（三级回退，粒度从细到粗）：
 *  1. autoShapeType（如 rectangle / isoscelesTriangle）——理想判据；
 *  2. 形状名前缀——PowerPoint 自动命名为「类型名+序号」（如「等腰三角形 2」「Rectangle 4」），
 *     Mac 宿主实测 autoShapeType 恒为 none/undefined 时用此判据；
 *  3. shape.type（GeometricShape 等大类）——仅作兜底，会把同类不同形都算上。
 *  ⚠️ 用户手动改过名的形状用判据 2 会失配（日志会显示判据便于诊断）。 */
function shapeKey(s: any): string {
  const ast = s.autoShapeType;
  if (ast && ast !== "none" && ast !== "undefined") return "ast:" + ast;
  const name: string = s.name || "";
  const kind = name.replace(/\s*\d+\s*$/, "").trim();
  if (kind && s.type === "GeometricShape") return "name:" + kind;
  return "type:" + s.type;
}

/** 计算本操作的目标形状集合：普通模式=选区；同类模式=当前页全部同类形状。
 *  返回的代理对象与 ctx 同上下文，可直接写入。 */
async function targetsIn(ctx: any, selected: any[]): Promise<{ list: any[]; note: string }> {
  if (!similarMode || selected.length === 0) return { list: selected, note: "" };
  const slides = ctx.presentation.getSelectedSlides?.();
  if (!slides) return { list: selected, note: "" };
  slides.load("items");
  await ctx.sync();
  if (slides.items.length === 0) return { list: selected, note: "" };
  const all = slides.items[0].shapes;
  all.load("items/shapeId,type,autoShapeType,name");
  await ctx.sync();
  const base = shapeKey(selected[0]);
  const matched = all.items.filter((s: any) => shapeKey(s) === base);
  if (matched.length <= 1) return { list: selected, note: "" };
  return { list: matched, note: `（同类模式 ×${matched.length}，判据 ${base.replace(":", "=")}）` };
}

/** 同宽/同高：以集合中的第一个形状为基准（VBA 版同语义）。读取+写入必须同一 run。 */
export async function opSameSize(dim: "width" | "height"): Promise<OpResult> {
  const opName = "同" + (dim === "width" ? "宽" : "高");
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+），请更新 Microsoft 365。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const shapes = await getSelectedShapesIn(ctx);
    if (shapes === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+），请更新 Microsoft 365。", api: "1.7" };
    if (shapes.length < 2)
      return { ok: false, op: opName, detail: "请先选中 2 个及以上形状。" };
    const target = shapes[0][dim];
    for (const s of shapes) s[dim] = target;
    await ctx.sync();
    return { ok: true, op: opName, detail: `${shapes.length} 个形状 → ${fmtPt(target)}`, api: "1.3" };
  });
}

/**
 * 贴合对齐（QPT Align Touch 的 JS 版）：
 * axis="h"：按 left 排序，右缘依次贴住下一个左缘，间隙归零；垂直中心取原分布的加权平均。
 * axis="v"：同理按 top 排序纵向堆叠。
 * 需要 PowerPointApi 1.7+（选区）与形状几何可写。
 */
export async function opAlignTouch(axis: "h" | "v"): Promise<OpResult> {
  const opName = axis === "h" ? "贴合对齐↔" : "贴合对齐↕";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+），请更新 Microsoft 365。", api: "1.7" };

  return await PowerPoint.run(async (ctx: any) => {
    const shapes = await getSelectedShapesIn(ctx);
    if (shapes === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+），请更新 Microsoft 365。", api: "1.7" };
    if (shapes.length < 2)
      return { ok: false, op: opName, detail: "请先选中 2 个及以上形状。" };

    const sorted = [...shapes].sort((a, b) => (axis === "h" ? a.left - b.left : a.top - b.top));
    const primary = axis === "h" ? "left" : "top";
    const secondary = axis === "h" ? "top" : "left";
    const extent = axis === "h" ? "width" : "height";

    // 次轴对齐：取基准形状的次轴坐标（第一个，与 VBA 版一致）
    const secondaryTarget = sorted[0][secondary];

    let cursor = sorted[0][primary];
    for (const s of sorted) {
      s[primary] = cursor;
      s[secondary] = secondaryTarget;
      cursor += s[extent];
    }
    await ctx.sync();
    return {
      ok: true,
      op: opName,
      detail: `${shapes.length} 个形状已贴合，间隙 0`,
      api: "1.7",
    };
  });
}

/** 删除所选形状文本。同一 run 内完成。
 *  ⚠️ Mac 实测坑：textRange.setText("") 会抛 InvalidArgument（空字符串参数被拒绝），
 *  必须用属性赋值 textRange.text = ""；且无文本框形状的 sync 会整体失败，
 *  故逐形状独立 sync 隔离错误，避免一个图片形状拖垮整批。 */
export async function opRemoveText(): Promise<OpResult> {
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: "删除文本", detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };

  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null)
      return { ok: false, op: "删除文本", detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
    if (sel.length === 0)
      return { ok: false, op: "删除文本", detail: "请先选中至少 1 个形状。" };
    const { list: shapes, note } = await targetsIn(ctx, sel);

    let cleared = 0;
    let skipped = 0;
    for (const s of shapes) {
      try {
        s.textFrame.textRange.text = "";
        await ctx.sync();
        cleared++;
      } catch {
        skipped++; // 无文本框的形状（图片/线条等）或宿主不支持，单独跳过
      }
    }
    if (cleared === 0)
      return { ok: false, op: "删除文本", detail: "所选形状均无文本框（或当前 API 版本不支持，需 1.5+）。" };
    return {
      ok: true,
      op: "删除文本",
      detail: note + (skipped > 0 ? `${cleared}/${shapes.length} 个形状已清空（${skipped} 个无文本框已跳过）` : `${cleared}/${shapes.length} 个形状已清空`),
      api: "1.5",
    };
  });
}

/** 自研咨询风色板（不复用 QPT 色板值）：深藏青/品牌蓝/青绿/琥珀/砖红/灰蓝 */
export const PALETTE: { hex: string; name: string }[] = [
  { hex: "0F2B46", name: "深藏青" },
  { hex: "1B6CA8", name: "品牌蓝" },
  { hex: "2E9E97", name: "青绿" },
  { hex: "E8A33D", name: "琥珀" },
  { hex: "C0504D", name: "砖红" },
  { hex: "8496A9", name: "灰蓝" },
];

/** 颜色工具：解析 #RRGGBB 或 rgb(r,g,b) → [r,g,b]；中性色判定；最近色板色。供 opUnifyPalette 使用。 */
function colorToRgb(c: string): [number, number, number] | null {
  c = c.trim();
  if (/^[0-9a-fA-F]{6}$/.test(c)) c = "#" + c; // 裸 hex（宿主可能不带 #）
  if (c.startsWith("#")) {
    let h = c.slice(1);
    if (h.length === 3) h = h.split("").map((x) => x + x).join("");
    if (h.length !== 6) return null;
    const n = parseInt(h, 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  const m = c.match(/^rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i);
  return m ? [+m[1], +m[2], +m[3]] : null;
}
function isNeutral(rgb: [number, number, number]): boolean {
  return Math.max(...rgb) - Math.min(...rgb) < 30; // 白/灰/黑等接近无色
}
function nearestPalette(hex: string): string {
  const c = colorToRgb(hex)!;
  let best = PALETTE[0].hex, bestD = Infinity;
  for (const p of PALETTE) {
    const pr = colorToRgb("#" + p.hex)!;
    const d = (pr[0] - c[0]) ** 2 + (pr[1] - c[1]) ** 2 + (pr[2] - c[2]) ** 2;
    if (d < bestD) { bestD = d; best = p.hex; }
  }
  return best;
}

/** 色板应用：对所选形状填充纯色。同一 run 内读写；无填充面的形状（图片等）跳过。 */
export async function opApplyFill(hex: string): Promise<OpResult> {
  const opName = "填充 " + hex;
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
    if (sel.length === 0)
      return { ok: false, op: opName, detail: "请先选中至少 1 个形状。" };
    const { list: shapes, note } = await targetsIn(ctx, sel);
    let painted = 0;
    for (const s of shapes) {
      try {
        s.fill.setSolidColor("#" + hex);
        painted++;
      } catch {
        /* 图片/线条等无填充面形状跳过 */
      }
    }
    await ctx.sync();
    if (painted === 0)
      return { ok: false, op: opName, detail: "所选形状均不支持纯色填充。" };
    return { ok: true, op: opName, detail: note + ` ${painted}/${shapes.length} 个形状已填充 #${hex}`, api: "1.3" };
  });
}

/** 等间距分布：n≥3，首尾形状固定，中间按间隙均分排列。 */
export async function opDistribute(axis: "h" | "v"): Promise<OpResult> {
  const opName = axis === "h" ? "等间距 ↔" : "等间距 ↕";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+），请更新 Microsoft 365。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const shapes = await getSelectedShapesIn(ctx);
    if (shapes === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+），请更新 Microsoft 365。", api: "1.7" };
    if (shapes.length < 3)
      return { ok: false, op: opName, detail: "等间距分布至少需要选中 3 个形状。" };

    const primary = axis === "h" ? "left" : "top";
    const extent = axis === "h" ? "width" : "height";
    const sorted = [...shapes].sort((a, b) => a[primary] - b[primary]);

    const start = sorted[0][primary];
    const end = sorted[sorted.length - 1][primary] + sorted[sorted.length - 1][extent];
    const sizeSum = sorted.reduce((acc: number, s: any) => acc + s[extent], 0);
    const gap = (end - start - sizeSum) / (sorted.length - 1);
    if (gap < 0)
      return { ok: false, op: opName, detail: "首尾间距不足以容纳中间形状（出现重叠），请先拉开首尾距离。" };

    let cursor: number = start;
    for (const s of sorted) {
      s[primary] = cursor;
      cursor += s[extent] + gap;
    }
    await ctx.sync();
    return { ok: true, op: opName, detail: `${shapes.length} 个形状，间距 ${fmtPt(gap)}`, api: "1.7" };
  });
}

/** 交换位置：交换两个形状的 left/top。 */
export async function opSwap(): Promise<OpResult> {
  const opName = "交换位置";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const shapes = await getSelectedShapesIn(ctx);
    if (shapes === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
    if (shapes.length !== 2)
      return { ok: false, op: opName, detail: "请恰好选中 2 个形状。" };

    const [a, b] = shapes;
    // 必须在任何写入前把两个坐标都取出来（写入后同 run 内的读值语义不可依赖）
    const ax = a.left, ay = a.top, bx = b.left, by = b.top;
    a.left = bx; a.top = by;
    b.left = ax; b.top = ay;
    await ctx.sync();
    return { ok: true, op: opName, detail: "2 个形状位置已互换", api: "1.7" };
  });
}

/** 选择相似形状：以第一个选中形状为基准（自选图形类型或形状类型），选中当前页所有同类形状。 */
export async function opSelectSimilar(): Promise<OpResult> {
  const opName = "选择相似形状";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const shapes = await getSelectedShapesIn(ctx);
    if (shapes === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
    if (shapes.length !== 1)
      return { ok: false, op: opName, detail: "请恰好选中 1 个基准形状。" };

    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides)
      return { ok: false, op: opName, detail: "宿主不支持获取当前页（getSelectedSlides）。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0)
      return { ok: false, op: opName, detail: "未获取到当前页。" };

    const base = shapeKey(shapes[0]);

    const all = slides.items[0].shapes;
    all.load("items/shapeId,type,autoShapeType,name");
    await ctx.sync();
    const matched = all.items.filter((s: any) => shapeKey(s) === base);
    if (matched.length <= 1)
      return { ok: false, op: opName, detail: "当前页没有其他同类型形状。" };

    if (typeof ctx.presentation.setSelectedShapes !== "function") {
      // Mac 等宿主不支持写入选区 → 自动开启同类模式作为变通
      setSimilarMode(true);
      return {
        ok: true,
        op: opName,
        detail: `找到 ${matched.length} 个同类形状（${base.replace(":", " ")}）。宿主不支持写入选区，已自动开启【同类模式】：填充/删除文本将作用于这 ${matched.length} 个形状。`,
        api: "1.7",
      };
    }
    try {
      ctx.presentation.setSelectedShapes(matched.map((s: any) => s.shapeId));
      await ctx.sync();
    } catch (e: any) {
      setSimilarMode(true);
      return { ok: false, op: opName, detail: `找到 ${matched.length} 个同类形状，写入选区失败（${e?.message || e}），已自动开启【同类模式】作为变通。` };
    }
    return { ok: true, op: opName, detail: `已选中 ${matched.length} 个同类形状（${base.replace(":", " ")}）`, api: "1.7" };
  });
}

/** Action Title：在当前页顶部插入行动式标题文本框（结论先行占位 + 咨询风样式）。 */
export async function opActionTitle(): Promise<OpResult> {
  const opName = "Action Title";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "插入文本框需 PowerPointApi 1.8+（addTextBox），请更新 Microsoft 365。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides)
      return { ok: false, op: opName, detail: "宿主不支持获取当前页（getSelectedSlides）。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0)
      return { ok: false, op: opName, detail: "未获取到当前页。" };

    const slide = slides.items[0];
    const tb = slide.shapes.addTextBox("行动式标题：把本页结论写成一句话（编辑替换此占位文本）");
    tb.left = 36; tb.top = 20; tb.width = 700; tb.height = 52;
    const font = tb.textFrame.textRange.font;
    font.bold = true;
    font.size = 20;
    font.color = "#0F2B46";
    await ctx.sync();
    return { ok: true, op: opName, detail: "已在当前页顶部插入标题文本框（36,20 700×52pt）", api: "1.8" };
  });
}

/** 统一本页色系：把当前页所有纯色填充形状 snap 到最近的咨询色板色。
 *  中性色（白/灰/黑，max-min<30）与无纯色填充的形状自动跳过，避免把白底/背景变成蓝块。 */
export async function opUnifyPalette(): Promise<OpResult> {
  const opName = "统一本页色系";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+，请更新 Microsoft 365。", api: "1.7" };

  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页（getSelectedSlides）。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "请先选中/激活一张幻灯片。" };

    const shapes = slides.items[0].shapes;
    // ⚠️ 集合 load 的嵌套属性也必须带 items/ 前缀——"fill/type" 会被静默忽略，
    // fill.type 读出 undefined → 全部被误判为"无纯色填充"（实测 9/9 跳过）。
    shapes.load("items/shapeId,items/fill/type,items/fill/foregroundColor");
    await ctx.sync();

    let changed = 0, skipped = 0, neutral = 0, inPalette = 0;
    const typeCounts: Record<string, number> = {};
    for (const s of shapes.items) {
      const fc: string | undefined = s.fill?.foregroundColor;
      const type: string | undefined = s.fill?.type;
      const tk = type ?? "undefined";
      typeCounts[tk] = (typeCounts[tk] || 0) + 1;
      // ⚠️ Mac 宿主 fill.type 返回 PascalCase（"Solid"/"NoFill"），与文档枚举
      // "solid"/"noFill" 大小写不符——必须大小写不敏感比较，否则纯色形状全被误跳过。
      if (!fc || (type || "").toLowerCase() !== "solid") { skipped++; continue; } // 无填充/渐变/图片
      const rgb = colorToRgb(fc);
      if (!rgb) { skipped++; continue; }
      if (isNeutral(rgb)) { neutral++; continue; } // 白/灰/黑底不碰
      const target = "#" + nearestPalette(fc);
      if (target.toLowerCase() === fc.toLowerCase()) { inPalette++; continue; } // 已在色板内（幂等）
      s.fill.setSolidColor(target);
      changed++;
    }
    await ctx.sync();
    const dist = Object.entries(typeCounts).map(([k, v]) => `${k}×${v}`).join("，");
    if (changed === 0)
      return { ok: false, op: opName, detail: `无需调整（${dist}；其中 ${inPalette} 个已在色板，${neutral} 中性色，${skipped} 无纯色填充）` };
    return { ok: true, op: opName, detail: `${changed} 个形状已 snap 到最近咨询色（${inPalette} 个已在色板，${neutral} 中性色 / ${skipped} 无纯色填充跳过）`, api: "1.7" };
  });
}

/** 四向边缘贴合（QPT 基因：向上/向下/向左/向右贴合）：
 *  把所有选中形状朝指定方向推齐到选区极值边缘（up: 顶缘取 min(top)；down: 底缘取 max(bottom)；依此类推）。
 *  与 opAlignTouch（串联堆叠）互补：这是「对齐到一条边」，那是「依次紧贴」。 */
export async function opTouchEdge(dir: "up" | "down" | "left" | "right"): Promise<OpResult> {
  const opName = { up: "向上贴合", down: "向下贴合", left: "向左贴合", right: "向右贴合" }[dir];
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const shapes = await getSelectedShapesIn(ctx);
    if (shapes === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
    if (shapes.length < 2)
      return { ok: false, op: opName, detail: "请先选中 2 个及以上形状。" };
    if (dir === "up" || dir === "down") {
      const edge = dir === "up" ? Math.min(...shapes.map((s) => s.top)) : Math.max(...shapes.map((s) => s.top + s.height)) ;
      for (const s of shapes) s.top = dir === "up" ? edge : edge - s.height;
    } else {
      const edge = dir === "left" ? Math.min(...shapes.map((s) => s.left)) : Math.max(...shapes.map((s) => s.left + s.width));
      for (const s of shapes) s.left = dir === "left" ? edge : edge - s.width;
    }
    await ctx.sync();
    return { ok: true, op: opName, detail: `${shapes.length} 个形状已贴合到选区${opName.slice(2)}边缘`, api: "1.7" };
  });
}

/** 主题标题：与 Action Title 配对——行动式标题写「结论」，主题标题写「本页讨论什么」。
 *  样式区分：18pt 不加粗 深藏青，位于页顶（若已有 Action Title 会覆盖其位置，用户自行挪开即可）。 */
export async function opSubjectTitle(): Promise<OpResult> {
  const opName = "主题标题";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "插入文本框需 PowerPointApi 1.8+（addTextBox）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页（getSelectedSlides）。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const tb = slides.items[0].shapes.addTextBox("主题标题：本页讨论的对象/主题（编辑替换此占位文本）");
    tb.left = 36; tb.top = 20; tb.width = 700; tb.height = 40;
    const font = tb.textFrame.textRange.font;
    font.bold = false;
    font.size = 18;
    font.color = "#0F2B46";
    await ctx.sync();
    return { ok: true, op: opName, detail: "已在当前页顶部插入主题标题（18pt 常规体）", api: "1.8" };
  });
}

/** 脚注：当前页底部左侧插入来源/注释框（10pt 灰色，咨询稿标配）。 */
export async function opFootnote(): Promise<OpResult> {
  const opName = "脚注";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "插入文本框需 PowerPointApi 1.8+（addTextBox）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页（getSelectedSlides）。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "未获取到当前页。" };
    // 16:9 页面 960×540pt，脚注落底部安全区
    const tb = slides.items[0].shapes.addTextBox("来源：数据出处/备注（脚注占位，编辑替换）");
    tb.left = 36; tb.top = 508; tb.width = 600; tb.height = 22;
    const font = tb.textFrame.textRange.font;
    font.size = 10;
    font.color = "#7F8C99";
    font.italic = true;
    await ctx.sync();
    return { ok: true, op: opName, detail: "已在页底插入脚注框（10pt 灰色斜体）", api: "1.8" };
  });
}

/** 合并文本（真合并）：把所选形状的文字按选区顺序合并进第一个含文本的形状
 *  （换行分隔），该形状扩到所选形状的整体包围盒，其余形状删除——N 个图元
 *  真正合为 1 个，而不是新建文本框保留原图元。
 *  ⚠️ 格式保护：写 textRange.text 会重置 run 结构，写前快照锚点字体属性、写后回写。 */
export async function opMergeText(): Promise<OpResult> {
  const opName = "合并文本";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
    if (sel.length < 2)
      return { ok: false, op: opName, detail: "请选中 2 个及以上含文本的形状。" };
    // 追加 load 文本内容 + 几何 + 锚点字体快照字段（未 load 的属性读取会抛错）
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/left,items/top,items/width,items/height," +
      "items/textFrame/textRange/text," +
      "items/textFrame/textRange/font/name,items/textFrame/textRange/font/size," +
      "items/textFrame/textRange/font/bold,items/textFrame/textRange/font/italic," +
      "items/textFrame/textRange/font/underline,items/textFrame/textRange/font/color");
    await ctx.sync();

    const parts: string[] = [];
    let anchor: any = null;
    let minL = Infinity, minT = Infinity, maxR = -Infinity, maxB = -Infinity;
    for (const s of sel2.items) {
      const l = s.left ?? 0, t = s.top ?? 0, w = s.width ?? 0, h = s.height ?? 0;
      if (l < minL) minL = l;
      if (t < minT) minT = t;
      if (l + w > maxR) maxR = l + w;
      if (t + h > maxB) maxB = t + h;
      const txt: string | undefined = s.textFrame?.textRange?.text;
      if (txt && txt.trim()) {
        parts.push(txt.replace(/\r+\n?/g, "\n").trim());
        if (!anchor) anchor = s; // 选区顺序中第一个含文本的形状作为锚点（格式基准）
      }
    }
    if (!anchor || parts.length < 2)
      return { ok: false, op: opName, detail: "所选形状中有效文本不足 2 段（无文本框/空文本已忽略）。" };

    // 字体快照（null = 混合格式，无法恢复该属性）
    const f = anchor.textFrame.textRange.font;
    const snap = { name: f.name, size: f.size, bold: f.bold, italic: f.italic, underline: f.underline, color: f.color };

    anchor.textFrame.textRange.text = parts.join("\n");
    if (typeof snap.name === "string") f.name = snap.name;
    if (typeof snap.size === "number") f.size = snap.size;
    if (typeof snap.bold === "boolean") f.bold = snap.bold;
    if (typeof snap.italic === "boolean") f.italic = snap.italic;
    if (typeof snap.underline === "string") f.underline = snap.underline;
    if (typeof snap.color === "string") f.color = snap.color;

    // 锚点扩到所选形状的整体包围盒
    anchor.left = minL;
    anchor.top = minT;
    anchor.width = maxR - minL;
    anchor.height = maxB - minT;

    // 先提交文本合并与几何扩展（避免后续 delete 的 sync 错误波及这部分写入）
    await ctx.sync();

    // 删除其余形状：先按选区代理删；Mac 宿主的 delete 失败不在调用点抛错，
    // 而是在 sync() 时浮出，且可能静默失效——故删除后用页面形状集合校验补删。
    const toDelete: number[] = sel2.items
      .filter((s: any) => s.shapeId !== anchor.shapeId)
      .map((s: any) => s.shapeId);
    for (const s of sel2.items) {
      if (s.shapeId === anchor.shapeId) continue;
      try { s.delete(); } catch { /* 错误在 sync 时浮出 */ }
    }
    try {
      await ctx.sync();
    } catch { /* 部分删除失败，走下面的校验补删 */ }

    let delFail = 0;
    const slides = ctx.presentation.getSelectedSlides?.();
    if (slides) {
      slides.load("items");
      await ctx.sync();
      const slideShapes = slides.items[0]?.shapes;
      if (slideShapes) {
        slideShapes.load("items/shapeId");
        await ctx.sync();
        const remain = slideShapes.items.filter((s: any) => toDelete.includes(s.shapeId));
        if (remain.length > 0) {
          for (const s of remain) {
            try { s.delete(); } catch { /* 同上 */ }
          }
          try {
            await ctx.sync();
            delFail = 0;
          } catch {
            delFail = remain.length; // 补删仍失败，如实报告
          }
        }
      }
    }
    const deleted = toDelete.length - delFail;

    const mixedNote = (snap.size === null || snap.bold === null || snap.color === null)
      ? "；锚点为混合格式，已按可恢复属性均一化" : "";
    if (deleted === 0)
      return { ok: false, op: opName, detail: `文本已合并进锚点形状，但原图元删除失败（宿主不支持 shape.delete 或静默失效），请手动删除其余 ${toDelete.length} 个。` };
    return {
      ok: true,
      op: opName,
      detail: `${parts.length} 段文本已合并为 1 个图元（删除原图元 ${deleted} 个${delFail ? `，${delFail} 个删除失败请手动删` : ""}）${mixedNote}`,
      api: "1.7",
    };
  });
}

/** 删除重复空格：清洗所选形状文本中的连续空白为单个空格（同 run 逐形状隔离失败）。
 *  ⚠️ 格式保护：直接 textRange.text = cleaned 会重置 run 结构、丢掉加粗/颜色/字号，
 *  故替换前快照 range 级字体属性、替换后回写（单一格式形状可完整恢复；
 *  混合格式形状会被均一化为首个格式——日志中已提示）。 */
export async function opCollapseSpaces(): Promise<OpResult> {
  const opName = "删除重复空格";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
    if (sel.length === 0)
      return { ok: false, op: opName, detail: "请先选中至少 1 个形状。" };
    // 读文本 + 字体快照字段
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/textFrame/textRange/text," +
      "items/textFrame/textRange/font/name,items/textFrame/textRange/font/size," +
      "items/textFrame/textRange/font/bold,items/textFrame/textRange/font/italic," +
      "items/textFrame/textRange/font/underline,items/textFrame/textRange/font/color");
    await ctx.sync();
    let fixed = 0, skipped = 0, homogenized = 0;
    for (const s of sel2.items) {
      const t: string | undefined = s.textFrame?.textRange?.text;
      if (typeof t !== "string" || !t) { skipped++; continue; }
      const cleaned = t.replace(/[ \t]{2,}/g, " ").replace(/[ \t]+([,.;:!?，。；：！？])/g, "$1");
      if (cleaned === t) continue; // 本就没有重复空格
      try {
        const font = s.textFrame.textRange.font;
        // 快照（undefined 表示 range 内格式不一，回写时跳过该字段）
        const snap = {
          name: font.name, size: font.size, bold: font.bold,
          italic: font.italic, underline: font.underline, color: font.color,
        };
        s.textFrame.textRange.text = cleaned;
        // 回写格式
        if (snap.name) font.name = snap.name;
        if (snap.size != null) font.size = snap.size;
        if (snap.bold != null) font.bold = snap.bold;
        if (snap.italic != null) font.italic = snap.italic;
        if (snap.underline != null) font.underline = snap.underline;
        if (snap.color) font.color = snap.color;
        await ctx.sync();
        fixed++;
        if (Object.values(snap).some((v) => v === null)) homogenized++;
      } catch { skipped++; }
    }
    if (fixed === 0)
      return { ok: false, op: opName, detail: `无需清洗（${skipped} 个形状无文本或本就没有重复空格）。` };
    return {
      ok: true,
      op: opName,
      detail: homogenized > 0
        ? `${fixed} 个形状文本已清洗（格式已保留；${homogenized} 个混合格式形状被均一化）`
        : `${fixed} 个形状文本已清洗（格式已保留）`,
      api: "1.5",
    };
  });
}

/** 移除轮廓：隐藏所选形状的边框线（line.format.visible=false）。无边框的形状静默跳过。 */
export async function opRemoveOutline(): Promise<OpResult> {
  const opName = "移除轮廓";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
    if (sel.length === 0)
      return { ok: false, op: opName, detail: "请先选中至少 1 个形状。" };
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/line/type");
    await ctx.sync();
    let done = 0, skipped = 0;
    for (const s of sel2.items) {
      try {
        s.line.format.visible = false;
        await ctx.sync();
        done++;
      } catch {
        skipped++; // 无边框的形状（图片等）跳过
      }
    }
    if (done === 0)
      return { ok: false, op: opName, detail: `所选形状均无边框可移除（${skipped} 个已跳过）。` };
    return { ok: true, op: opName, detail: `${done}/${sel2.items.length} 个形状已去边框${skipped ? `（${skipped} 个无边框跳过）` : ""}`, api: "1.3" };
  });
}

/** 插入文本框：在当前页（或选区中心）插入一个待编辑文本框，方便快速补文字。 */
export async function opInsertTextbox(): Promise<OpResult> {
  const opName = "插入文本框";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addTextBox）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    const slide = slides.items[0];
    if (!slide) return { ok: false, op: opName, detail: "未选中幻灯片页。" };
    // 默认置于页面中上部（Mac 幻灯片 720×405pt 的常见工作区）
    const tb = slide.shapes.addTextBox("输入文字");
    tb.left = 280; tb.top = 150; tb.width = 300; tb.height = 56;
    const font = tb.textFrame.textRange.font;
    font.size = 16; font.color = "#333333";
    await ctx.sync();
    return { ok: true, op: opName, detail: "已在当前页插入文本框（双击编辑文字）", api: "1.8" };
  });
}

/** 统一字体：把除首个外的所有选中形状文本，套用「首个含文本形状」的字体名/字号/颜色。
 *  无需弹窗，以选区第一项为基准，适合快速统一一页内的字体。 */
export async function opUnifyFont(): Promise<OpResult> {
  const opName = "统一字体";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
    if (sel.length < 2)
      return { ok: false, op: opName, detail: "请选中 2 个及以上形状。" };
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/textFrame/textRange/text," +
      "items/textFrame/textRange/font/name,items/textFrame/textRange/font/size," +
      "items/textFrame/textRange/font/bold,items/textFrame/textRange/font/color");
    await ctx.sync();
    // 找基准：第一个含非空文本的形状
    let base: any = null;
    for (const s of sel2.items) {
      const t: string | undefined = s.textFrame?.textRange?.text;
      if (t && t.trim()) { base = s; break; }
    }
    if (!base) return { ok: false, op: opName, detail: "所选形状均无可套用的字体基准（需至少一个含文本）。" };
    const bf = base.textFrame.textRange.font;
    const snap = { name: bf.name, size: bf.size, bold: bf.bold, color: bf.color };
    let applied = 0, skipped = 0;
    for (const s of sel2.items) {
      if (s.shapeId === base.shapeId) continue;
      try {
        const f = s.textFrame.textRange.font;
        if (typeof snap.name === "string") f.name = snap.name;
        if (typeof snap.size === "number") f.size = snap.size;
        if (typeof snap.bold === "boolean") f.bold = snap.bold;
        if (typeof snap.color === "string") f.color = snap.color;
        await ctx.sync();
        applied++;
      } catch { skipped++; }
    }
    if (applied === 0)
      return { ok: false, op: opName, detail: `无需统一（${skipped} 个形状无文本框已跳过）。` };
    const baseInfo = [snap.name, snap.size ? snap.size + "pt" : null].filter(Boolean).join(" / ");
    return { ok: true, op: opName, detail: `${applied} 个形状已套用基准字体（${baseInfo}）${skipped ? `（${skipped} 个无文本框跳过）` : ""}`, api: "1.7" };
  });
}

/** 复制到所有幻灯片：把所选形状（文本/自选图形）克隆到除当前页外的每一张幻灯片。
 *  常用于页眉/页脚/Logo 一键铺满。绘制类（图片等）跳过；逐个隔离失败。 */
export async function opCopyToAllSlides(): Promise<OpResult> {
  const opName = "复制到所有幻灯片";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addTextBox/addGeometricShape）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null)
      return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 PowerPointApi 1.7+）。", api: "1.7" };
    if (sel.length === 0)
      return { ok: false, op: opName, detail: "请先选中至少 1 个形状。" };
    // 读取源形状几何 + 填充 + 文本 + 字体
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/autoShapeType,items/left,items/top,items/width,items/height," +
      "items/fill/type,items/fill/foregroundColor," +
      "items/textFrame/textRange/text,items/textFrame/textRange/font/name,items/textFrame/textRange/font/size," +
      "items/textFrame/textRange/font/bold,items/textFrame/textRange/font/color");
    await ctx.sync();

    // 当前选中幻灯片 id（克隆时跳过自身）
    const cur = ctx.presentation.getSelectedSlides?.();
    let curId: any = null;
    if (cur) { cur.load("items/id"); await ctx.sync(); curId = cur.items[0]?.id; }

    const all = ctx.presentation.slides;
    all.load("items/id");
    await ctx.sync();

    let copied = 0, skipped = 0, slideCount = 0;
    for (const slide of all.items) {
      slideCount++;
      if (slide.id === curId) { skipped++; continue; } // 当前页不重复
      for (const src of sel2.items) {
        try {
          const ast = src.autoShapeType;
          let dst: any;
          if (ast && ast !== "none" && ast !== undefined) {
            dst = slide.shapes.addGeometricShape(ast);
          } else {
            dst = slide.shapes.addTextBox(src.textFrame?.textRange?.text || "");
          }
          dst.left = src.left; dst.top = src.top; dst.width = src.width; dst.height = src.height;
          if (src.fill?.type === "solid" && src.fill?.foregroundColor) {
            try { dst.fill.setSolidColor(src.fill.foregroundColor); } catch { /* 填充失败忽略 */ }
          }
          const txt: string | undefined = src.textFrame?.textRange?.text;
          if (txt) {
            const sf = src.textFrame.textRange.font;
            const df = dst.textFrame.textRange.font;
            if (typeof sf.name === "string") df.name = sf.name;
            if (typeof sf.size === "number") df.size = sf.size;
            if (typeof sf.bold === "boolean") df.bold = sf.bold;
            if (typeof sf.color === "string") df.color = sf.color;
          }
          copied++;
        } catch {
          skipped++;
        }
      }
    }
    await ctx.sync();
    if (copied === 0)
      return { ok: false, op: opName, detail: `未能克隆（${skipped} 次失败；多为图片等非文本/非自选图形，或宿主 API 限制）。` };
    return { ok: true, op: opName, detail: `${sel2.items.length} 个形状已克隆到 ${slideCount - 1} 张其他幻灯片（${copied} 次复制${skipped ? `，${skipped} 次跳过` : ""}）`, api: "1.8" };
  });
}

/**
 * 能力探测（Phase 0 产出）：逐项检查宿主 API 能力，驱动任务窗格的能力矩阵。
 */
export function probeCapabilities(): { name: string; min: string; ok: boolean }[] {
  const sets: [string, string, string][] = [
    ["基础演示读写", "1.1", "PowerPointApi"],
    ["形状增删/枚举", "1.3", "PowerPointApi"],
    ["形状几何(textFrame等)", "1.5", "PowerPointApi"],
    ["选区(SelectedShapes)", "1.7", "PowerPointApi"],
    ["形状插入(addTextBox)", "1.8", "PowerPointApi"],
    ["最新能力集", "1.9", "PowerPointApi"],
  ];
  return sets.map(([name, min, set]) => ({ name, min, ok: apiSupported(set, min) }));
}

/* ===================== A 组：形状样式 ===================== */

/** 调整圆角：对选中几何形状的第 0 调整点（圆角比例，0~0.5）按 delta 增减。
 *  用 Shape.adjustments（PowerPointApi 1.10+）：count / get(index) / set(index, value)。 */
export async function opAdjustRounding(delta: number): Promise<OpResult> {
  const opName = delta >= 0 ? "增大圆角" : "减小圆角";
  if (!apiSupported("PowerPointApi", "1.10"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.10+（形状调整点 API），请更新 Microsoft 365。", api: "1.10" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null) return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 1.7+）。", api: "1.7" };
    if (sel.length === 0) return { ok: false, op: opName, detail: "请先选中至少 1 个圆角形状。" };
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/type");
    await ctx.sync();
    let done = 0, skipped = 0;
    for (const s of sel2.items) {
      try {
        if (s.type !== "GeometricShape") { skipped++; continue; }
        const adj = s.adjustments;
        adj.load("count");
        await ctx.sync();
        if (!adj.count || adj.count < 1) { skipped++; continue; }
        const r = adj.get(0);   // 圆角矩形的第 0 调整点 = 圆角比例
        await ctx.sync();
        const v = r.value;
        if (typeof v !== "number") { skipped++; continue; }
        adj.set(0, Math.max(0, Math.min(0.5, v + delta)));
        await ctx.sync();
        done++;
      } catch { skipped++; }
    }
    if (done === 0) return { ok: false, op: opName, detail: `所选形状无圆角可调（${skipped} 个跳过；仅几何形状支持，需 1.10 API）。` };
    return { ok: true, op: opName, detail: `${done} 个形状圆角已${delta >= 0 ? "增大" : "减小"}${skipped ? `（${skipped} 个无圆角跳过）` : ""}`, api: "1.10" };
  });
}

/** 拉直形状：把选中几何形状的第 0 调整点（圆角比例）归零，使圆角矩形变方。
 *  用 Shape.adjustments（PowerPointApi 1.10+）。纯线条无调整点，会跳过并提示。 */
export async function opStraightenLines(): Promise<OpResult> {
  const opName = "拉直形状";
  if (!apiSupported("PowerPointApi", "1.10"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.10+（形状调整点 API）。", api: "1.10" };
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+（选区）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null) return { ok: false, op: opName, detail: "当前 Office 不支持选区 API。", api: "1.7" };
    if (sel.length === 0) return { ok: false, op: opName, detail: "请先选中至少 1 个形状。" };
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/type");
    await ctx.sync();
    let done = 0, skipped = 0;
    for (const s of sel2.items) {
      try {
        if (s.type !== "GeometricShape") { skipped++; continue; }
        const adj = s.adjustments;
        adj.load("count");
        await ctx.sync();
        if (!adj.count || adj.count < 1) { skipped++; continue; }
        const r = adj.get(0);
        await ctx.sync();
        if (typeof r.value !== "number") { skipped++; continue; }
        adj.set(0, 0);
        await ctx.sync();
        done++;
      } catch { skipped++; }
    }
    if (done === 0) return { ok: false, op: opName, detail: `所选形状无曲线/圆角可调（${skipped} 个跳过；纯线条请用 PowerPoint 原生『编辑顶点』）。` };
    return { ok: true, op: opName, detail: `${done} 个形状已拉直${skipped ? `（${skipped} 个无曲线可调跳过）` : ""}`, api: "1.10" };
  });
}

/* ===================== B 组：页操作 ===================== */

/** 删除当前页演讲者备注。需 PowerPointApi 1.7+（notesSlide）。 */
export async function opRemoveNotes(): Promise<OpResult> {
  const opName = "删除备注";
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "请先选中/激活一张幻灯片。" };
    const slide = slides.items[0];
    slide.load("notesSlide");
    await ctx.sync();
    const ns = slide.notesSlide;
    if (!ns) return { ok: false, op: opName, detail: "当前页没有备注可删。" };
    ns.load("textFrame");
    await ctx.sync();
    ns.textFrame.textRange.text = "";
    await ctx.sync();
    return { ok: true, op: opName, detail: "当前页备注已清空。", api: "1.7" };
  });
}

/** 移除当前页动画。⚠️ PowerPoint JavaScript API 目前（≤1.9）仍未公开动画读写接口，
 *  故做能力探测：若宿主暴露 slide.animations 则清空，否则诚实提示用原生方式。 */
export async function opRemoveAnimations(): Promise<OpResult> {
  const opName = "移除动画";
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "请先选中/激活一张幻灯片。" };
    const slide = slides.items[0];
    const anySlide: any = slide;
    if (typeof anySlide.animations === "undefined") {
      return { ok: false, op: opName, detail: "当前 Office.js 暂未开放动画读写接口（PowerPointApi ≤1.9）。请在 PowerPoint 中清除：动画 → 『高级动画』组『动画窗格』→ 全选删除，或把效果设为『无』。" };
    }
    try {
      const anims: any[] = anySlide.animations.items || [];
      for (const a of anims) a.delete?.();
      await ctx.sync();
      return { ok: true, op: opName, detail: "当前页动画已移除。", api: "1.9" };
    } catch (e: any) {
      return { ok: false, op: opName, detail: "动画接口存在但操作失败（" + (e?.message || e) + "）。请改用 PowerPoint 原生动画窗格删除。" };
    }
  });
}

/** 复制当前页到其后。需 PowerPointApi 1.2+（slide.duplicate）。 */
export async function opDuplicateSlide(): Promise<OpResult> {
  const opName = "复制幻灯片";
  if (!apiSupported("PowerPointApi", "1.2"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.2+。", api: "1.2" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "请先选中/激活一张幻灯片。" };
    const slide = slides.items[0];
    slide.duplicate();
    await ctx.sync();
    return { ok: true, op: opName, detail: "当前页已复制一份到其后。", api: "1.2" };
  });
}

/* ===================== C 组：咨询基因 ===================== */

/** 哈维球（Harvey Ball）：在选中的圆形（椭圆）上叠加一个扇形（pie）表示比例 level（0~1）。
 *  从 12 点方向顺时针扫过 level*360°。需 PowerPointApi 1.8+（addGeometricShape）。
 *  ⚠️ pie 的 adjustment 角度语义各宿主略有差异，已用归一化绝对角尝试；若 ¾/满 显示异常，
 *  真机验证后再微调（代码已对 level≥1 降级为直接填充椭圆）。 */
export async function opHarveyBall(level: number): Promise<OpResult> {
  const opName = "哈维球 " + Math.round(level * 100) + "%";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addGeometricShape）。", api: "1.8" };
  if (level <= 0) return { ok: false, op: opName, detail: "空哈维球无需绘制（原圆保持空心即可）。" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const slide = slides.items[0];
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/type,items/autoShapeType,items/name,items/left,items/top,items/width,items/height,items/fill/type,items/fill/foregroundColor");
    await ctx.sync();
    let done = 0, skipped = 0;
    for (const s of sel2.items) {
      const nm: string = s.name || "";
      const isOval = s.autoShapeType === "oval" || s.type === "GeometricShape" && (nm.includes("椭圆") || nm.includes("Oval") || nm.toLowerCase().includes("ellipse"));
      if (!isOval) { skipped++; continue; }
      try {
        if (level >= 1) {
          // 满：直接填充椭圆为品牌色（不画 pie，避免整圆 adj 歧义）
          s.fill.setSolidColor("#0F2B46");
        } else {
          const pie = slide.shapes.addGeometricShape("pie");
          pie.left = s.left; pie.top = s.top; pie.width = s.width; pie.height = s.height;
          let fillColor = "#0F2B46";
          if ((s.fill?.type || "").toLowerCase() === "solid" && s.fill?.foregroundColor) fillColor = s.fill.foregroundColor;
          pie.fill.setSolidColor(fillColor);
          const adj = pie.adjustments;   // PowerPointApi 1.10 Adjustments：get/set(index)
          adj.load("count");
          await ctx.sync();
          if (adj.count >= 2) {
            // 12 点方向 = 270°，归一化绝对角 = 角度/360
            adj.set(0, 270 / 360);
            adj.set(1, (270 + level * 360) / 360);
            await ctx.sync();
          }
        }
        done++;
      } catch {
        // pie 调整点失败时降级：直接填充椭圆
        try { s.fill.setSolidColor("#0F2B46"); done++; } catch { skipped++; }
      }
    }
    await ctx.sync();
    if (done === 0) return { ok: false, op: opName, detail: `请在当前页选中圆形（椭圆）形状（${skipped} 个非圆形已跳过）。` };
    return { ok: true, op: opName, detail: `${done} 个圆形已绘制 ${Math.round(level * 100)}% 哈维球${skipped ? `（${skipped} 个非圆跳过）` : ""}`, api: "1.8" };
  });
}

/** 议程页（目录）：在当前页插入「目录」标题与每张幻灯片一条占位条目。需 1.8+。 */
export async function opAgenda(): Promise<OpResult> {
  const opName = "议程页";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addTextBox）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const slide = slides.items[0];
    const all = ctx.presentation.slides;
    all.load("items");
    await ctx.sync();
    const n = all.items.length;
    const title = slide.shapes.addTextBox("目录  Agenda");
    title.left = 60; title.top = 50; title.width = 600; title.height = 50;
    const tf0 = title.textFrame.textRange.font; tf0.bold = true; tf0.size = 28; tf0.color = "#0F2B46";
    for (let i = 0; i < n; i++) {
      const tb = slide.shapes.addTextBox(`${i + 1}. 议题 ${i + 1}（编辑替换）`);
      tb.left = 80; tb.top = 130 + i * 46; tb.width = 560; tb.height = 36;
      const tf = tb.textFrame.textRange.font; tf.size = 18; tf.color = "#333333";
    }
    await ctx.sync();
    return { ok: true, op: opName, detail: `已插入目录标题 + ${n} 条议题占位（对应演示共 ${n} 页）。`, api: "1.8" };
  });
}

/** 价值链（波特基础活动 + 辅助活动条）：在当前页绘制一套价值链示意图。需 1.8+。 */
export async function opValueChain(): Promise<OpResult> {
  const opName = "价值链";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addGeometricShape/addTextBox）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const slide = slides.items[0];
    const primaries = ["进厂物流", "生产运营", "出厂物流", "营销销售", "售后服务"];
    const supports = ["企业基础设施", "人力资源", "技术开发", "采购"];
    const palette = ["#1B6CA8", "#2E9E97", "#E8A33D", "#C0504D", "#8496A9"];
    const x0 = 60, y = 220, w = 150, h = 70, gap = 14;
    // 辅助活动长条（横贯上方）
    const bar = slide.shapes.addGeometricShape("rectangle");
    bar.left = x0; bar.top = 130; bar.width = primaries.length * (w + gap) - gap; bar.height = 40;
    bar.fill.setSolidColor("#0F2B46");
    const btb = slide.shapes.addTextBox("辅助活动：" + supports.join(" / "));
    btb.left = bar.left; btb.top = bar.top + 10; btb.width = bar.width; btb.height = 24;
    const btf = btb.textFrame.textRange.font; btf.size = 13; btf.color = "#FFFFFF";
    // 基础活动
    primaries.forEach((name, i) => {
      const r = slide.shapes.addGeometricShape("rectangle");
      r.left = x0 + i * (w + gap); r.top = y; r.width = w; r.height = h;
      r.fill.setSolidColor(palette[i % palette.length]);
      const tb = slide.shapes.addTextBox(name);
      tb.left = r.left; tb.top = r.top + 24; tb.width = w; tb.height = 24;
      const tf = tb.textFrame.textRange.font; tf.size = 13; tf.color = "#FFFFFF"; tf.bold = true;
    });
    // 标题
    const t = slide.shapes.addTextBox("价值链（Value Chain）");
    t.left = 60; t.top = 55; t.width = 600; t.height = 40;
    const tfh = t.textFrame.textRange.font; tfh.bold = true; tfh.size = 22; tfh.color = "#0F2B46";
    await ctx.sync();
    return { ok: true, op: opName, detail: `已绘制价值链：5 个基础活动 + 1 条辅助活动（${supports.length} 类）。`, api: "1.8" };
  });
}

/* ===================== 思维模型（面板专属：Ribbon 放不下的功能） ===================== */

/** 2×2 矩阵：双轴四象限占位示意图（象限/轴文案均可后期编辑）。 */
export async function opMatrix2x2(): Promise<OpResult> {
  const opName = "2×2 矩阵";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addGeometricShape/addTextBox）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const slide = slides.items[0];
    const x0 = 240, y0 = 130, w = 180, h = 110, g = 6;
    const quads: Array<[string, string, number, number]> = [
      ["重点投入", "#1B6CA8", 0, 0], ["维持优势", "#2E9E97", 1, 0],
      ["机会观察", "#E8A33D", 0, 1], ["谨慎退出", "#8496A9", 1, 1],
    ];
    for (const [label, color, col, row] of quads) {
      const r = slide.shapes.addGeometricShape("rectangle");
      r.left = x0 + col * (w + g); r.top = y0 + row * (h + g); r.width = w; r.height = h;
      r.fill.setSolidColor(color);
      const tb = slide.shapes.addTextBox(label);
      tb.left = r.left; tb.top = r.top + h / 2 - 14; tb.width = w; tb.height = 28;
      const tf = tb.textFrame.textRange.font; tf.size = 15; tf.bold = true; tf.color = "#FFFFFF";
    }
    const axisY = slide.shapes.addTextBox("↑ 市场吸引力");
    axisY.left = x0 - 140; axisY.top = y0 + h - 10; axisY.width = 130; axisY.height = 24;
    axisY.textFrame.textRange.font.size = 12;
    const axisX = slide.shapes.addTextBox("竞争地位 →");
    axisX.left = x0 + w; axisX.top = y0 + 2 * h + g + 8; axisX.width = 140; axisX.height = 24;
    axisX.textFrame.textRange.font.size = 12;
    const t = slide.shapes.addTextBox("2×2 矩阵");
    t.left = x0; t.top = y0 - 56; t.width = 300; t.height = 40;
    const tfh = t.textFrame.textRange.font; tfh.bold = true; tfh.size = 22; tfh.color = "#0F2B46";
    await ctx.sync();
    return { ok: true, op: opName, detail: "已绘制 2×2 四象限矩阵（轴标签/象限文案可直接编辑）。", api: "1.8" };
  });
}

/** 金字塔图：顶三角 + 3 层梯形，层级文案可编辑。 */
export async function opPyramid(): Promise<OpResult> {
  const opName = "金字塔图";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addGeometricShape/addTextBox）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const slide = slides.items[0];
    const cx = 400, y0 = 120, h = 72, g = 4;
    const layers: Array<[string, number, string]> = [
      ["愿景 / 目标", 140, "#0F2B46"],
      ["战略举措", 250, "#1B6CA8"],
      ["关键任务", 360, "#2E9E97"],
      ["日常执行", 470, "#E8A33D"],
    ];
    let y = y0;
    for (const [label, w, color] of layers) {
      const r = slide.shapes.addGeometricShape(w === 140 ? "triangle" : "trapezoid");
      r.left = cx - w / 2; r.top = y; r.width = w; r.height = h;
      r.fill.setSolidColor(color);
      const tb = slide.shapes.addTextBox(label);
      tb.left = cx - w / 2; tb.top = y + h / 2 - 13; tb.width = w; tb.height = 26;
      const tf = tb.textFrame.textRange.font; tf.size = 14; tf.bold = true; tf.color = "#FFFFFF";
      y += h + g;
    }
    const t = slide.shapes.addTextBox("金字塔（目标分解）");
    t.left = cx - 150; t.top = y0 - 56; t.width = 300; t.height = 40;
    const tfh = t.textFrame.textRange.font; tfh.bold = true; tfh.size = 22; tfh.color = "#0F2B46";
    await ctx.sync();
    return { ok: true, op: opName, detail: "已绘制 4 层金字塔（顶三角 + 3 层梯形），文案可直接编辑。", api: "1.8" };
  });
}

/* ===================== D 组：配色方案管理（localStorage） ===================== */

const SCHEME_KEY = "qslide.schemes.v1";

/** 保存配色方案：抓取当前页所有纯色（非中性）填充色作为方案，写入 localStorage。 */
export async function opSaveScheme(name: string): Promise<OpResult> {
  const opName = "保存配色方案";
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "请先激活一张幻灯片。" };
    const shapes = slides.items[0].shapes;
    shapes.load("items/fill/type,items/fill/foregroundColor");
    await ctx.sync();
    const colors: string[] = [];
    for (const s of shapes.items) {
      const fc: string | undefined = s.fill?.foregroundColor;
      const type: string | undefined = s.fill?.type;
      if (fc && (type || "").toLowerCase() === "solid") {
        const rgb = colorToRgb(fc);
        if (rgb && !isNeutral(rgb)) colors.push(fc.replace("#", "").toUpperCase());
      }
    }
    if (colors.length === 0) return { ok: false, op: opName, detail: "当前页没有可保存的纯色（非中性）填充。" };
    const uniq = Array.from(new Set(colors)).slice(0, 8);
    try {
      const raw = localStorage.getItem(SCHEME_KEY);
      const m = raw ? JSON.parse(raw) : {};
      m[name] = uniq;
      localStorage.setItem(SCHEME_KEY, JSON.stringify(m));
    } catch {
      return { ok: false, op: opName, detail: "本地存储不可用，无法保存方案。" };
    }
    return { ok: true, op: opName, detail: `已保存方案「${name}」（${uniq.length} 色：${uniq.join(" ")}）`, api: "1.7" };
  });
}

/** 列出已保存方案名。 */
export function listSchemes(): string[] {
  try {
    const raw = localStorage.getItem(SCHEME_KEY);
    const m = raw ? JSON.parse(raw) : {};
    return Object.keys(m);
  } catch {
    return [];
  }
}

/** 应用配色方案：把当前页所有纯色（非中性）填充形状 snap 到方案里的最近色。 */
export async function opApplyScheme(name: string): Promise<OpResult> {
  const opName = "应用配色方案";
  let colors: string[] = [];
  try {
    const raw = localStorage.getItem(SCHEME_KEY);
    const m = raw ? JSON.parse(raw) : {};
    colors = m[name] || [];
  } catch { /* ignore */ }
  if (colors.length === 0) return { ok: false, op: opName, detail: `方案「${name}」不存在或为空（先用「保存当前页配色为方案」）。` };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "请先激活一张幻灯片。" };
    const shapes = slides.items[0].shapes;
    shapes.load("items/shapeId,items/fill/type,items/fill/foregroundColor");
    await ctx.sync();
    let changed = 0, skipped = 0;
    for (const s of shapes.items) {
      const fc: string | undefined = s.fill?.foregroundColor;
      const type: string | undefined = s.fill?.type;
      if (!fc || (type || "").toLowerCase() !== "solid") { skipped++; continue; }
      const rgb = colorToRgb(fc);
      if (!rgb || isNeutral(rgb)) { skipped++; continue; }
      let best = colors[0], bestD = Infinity;
      for (const c of colors) {
        const pr = colorToRgb("#" + c)!;
        const d = (pr[0] - rgb[0]) ** 2 + (pr[1] - rgb[1]) ** 2 + (pr[2] - rgb[2]) ** 2;
        if (d < bestD) { bestD = d; best = c; }
      }
      s.fill.setSolidColor("#" + best);
      changed++;
    }
    await ctx.sync();
    return { ok: true, op: opName, detail: `${changed} 个形状已 snap 到方案「${name}」${skipped ? `（${skipped} 跳过）` : ""}`, api: "1.7" };
  });
}

/* ===================== E 元能力：频率埋点 ===================== */

const USAGE_KEY = "qslide.usage.v1";

/** 记录一次操作使用（localStorage 计数）。 */
export function trackUsage(op: string): void {
  try {
    const raw = localStorage.getItem(USAGE_KEY);
    const m: Record<string, number> = raw ? JSON.parse(raw) : {};
    m[op] = (m[op] || 0) + 1;
    localStorage.setItem(USAGE_KEY, JSON.stringify(m));
  } catch { /* localStorage 不可用则忽略 */ }
}

/** 读取使用频率统计（降序）。 */
export function getUsageStats(): { op: string; n: number }[] {
  try {
    const raw = localStorage.getItem(USAGE_KEY);
    const m: Record<string, number> = raw ? JSON.parse(raw) : {};
    return Object.entries(m).map(([op, n]) => ({ op, n })).sort((a, b) => b.n - a.n);
  } catch {
    return [];
  }
}

/* ===================== F 组：文本样式（对照 QPT 字号/行距增减） ===================== */

/** 字号增减：对所选形状文本按 delta(pt) 调整字号，下限 4pt 防止不可读。需 1.5+。 */
export async function opFontSize(delta: number): Promise<OpResult> {
  const opName = delta >= 0 ? "增大字号" : "减小字号";
  if (!apiSupported("PowerPointApi", "1.5"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.5+（文本格式）。", api: "1.5" };
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+（选区）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/textFrame/textRange/font/size");
    await ctx.sync();
    let done = 0, skipped = 0;
    for (const s of sel2.items) {
      try {
        const f = s.textFrame?.textRange?.font;
        if (!f) { skipped++; continue; }
        const cur = typeof f.size === "number" ? f.size : 18;
        f.size = Math.max(4, cur + delta);
        await ctx.sync();
        done++;
      } catch { skipped++; }
    }
    if (done === 0) return { ok: false, op: opName, detail: `所选形状均无文本可调整字号（${skipped} 个跳过）。` };
    return { ok: true, op: opName, detail: `${done} 个形状字号已${delta >= 0 ? "增大" : "减小"} ${Math.abs(delta)}pt${skipped ? `（${skipped} 个无文本跳过）` : ""}`, api: "1.5" };
  });
}

/** 行距增减：对所选形状文本段落行距按 delta(pt) 调整。需 1.5+（paragraphFormat）。 */
export async function opLineSpacing(delta: number): Promise<OpResult> {
  const opName = delta >= 0 ? "增大行距" : "减小行距";
  if (!apiSupported("PowerPointApi", "1.5"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.5+（段落格式）。", api: "1.5" };
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+（选区）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/textFrame/textRange/paragraphFormat/lineSpacing");
    await ctx.sync();
    let done = 0, skipped = 0;
    for (const s of sel2.items) {
      try {
        const pf = s.textFrame?.textRange?.paragraphFormat;
        if (!pf) { skipped++; continue; }
        const cur = typeof pf.lineSpacing === "number" ? pf.lineSpacing : 18;
        pf.lineSpacing = Math.max(8, cur + delta);
        await ctx.sync();
        done++;
      } catch { skipped++; }
    }
    if (done === 0) return { ok: false, op: opName, detail: `所选形状均无段落可调整行距（${skipped} 个跳过）。` };
    return { ok: true, op: opName, detail: `${done} 个形状行距已${delta >= 0 ? "增大" : "减小"} ${Math.abs(delta)}pt${skipped ? `（${skipped} 个跳过）` : ""}`, api: "1.5" };
  });
}

/* ===================== G 组：形状特效 / 格式刷（对照 QPT 移除特效、吸取应用格式） ===================== */

/** 移除形状特效：去阴影/反射/发光/柔化边缘。Office.js 当前仅可靠去阴影；
 *  反射/发光/柔化边缘按宿主 API 暴露情况探测性尝试，未开放则如实说明。 */
export async function opRemoveEffects(): Promise<OpResult> {
  const opName = "移除形状特效";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+（选区）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId");
    await ctx.sync();
    let shadowed = 0, skipped = 0;
    const unsupported: string[] = [];
    for (const s of sel2.items) {
      try {
        let any = false;
        if (s.shadow && typeof s.shadow.visible === "boolean") { s.shadow.visible = false; any = true; shadowed++; }
        // 反射/发光/柔化边缘：仅当属性在该宿主存在才尝试
        for (const prop of ["reflection", "glow", "softEdge"]) {
          try {
            const ef: any = s[prop];
            if (ef && typeof ef.visible === "boolean") { ef.visible = false; any = true; }
          } catch { if (!unsupported.includes(prop)) unsupported.push(prop); }
        }
        if (any) await ctx.sync();
      } catch { skipped++; }
    }
    if (shadowed === 0 && skipped === sel2.items.length)
      return { ok: false, op: opName, detail: "所选形状均无可见特效，或当前 API 不支持读取特效属性。" };
    const extra = unsupported.length
      ? `；反射/发光/柔化边缘当前 Office.js（≤1.9）未开放，请在 PowerPoint『格式→形状效果』清除`
      : "";
    return { ok: true, op: opName, detail: `${shadowed} 个形状已去阴影${skipped ? `（${skipped} 个无特效/跳过）` : ""}${extra}`, api: "1.3" };
  });
}

/** 格式刷：吸取 + 应用。吸取把首个选中形状的填充/线/字体/尺寸存入模块变量；
 *  应用把该格式写到后续选中的所有形状（各项逐属性 try 隔离）。 */
interface PickedFormat {
  fill?: { type: string; color?: string };
  line?: { color?: string; weight?: number };
  font?: { name?: string; size?: number; bold?: boolean; color?: string };
  size?: { width: number; height: number };
}
let pickedFormat: PickedFormat | null = null;

export async function opPickFormat(): Promise<OpResult> {
  const opName = "吸取格式";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+（选区）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/fill/type,items/fill/foregroundColor," +
      "items/line/color,items/line/weight," +
      "items/textFrame/textRange/font/name,items/textFrame/textRange/font/size," +
      "items/textFrame/textRange/font/bold,items/textFrame/textRange/font/color," +
      "items/width,items/height");
    await ctx.sync();
    if (sel2.items.length === 0) return { ok: false, op: opName, detail: "请先选中至少 1 个形状作为格式源。" };
    const s = sel2.items[0];
    pickedFormat = {
      fill: { type: s.fill?.type || "solid", color: s.fill?.foregroundColor },
      line: { color: s.line?.color, weight: s.line?.weight },
      font: {
        name: s.textFrame?.textRange?.font?.name,
        size: s.textFrame?.textRange?.font?.size,
        bold: s.textFrame?.textRange?.font?.bold,
        color: s.textFrame?.textRange?.font?.color,
      },
      size: { width: s.width, height: s.height },
    };
    return { ok: true, op: opName, detail: `已吸取（填充 ${pickedFormat.fill?.color || "-"} / 字体 ${(pickedFormat.font?.name || "-")} ${pickedFormat.font?.size || "-"}pt）`, api: "1.7" };
  });
}

export async function opApplyFormat(): Promise<OpResult> {
  const opName = "应用格式";
  if (!pickedFormat) return { ok: false, op: opName, detail: "请先点『吸取格式』选中一个源形状。" };
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+（选区）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId");
    await ctx.sync();
    if (sel2.items.length === 0) return { ok: false, op: opName, detail: "请选中要套用格式的目标形状。" };
    let applied = 0;
    const pf = pickedFormat!;
    for (const s of sel2.items) {
      try {
        if (pf.fill?.color && (pf.fill.type || "").toLowerCase() === "solid") {
          try { s.fill.setSolidColor(pf.fill.color); } catch { /* 无填充面 */ }
        }
        if (typeof pf.line?.color === "string") {
          try { s.line.color = pf.line.color; } catch { /* 无边框 */ }
        }
        const f = s.textFrame?.textRange?.font;
        if (f && pf.font) {
          if (typeof pf.font.name === "string") try { f.name = pf.font.name; } catch { /* */ }
          if (typeof pf.font.size === "number") try { f.size = pf.font.size; } catch { /* */ }
          if (typeof pf.font.bold === "boolean") try { f.bold = pf.font.bold; } catch { /* */ }
          if (typeof pf.font.color === "string") try { f.color = pf.font.color; } catch { /* */ }
        }
        await ctx.sync();
        applied++;
      } catch { /* 单个形状失败跳过 */ }
    }
    if (applied === 0) return { ok: false, op: opName, detail: "所选形状均无法套用（多为图片/无文本框）。" };
    return { ok: true, op: opName, detail: `${applied} 个形状已套用吸取的格式（填充/线/字体）`, api: "1.7" };
  });
}

/* ===================== H 组：内边距归零（对照 QPT 删除内部边距） ===================== */

/** 内边距归零：把形状文本框四边内部边距设为 0（消灭 PPT 默认 0.1″，咨询稿对齐关键）。
 *  allSlides=true 时对全文档所有幻灯片生效。需 1.5+。 */
export async function opZeroMargins(allSlides: boolean): Promise<OpResult> {
  const opName = allSlides ? "内边距归零（全文档）" : "内边距归零";
  if (!apiSupported("PowerPointApi", "1.5"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.5+（文本框边距）。", api: "1.5" };
  return await PowerPoint.run(async (ctx: any) => {
    let slideScope: any;
    if (allSlides) {
      slideScope = ctx.presentation.slides;
    } else {
      slideScope = ctx.presentation.getSelectedSlides?.();
      if (!slideScope) {
        // 退化：取第一页
        const all = ctx.presentation.slides;
        all.load("items");
        await ctx.sync();
        slideScope = { items: all.items[0] ? [all.items[0]] : [] };
      }
    }
    slideScope.load("items");
    await ctx.sync();
    let cells = 0, skipped = 0;
    for (const slide of slideScope.items) {
      const shapes = slide.shapes;
      shapes.load("items/shapeId,items/textFrame/marginTop,items/textFrame/marginBottom,items/textFrame/marginLeft,items/textFrame/marginRight");
      await ctx.sync();
      for (const s of shapes.items) {
        try {
          const tf = s.textFrame;
          if (!tf) { skipped++; continue; }
          tf.marginTop = 0; tf.marginBottom = 0; tf.marginLeft = 0; tf.marginRight = 0;
          await ctx.sync();
          cells++;
        } catch { skipped++; }
      }
    }
    if (cells === 0) return { ok: false, op: opName, detail: `没有可重置边距的文本框（${skipped} 个无文本框形状跳过）。` };
    return { ok: true, op: opName, detail: `${cells} 个文本框内边距已归零${allSlides ? "（全文档）" : ""}${skipped ? `（${skipped} 个跳过）` : ""}`, api: "1.5" };
  });
}

/* ===================== I 组：自绘表格（对照 QPT 表格） ===================== */

/** 自绘表格：在当前页画 rows×cols 单元格（矩形+居中文本），第 0 行为品牌蓝表头，
 *  其余行浅蓝/白斑马纹。保持短路径：行列数由入参指定（面板用输入框）。需 1.8+。 */
export async function opInsertTable(rows: number, cols: number): Promise<OpResult> {
  const opName = "插入表格";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addGeometricShape）。", api: "1.8" };
  if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1 || rows > 20 || cols > 20)
    return { ok: false, op: opName, detail: "行列数需为 1~20 的整数。" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const slide = slides.items[0];
    const x0 = 80, y0 = 140, cw = 150, ch = 44, gap = 2;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const rect = slide.shapes.addGeometricShape("rectangle");
        rect.left = x0 + c * (cw + gap);
        rect.top = y0 + r * (ch + gap);
        rect.width = cw; rect.height = ch;
        if (r === 0) rect.fill.setSolidColor("#0F2B46");
        else rect.fill.setSolidColor(r % 2 === 1 ? "#EAF1F7" : "#FFFFFF");
        const tb = slide.shapes.addTextBox(r === 0 ? `表头 ${c + 1}` : `单元格 ${r + 1}-${c + 1}`);
        tb.left = rect.left + 4; tb.top = rect.top + (ch - 22) / 2; tb.width = cw - 8; tb.height = 22;
        const f = tb.textFrame.textRange.font;
        f.size = 12; f.bold = r === 0; f.color = r === 0 ? "#FFFFFF" : "#333333";
        try { tb.textFrame.marginTop = 0; tb.textFrame.marginBottom = 0; tb.textFrame.marginLeft = 0; tb.textFrame.marginRight = 0; } catch { /* */ }
      }
    }
    await ctx.sync();
    return { ok: true, op: opName, detail: `已插入 ${rows}×${cols} 表格（品牌蓝表头 + 斑马纹，编辑替换占位文本）`, api: "1.8" };
  });
}

/* ===================== J 组：灰度优化（对照 QPT 灰度优化） ===================== */

/** 灰度优化：把全文档所有纯色填充与文本颜色转灰（0.3R+0.59G+0.11B），打印/投影友好。 */
export async function opGrayscale(): Promise<OpResult> {
  const opName = "灰度优化";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.slides;
    slides.load("items");
    await ctx.sync();
    let fills = 0, texts = 0, skipped = 0;
    const toGray = (hex: string): string | null => {
      const rgb = colorToRgb(hex);
      if (!rgb) return null;
      const g = Math.round(0.3 * rgb[0] + 0.59 * rgb[1] + 0.11 * rgb[2]);
      const h = g.toString(16).padStart(2, "0");
      return "#" + h + h + h;
    };
    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items/shapeId,items/fill/type,items/fill/foregroundColor," +
        "items/textFrame/textRange/text,items/textFrame/textRange/font/color");
      await ctx.sync();
      for (const s of shapes.items) {
        try {
          const ft = s.fill?.type || "";
          if (ft.toLowerCase() === "solid" && s.fill?.foregroundColor) {
            const g = toGray(s.fill.foregroundColor);
            if (g) { s.fill.setSolidColor(g); fills++; }
          }
        } catch { skipped++; }
        try {
          const t = s.textFrame?.textRange?.text;
          if (t && s.textFrame?.textRange?.font?.color) {
            const g = toGray(s.textFrame.textRange.font.color);
            if (g) { s.textFrame.textRange.font.color = g; texts++; }
          }
        } catch { /* 文本颜色失败忽略 */ }
      }
      try { await ctx.sync(); } catch { /* 部分失败继续 */ }
    }
    if (fills === 0 && texts === 0) return { ok: false, op: opName, detail: "未发现可转灰的纯色填充或文本颜色。" };
    return { ok: true, op: opName, detail: `${fills} 个填充 + ${texts} 处文本色已转灰（全文档）`, api: "1.7" };
  });
}

/* ===================== K 组：对齐到页 / 叠放次序 / 旋转归零（P0） ===================== */

/** 探测幻灯片页面尺寸（pt）。Office.js ≤1.9 未正式暴露 PageSetup.SlideWidth，
 *  先尝试读 presentation.slideWidth/Height（部分宿主预览版支持），失败回退 16:9 960×540。 */
async function getSlideSize(ctx: any): Promise<{ w: number; h: number }> {
  try {
    const p: any = ctx.presentation;
    p.load("slideWidth,slideHeight");
    await ctx.sync();
    const w = p.slideWidth, h = p.slideHeight;
    if (typeof w === "number" && typeof h === "number" && w > 100 && h > 100) return { w, h };
  } catch { /* 读不到则用假设值 */ }
  return { w: 960, h: 540 };
}

/** 居中于页：把所选形状水平/垂直/双方向居中到页面中心。需 1.7+。 */
export async function opCenterOnSlide(mode: "h" | "v" | "both"): Promise<OpResult> {
  const opName = mode === "h" ? "水平居中于页" : mode === "v" ? "垂直居中于页" : "居中于页";
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+（选区）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null) return { ok: false, op: opName, detail: "当前 Office 不支持选区 API（需 1.7+）。", api: "1.7" };
    if (sel.length === 0) return { ok: false, op: opName, detail: "请先选中至少 1 个形状。" };
    const { w, h } = await getSlideSize(ctx);
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/left,items/top,items/width,items/height");
    await ctx.sync();
    let done = 0;
    for (const s of sel2.items) {
      if (mode === "h" || mode === "both") s.left = (w - s.width) / 2;
      if (mode === "v" || mode === "both") s.top = (h - s.height) / 2;
      done++;
    }
    await ctx.sync();
    const sizeNote = `（基于页面 ${Math.round(w)}×${Math.round(h)}pt${w === 960 && h === 540 ? "；若页面为 4:3 请在反馈里告知" : ""}）`;
    return { ok: true, op: opName, detail: `${done} 个形状已${mode === "h" ? "水平" : mode === "v" ? "垂直" : "双向"}居中于页${sizeNote}`, api: "1.7" };
  });
}

/** 旋转归零：把所选形状 rotation 设为 0。需 1.3+。 */
export async function opResetRotation(): Promise<OpResult> {
  const opName = "旋转归零";
  if (!apiSupported("PowerPointApi", "1.3"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.3+。", api: "1.3" };
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+（选区）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null) return { ok: false, op: opName, detail: "需 1.7+。", api: "1.7" };
    if (sel.length === 0) return { ok: false, op: opName, detail: "请先选中至少 1 个形状。" };
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/rotation");
    await ctx.sync();
    let done = 0, skipped = 0;
    for (const s of sel2.items) {
      try { s.rotation = 0; done++; } catch { skipped++; }
    }
    await ctx.sync();
    if (done === 0) return { ok: false, op: opName, detail: `所选形状均无旋转属性（${skipped} 个跳过）。` };
    return { ok: true, op: opName, detail: `${done} 个形状旋转已归零${skipped ? `（${skipped} 个跳过）` : ""}`, api: "1.3" };
  });
}

/** 叠放次序：front/back/forward/backward。
 *  ⚠️ 1) Shape.zOrder 要求 ShapeZOrder 枚举值（BringToFront/BringForward/SendToBack/SendBackward）。
 *  ⚠️ 2) Mac 宿主 isSetSupported 对 1.13 谎报 false（实际能执行），故不设版本门禁，
 *     改为运行时 typeof 特性探测。 */
export async function opZOrder(pos: "front" | "back" | "forward" | "backward"): Promise<OpResult> {
  const map = { front: "置顶", back: "置底", forward: "上移一层", backward: "下移一层" };
  const cmdMap = { front: "BringToFront", back: "SendToBack", forward: "BringForward", backward: "SendBackward" };
  const opName = map[pos];
  if (!apiSupported("PowerPointApi", "1.7"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+（选区）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null) return { ok: false, op: opName, detail: "需 1.7+。", api: "1.7" };
    if (sel.length === 0) return { ok: false, op: opName, detail: "请先选中至少 1 个形状。" };
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId");
    await ctx.sync();
    // 特性探测：宿主 requirements 谎报不可信，直接看代理对象上有没有方法
    if (typeof sel2.items[0]?.zOrder !== "function")
      return { ok: false, op: opName, detail: "当前 PowerPoint 运行时缺少 Shape.zOrder 方法，请更新 Microsoft 365。", api: "1.13" };
    let done = 0;
    let lastErr = "";
    for (const s of sel2.items) {
      try { s.zOrder(cmdMap[pos]); done++; } catch (e: any) { lastErr = e?.message || e?.code || String(e); }
    }
    await ctx.sync();
    if (done === 0)
      return { ok: false, op: opName, detail: `调整叠放次序失败${lastErr ? `：${lastErr}` : "（多为特殊对象）"}`, api: "1.13" };
    return { ok: true, op: opName, detail: `${done} 个形状已${opName}`, api: "1.13" };
  });
}

/* ===================== L 组：文本拆形状（P0：Lines2Shapes） ===================== */

/** 文本拆形状：选中 1 个文本框 → 按段落（\n）拆成多个独立文本框（逐行），纵向堆叠。
 *  原文本框删除。用于「逐行动画 / 矩阵式排版」。需 1.8+。 */
export async function opLines2Shapes(): Promise<OpResult> {
  const opName = "文本拆形状";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addTextBox）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null) return { ok: false, op: opName, detail: "需 1.7+。", api: "1.7" };
    if (sel.length !== 1) return { ok: false, op: opName, detail: "请只选中 1 个文本框。" };
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/left,items/top,items/width,items/height," +
      "items/textFrame/textRange/text," +
      "items/textFrame/textRange/font/name,items/textFrame/textRange/font/size," +
      "items/textFrame/textRange/font/bold,items/textFrame/textRange/font/color");
    await ctx.sync();
    const src = sel2.items[0];
    const t: string | undefined = src.textFrame?.textRange?.text;
    if (!t || !t.trim()) return { ok: false, op: opName, detail: "选中的形状没有文本内容。" };
    const rawLines = t.split(/\r?\n/).map((x) => x.trimEnd());
    const textLines = rawLines.filter((x) => x.length > 0);
    if (textLines.length < 2) return { ok: false, op: opName, detail: "文本框只有单行，无需拆分（多行文本才会拆成多个形状）。" };
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    const slide = slides.items[0];
    if (!slide) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const font = src.textFrame.textRange.font;
    const snap = { name: font.name, size: font.size, bold: font.bold, color: font.color };
    const x = src.left ?? 80, w = src.width ?? 300;
    const lineH = Math.max((typeof snap.size === "number" ? snap.size : 18) * 1.25, 20);
    let y = src.top ?? 120;
    let made = 0;
    for (const line of textLines) {
      try {
        const tb = slide.shapes.addTextBox(line);
        tb.left = x; tb.top = y; tb.width = w; tb.height = lineH;
        const f = tb.textFrame.textRange.font;
        if (typeof snap.name === "string") f.name = snap.name;
        if (typeof snap.size === "number") f.size = snap.size;
        if (typeof snap.bold === "boolean") f.bold = snap.bold;
        if (typeof snap.color === "string") f.color = snap.color;
        made++;
      } catch { /* 单个失败 */ }
      y += lineH + 4;
    }
    try { src.delete(); } catch { /* 删除失败忽略 */ }
    await ctx.sync();
    return { ok: true, op: opName, detail: `已将 1 个文本框拆为 ${made} 个独立形状（逐行，格式保留）`, api: "1.8" };
  });
}

/* ===================== M 组：排版体检 + 一键合规（P0：QPT 灵魂能力） ===================== */

export interface AuditIssue { level: "warn" | "info"; rule: string; count: number; detail: string; }
export interface AuditReport {
  slideCount: number; scope: string;
  fontKinds: string[]; colorKinds: number;
  issues: AuditIssue[]; marginShapes: number; offslide: number; noActionTitle: number;
  checkedShapes: number; page: number;
}
const AUDIT_KEY = "qslide.audit.v1";

/** 排版体检：扫描全文档/本页，检查字体数、颜色数、默认内边距、超版心、缺 Action Title。
 *  报告写入 localStorage，由面板「排版体检」区渲染；返回摘要。 */
export async function opAudit(scope: "slide" | "all"): Promise<OpResult> {
  const opName = scope === "all" ? "排版体检（全文档）" : "排版体检（本页）";
  return await PowerPoint.run(async (ctx: any) => {
    const slides = scope === "all" ? ctx.presentation.slides : ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取幻灯片。" };
    slides.load("items");
    await ctx.sync();
    const slideItems = slides.items;
    const { w, h } = await getSlideSize(ctx);
    const fonts = new Set<string>();
    const colors = new Set<string>();
    let marginShapes = 0, offslide = 0, noActionTitle = 0, checkedShapes = 0;
    const issues: AuditIssue[] = [];
    for (const slide of slideItems) {
      const shapes = slide.shapes;
      shapes.load("items/shapeId,items/type,items/left,items/top,items/width,items/height," +
        "items/fill/type,items/fill/foregroundColor," +
        "items/textFrame/textRange/text,items/textFrame/textRange/font/name," +
        "items/textFrame/marginLeft,items/textFrame/marginTop");
      await ctx.sync();
      let hasTitle = false;
      for (const s of shapes.items) {
        checkedShapes++;
        const nm: string | undefined = s.textFrame?.textRange?.font?.name;
        if (nm) fonts.add(nm);
        const fc: string | undefined = s.fill?.foregroundColor;
        const ft: string | undefined = s.fill?.type;
        if (fc && (ft || "").toLowerCase() === "solid") {
          const rgb = colorToRgb(fc);
          if (rgb && !isNeutral(rgb)) colors.add(fc.toUpperCase());
        }
        if (typeof s.textFrame?.marginLeft === "number" && s.textFrame.marginLeft > 0.5) marginShapes++;
        if (typeof s.textFrame?.marginTop === "number" && s.textFrame.marginTop > 0.5) marginShapes++;
        if (typeof s.left === "number" && typeof s.top === "number" && typeof s.width === "number" && typeof s.height === "number") {
          if (s.left < -1 || s.top < -1 || s.left + s.width > w + 1 || s.top + s.height > h + 1) offslide++;
        }
        if (typeof s.top === "number" && s.top < 80 && s.textFrame?.textRange?.text?.trim()) hasTitle = true;
      }
      if (!hasTitle) noActionTitle++;
    }
    const fontKinds = Array.from(fonts);
    const colorKinds = colors.size;
    if (fontKinds.length > 2) issues.push({ level: "warn", rule: "字体数超标", count: fontKinds.length, detail: `用了 ${fontKinds.length} 种字体：${fontKinds.join("、")}（建议 ≤2）` });
    if (colorKinds > 6) issues.push({ level: "warn", rule: "颜色数超标", count: colorKinds, detail: `用了 ${colorKinds} 种非中性色（建议 ≤6）` });
    if (marginShapes > 0) issues.push({ level: "warn", rule: "默认内边距", count: marginShapes, detail: `${marginShapes} 个文本框含默认内边距（建议归零）` });
    if (offslide > 0) issues.push({ level: "warn", rule: "超出版心", count: offslide, detail: `${offslide} 个形状超出页面（${Math.round(w)}×${Math.round(h)}pt）` });
    if (noActionTitle > 0) issues.push({ level: "info", rule: "缺 Action Title", count: noActionTitle, detail: `${noActionTitle} 页缺少顶部标题（top<80 含文本）` });
    if (issues.length === 0) issues.push({ level: "info", rule: "合规", count: 0, detail: "未发现明显问题（字体数/颜色数/内边距/版心/标题均达标）" });
    const report: AuditReport = { slideCount: slideItems.length, scope, fontKinds, colorKinds, issues, marginShapes, offslide, noActionTitle, checkedShapes, page: w === 960 && h === 540 ? 0 : 1 };
    try { localStorage.setItem(AUDIT_KEY, JSON.stringify(report)); } catch { /* ignore */ }
    const warns = issues.filter((i) => i.level === "warn").length;
    const summary = `扫描 ${slideItems.length} 页 / ${checkedShapes} 形状：字体 ${fontKinds.length} 种、颜色 ${colorKinds} 种、内边距 ${marginShapes}、超版心 ${offslide}、缺标题 ${noActionTitle}；发现 ${warns} 项需关注。`;
    return { ok: true, op: opName, detail: summary + "（详见下方体检报告）", api: "1.7" };
  });
}

export function getAuditReport(): AuditReport | null {
  try {
    const raw = localStorage.getItem(AUDIT_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch { return null; }
}

/** 一键合规：安全项——全文档内边距归零 + 全文档统一色系。返回汇总。 */
export async function opRepair(): Promise<OpResult> {
  const opName = "一键合规";
  if (!apiSupported("PowerPointApi", "1.7")) return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+。", api: "1.7" };
  const parts: string[] = [];
  const r1 = await opZeroMargins(true);
  parts.push("内边距：" + r1.detail);
  return await PowerPoint.run(async (ctx: any) => {
    const all = ctx.presentation.slides;
    all.load("items");
    await ctx.sync();
    let changed = 0, page = 0;
    for (const slide of all.items) {
      page++;
      const shapes = slide.shapes;
      shapes.load("items/shapeId,items/fill/type,items/fill/foregroundColor");
      await ctx.sync();
      for (const s of shapes.items) {
        const fc: string | undefined = s.fill?.foregroundColor;
        const type: string | undefined = s.fill?.type;
        if (!fc || (type || "").toLowerCase() !== "solid") continue;
        const rgb = colorToRgb(fc);
        if (!rgb || isNeutral(rgb)) continue;
        const target = "#" + nearestPalette(fc);
        if (target.toLowerCase() === fc.toLowerCase()) continue;
        try { s.fill.setSolidColor(target); changed++; } catch { /* 单形状失败 */ }
      }
    }
    await ctx.sync();
    parts.push(`统一色系：全 ${page} 页共 snap ${changed} 个形状`);
    return { ok: true, op: opName, detail: parts.join("；"), api: "1.7" };
  });
}

/* ===================== N 组：P1 生成/标注类 ===================== */

export const PICKED_COLOR_KEY = "qslide.pickedColor.v1";

/** 取色器：吸取首个选中形状的纯色填充（或文字色）存入本地，供面板「我的色板」点击填充。需 1.7+。 */
export async function opPickColor(): Promise<OpResult> {
  const opName = "取色器";
  if (!apiSupported("PowerPointApi", "1.7")) return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/fill/type,items/fill/foregroundColor,items/textFrame/textRange/font/color");
    await ctx.sync();
    if (sel2.items.length === 0) return { ok: false, op: opName, detail: "请选中 1 个形状（取其填充色或文字色）。" };
    const s = sel2.items[0];
    let hex: string | null = null;
    const ft = s.fill?.type || "";
    if (ft.toLowerCase() === "solid" && s.fill?.foregroundColor) hex = s.fill.foregroundColor.replace("#", "").toUpperCase();
    if (!hex && s.textFrame?.textRange?.font?.color) hex = s.textFrame.textRange.font.color.replace("#", "").toUpperCase();
    if (!hex) return { ok: false, op: opName, detail: "所选形状无纯色填充也无文字色可吸（图片/渐变不支持）。" };
    try { localStorage.setItem(PICKED_COLOR_KEY, hex); } catch { /* ignore */ }
    return { ok: true, op: opName, detail: `已吸取 #${hex}（点击「我的色板」色块即可填充选中形状）`, api: "1.7" };
  });
}

export function getPickedColor(): string | null {
  try { return localStorage.getItem(PICKED_COLOR_KEY); } catch { return null; }
}

/** Roadmap 页眉：当前页顶部插入咨询风页眉（左侧色条 + 标题 + 右侧阶段标签）。需 1.8+。 */
export async function opRoadmapHeader(): Promise<OpResult> {
  const opName = "Roadmap 页眉";
  if (!apiSupported("PowerPointApi", "1.8")) return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addGeometricShape）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const slide = slides.items[0];
    const bar = slide.shapes.addGeometricShape("rectangle");
    bar.left = 0; bar.top = 0; bar.width = 220; bar.height = 14; bar.fill.setSolidColor("#0F2B46");
    const tb = slide.shapes.addTextBox("Roadmap 阶段占位（编辑替换）");
    tb.left = 36; tb.top = 22; tb.width = 560; tb.height = 34;
    const tf = tb.textFrame.textRange.font; tf.bold = true; tf.size = 20; tf.color = "#0F2B46";
    const tag = slide.shapes.addTextBox("阶段 1 / N");
    tag.left = 760; tag.top = 24; tag.width = 160; tag.height = 28;
    const ttf = tag.textFrame.textRange.font; ttf.size = 12; ttf.color = "#8496A9";
    await ctx.sync();
    return { ok: true, op: opName, detail: "已在页顶插入 Roadmap 页眉（左侧色条 + 标题 + 右侧阶段标签，可编辑）", api: "1.8" };
  });
}

/** 便签：当前页插入黄色圆角矩形便签（占位文本）。需 1.8+。 */
export async function opStickyNote(): Promise<OpResult> {
  const opName = "便签";
  if (!apiSupported("PowerPointApi", "1.8")) return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    if (slides.items.length === 0) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const slide = slides.items[0];
    const note = slide.shapes.addGeometricShape("roundRect");
    note.left = 560; note.top = 300; note.width = 200; note.height = 150;
    note.fill.setSolidColor("#FFF3B0");
    try { note.line.format.visible = false; } catch { /* */ }
    const tb = slide.shapes.addTextBox("便签：记一笔（编辑替换）");
    tb.left = note.left + 8; tb.top = note.top + 10; tb.width = 184; tb.height = 130;
    const tf = tb.textFrame.textRange.font; tf.size = 13; tf.color = "#5b4a00";
    try { tb.textFrame.marginTop = 0; tb.textFrame.marginBottom = 0; tb.textFrame.marginLeft = 0; tb.textFrame.marginRight = 0; } catch { /* */ }
    await ctx.sync();
    return { ok: true, op: opName, detail: "已插入黄色便签（圆角矩形 + 文本框）", api: "1.8" };
  });
}

/** 单位标注：为所选形状在右侧标注宽×高（pt）。需 1.8+。 */
export async function opDimensionLabel(): Promise<OpResult> {
  const opName = "单位标注";
  if (!apiSupported("PowerPointApi", "1.8")) return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel = await getSelectedShapesIn(ctx);
    if (sel === null) return { ok: false, op: opName, detail: "需 1.7+。", api: "1.7" };
    if (sel.length === 0) return { ok: false, op: opName, detail: "请先选中至少 1 个形状。" };
    const slides = ctx.presentation.getSelectedSlides?.();
    if (!slides) return { ok: false, op: opName, detail: "宿主不支持获取当前页。" };
    slides.load("items");
    await ctx.sync();
    const slide = slides.items[0];
    if (!slide) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/left,items/top,items/width,items/height");
    await ctx.sync();
    let made = 0;
    for (const s of sel2.items) {
      try {
        const w = Math.round(s.width), h = Math.round(s.height);
        const label = slide.shapes.addTextBox(`${w}×${h}pt`);
        label.left = s.left + s.width + 6; label.top = s.top + s.height / 2 - 10;
        label.width = 100; label.height = 20;
        const tf = label.textFrame.textRange.font; tf.size = 10; tf.color = "#8496A9";
        try { label.textFrame.marginTop = 0; label.textFrame.marginBottom = 0; label.textFrame.marginLeft = 0; label.textFrame.marginRight = 0; } catch { /* */ }
        made++;
      } catch { /* 单形状失败 */ }
    }
    await ctx.sync();
    return { ok: true, op: opName, detail: `已为 ${made} 个形状标注尺寸（宽×高 pt，位于右侧）`, api: "1.8" };
  });
}

/** 自动换行切换：对所选形状 textFrame.wordWrap 取反。需 1.5+。 */
export async function opToggleWrap(): Promise<OpResult> {
  const opName = "自动换行切换";
  if (!apiSupported("PowerPointApi", "1.5")) return { ok: false, op: opName, detail: "需 PowerPointApi 1.5+。", api: "1.5" };
  if (!apiSupported("PowerPointApi", "1.7")) return { ok: false, op: opName, detail: "需 PowerPointApi 1.7+（选区）。", api: "1.7" };
  return await PowerPoint.run(async (ctx: any) => {
    const sel2 = ctx.presentation.getSelectedShapes();
    sel2.load("items/shapeId,items/textFrame/wordWrap");
    await ctx.sync();
    if (sel2.items.length === 0) return { ok: false, op: opName, detail: "请先选中至少 1 个形状。" };
    let on = 0, off = 0;
    for (const s of sel2.items) {
      try {
        const cur = s.textFrame?.wordWrap;
        const next = !(cur === true);
        s.textFrame.wordWrap = next;
        if (next) on++; else off++;
      } catch { /* 单形状失败 */ }
    }
    await ctx.sync();
    return { ok: true, op: opName, detail: `${on} 个开 / ${off} 个关自动换行（取反切换）`, api: "1.5" };
  });
}

/* ===================== 插入组：形状 / 直线 / 图片 ===================== */

/** 取当前页（首个选中页）的代理。需调用方随后 load+sync。 */
async function firstSlideIn(ctx: any): Promise<any | null> {
  const slides = ctx.presentation.getSelectedSlides?.();
  if (!slides) return null;
  slides.load("items");
  await ctx.sync();
  return slides.items.length > 0 ? slides.items[0] : null;
}

/** 插入常见形状：置于当前页中心，白底深蓝描边（咨询风默认样式）。需 1.8+。 */
export async function opInsertShape(ast: string, label: string): Promise<OpResult> {
  const opName = "插入" + label;
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addGeometricShape）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slide = await firstSlideIn(ctx);
    if (!slide) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const { w, h } = await getSlideSize(ctx);
    const SW = 160, SH = 110;
    const s = slide.shapes.addGeometricShape(ast);
    s.left = Math.round((w - SW) / 2); s.top = Math.round((h - SH) / 2);
    s.width = SW; s.height = SH;
    try { s.fill.setSolidColor("#FFFFFF"); } catch { /* 部分形状无独立填充 */ }
    try { s.line.color = "#0F2B46"; s.line.weight = 1.5; } catch { /* */ }
    await ctx.sync();
    return { ok: true, op: opName, detail: "已插入到当前页中心（白底深蓝描边）。", api: "1.8" };
  });
}

/** 插入直线：当前页中部水平直线，深蓝 2pt。需 1.8+（addLine）。 */
export async function opInsertLine(): Promise<OpResult> {
  const opName = "插入直线";
  if (!apiSupported("PowerPointApi", "1.8"))
    return { ok: false, op: opName, detail: "需 PowerPointApi 1.8+（addLine）。", api: "1.8" };
  return await PowerPoint.run(async (ctx: any) => {
    const slide = await firstSlideIn(ctx);
    if (!slide) return { ok: false, op: opName, detail: "未获取到当前页。" };
    const { w, h } = await getSlideSize(ctx);
    let ln: any;
    try {
      ln = slide.shapes.addLine({ left: Math.round(w / 2 - 120), top: Math.round(h / 2), width: 240, height: 0 });
    } catch {
      return { ok: false, op: opName, detail: "当前 Office 不支持 addLine，请更新 Microsoft 365。", api: "1.8" };
    }
    try { ln.line.color = "#0F2B46"; ln.line.weight = 2; } catch { /* */ }
    await ctx.sync();
    return { ok: true, op: opName, detail: "已插入 240pt 水平直线到当前页中部。", api: "1.8" };
  });
}

/** 插入图片（base64）：置于当前页中心，默认最长边适配页面 70%。
 *  主路径 addImageFromBase64（1.8+，可精确定位缩放）；部分 Mac 宿主运行时虽上报支持
 *  1.8/1.13，但 ShapeCollection 上实际没有该方法（TypeError: is not a function），
 *  故先 typeof 探测，缺方法直接转兜底 setSelectedDataAsync（CoercionType.Image，1.1+ 全平台）。
 *  不同宿主对 base64 是否带 data: 前缀要求不一：双格式各试一次（独立 run，避免失败代理污染下一轮）。 */
export async function opInsertImageFromBase64(b64: string, wPt?: number, hPt?: number): Promise<OpResult> {
  const opName = "插入图片";
  if (!b64) return { ok: false, op: opName, detail: "未收到图片数据。" };
  const stripped = b64.replace(/^data:[^,]*,/, "");
  const variants = b64.startsWith("data:")
    ? [b64, stripped]
    : ["data:image/png;base64," + b64, b64];

  // 统一计算目标尺寸（最长边适配页面 70%）；getSlideSize 自带 960×540 假设值兜底
  let iw = wPt && wPt > 20 ? wPt : 320;
  let ih = hPt && hPt > 20 ? hPt : 220;
  const size = await PowerPoint.run(async (ctx: any) => getSlideSize(ctx)).catch(() => ({ w: 960, h: 540 }));
  const k = Math.min(1, size.w * 0.7 / iw, size.h * 0.7 / ih);
  iw = Math.round(iw * k); ih = Math.round(ih * k);

  // ── 主路径：addImageFromBase64（1.8+）──
  if (apiSupported("PowerPointApi", "1.8")) {
    let hostMissing = false; // 宿主缺方法 → 转兜底
    for (const variant of variants) {
      try {
        return await PowerPoint.run(async (ctx: any) => {
          const slide = await firstSlideIn(ctx);
          if (!slide) return { ok: false, op: opName, detail: "未获取到当前页。" };
          const sc: any = slide.shapes;
          if (typeof sc.addImageFromBase64 !== "function") throw new Error("__hostMissing_addImageFromBase64__");
          sc.addImageFromBase64(variant, {
            left: Math.round((size.w - iw) / 2), top: Math.round((size.h - ih) / 2), width: iw, height: ih,
          });
          await ctx.sync();
          return { ok: true, op: opName, detail: `已插入图片（${iw}×${ih}pt）到当前页中心。`, api: "1.8" };
        });
      } catch (e: any) {
        const m = e?.message || e?.code || String(e);
        if (m.includes("__hostMissing_addImageFromBase64__")) { hostMissing = true; break; }
        // API 在但该格式被拒 → 记录后试下一格式
      }
    }
  }

  // ── 兜底：setSelectedDataAsync（CoercionType.Image，1.1+，插到页面中心）──
  // 注意：若当前页有选中形状，宿主可能替换选区——兜底路径仅在主路径不可用时启用。
  return await new Promise<OpResult>((resolve) => {
    const tryInsert = (idx: number) => {
      try {
        Office.context.document.setSelectedDataAsync(variants[idx], {
          coercionType: (Office as any).CoercionType.Image,
          imageWidth: iw,
          imageHeight: ih,
        } as any, (res: any) => {
          if (res?.status === (Office as any).AsyncResultStatus.Succeeded) {
            resolve({ ok: true, op: opName, detail: `已插入图片（${iw}×${ih}pt，页面中心）。`, api: "1.1" });
          } else if (idx + 1 < variants.length) {
            tryInsert(idx + 1);
          } else {
            resolve({ ok: false, op: opName, detail: `插入图片失败：${res?.error?.message || "宿主拒绝写入"}` });
          }
        });
      } catch (e: any) {
        if (idx + 1 < variants.length) tryInsert(idx + 1);
        else resolve({ ok: false, op: opName, detail: `插入图片失败：${e?.message || String(e)}` });
      }
    };
    tryInsert(0);
  });
}
