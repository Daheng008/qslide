import {
  opSameSize, opAlignTouch, opTouchEdge, opRemoveText, opDistribute, opSwap, opSelectSimilar,
  opActionTitle, opSubjectTitle, opFootnote, opMergeText, opCollapseSpaces, opApplyFill,
  opUnifyPalette, opRemoveOutline, opInsertTextbox, opUnifyFont, opCopyToAllSlides,
  opAdjustRounding, opStraightenLines, opRemoveNotes, opRemoveAnimations, opDuplicateSlide,
  opHarveyBall, opAgenda, opValueChain, opSaveScheme, opApplyScheme, listSchemes, getUsageStats,
  PALETTE, probeCapabilities, setSimilarMode, getSimilarMode, trackUsage, type OpResult,
  opFontSize, opLineSpacing, opRemoveEffects, opPickFormat, opApplyFormat,
  opZeroMargins, opInsertTable, opGrayscale, opMatrix2x2, opPyramid,
  opCenterOnSlide, opResetRotation, opZOrder, opLines2Shapes, opAudit, opRepair,
  opPickColor, opRoadmapHeader, opStickyNote, opDimensionLabel, opToggleWrap,
  getAuditReport, getPickedColor,
} from "../core/ops";

/** 与 commands.ts 约定的 Ribbon 反馈通道 */
const NOTIFY_KEY = "qslide.notify.v1";
type NotifyMsg = { t: number; ok: boolean; op: string; detail: string; api?: string };

const log = document.getElementById("log")!;

function appendLog(r: OpResult): void {
  const empty = log.querySelector(".empty");
  if (empty) empty.remove();
  const div = document.createElement("div");
  div.className = r.ok ? "ok" : "warn";
  div.textContent = `${r.ok ? "✅" : "⚠️"} ${r.op} — ${r.detail}${r.api ? "（API " + r.api + "）" : ""}`;
  log.prepend(div);
}

let lastNotifyTs = 0;

/** 展示来自 Ribbon 的反馈（按时间戳去重） */
function showNotify(m: NotifyMsg | null): void {
  if (!m || typeof m.t !== "number" || m.t <= lastNotifyTs) return;
  lastNotifyTs = m.t;
  appendLog({ ok: m.ok, op: m.op, detail: m.detail, api: m.api });
}

function pollNotify(): void {
  try {
    const raw = localStorage.getItem(NOTIFY_KEY);
    if (!raw) return;
    const list: NotifyMsg[] = JSON.parse(raw);
    // 只取最近 3 条判断新消息（时间戳去重，避免整表重放）
    for (const m of list.slice(-3)) showNotify(m);
  } catch { /* ignore */ }
}

const OPS: Record<string, () => Promise<OpResult>> = {
  sameWidth: () => opSameSize("width"),
  sameHeight: () => opSameSize("height"),
  touchLeft: () => opAlignTouch("h"),
  touchUp: () => opAlignTouch("v"),
  edgeUp: () => opTouchEdge("up"),
  edgeDown: () => opTouchEdge("down"),
  edgeLeft: () => opTouchEdge("left"),
  edgeRight: () => opTouchEdge("right"),
  distributeH: () => opDistribute("h"),
  distributeV: () => opDistribute("v"),
  swap: () => opSwap(),
  selectSimilar: () => opSelectSimilar(),
  actionTitle: () => opActionTitle(),
  subjectTitle: () => opSubjectTitle(),
  footnote: () => opFootnote(),
  mergeText: () => opMergeText(),
  collapseSpaces: () => opCollapseSpaces(),
  removeText: () => opRemoveText(),
  unifyPalette: () => opUnifyPalette(),
  removeOutline: () => opRemoveOutline(),
  insertTextbox: () => opInsertTextbox(),
  unifyFont: () => opUnifyFont(),
  copyToAll: () => opCopyToAllSlides(),
  roundUp: () => opAdjustRounding(0.1),
  roundDown: () => opAdjustRounding(-0.1),
  straighten: () => opStraightenLines(),
  removeNotes: () => opRemoveNotes(),
  removeAnimations: () => opRemoveAnimations(),
  duplicateSlide: () => opDuplicateSlide(),
  harveyEmpty: () => opHarveyBall(0),
  harveyQuarter: () => opHarveyBall(0.25),
  harveyHalf: () => opHarveyBall(0.5),
  harveyThreeQuarter: () => opHarveyBall(0.75),
  harveyFull: () => opHarveyBall(1),
  agenda: () => opAgenda(),
  valueChain: () => opValueChain(),
  matrix2x2: () => opMatrix2x2(),
  pyramid: () => opPyramid(),
  fontUp: () => opFontSize(2),
  fontDown: () => opFontSize(-2),
  lineUp: () => opLineSpacing(3),
  lineDown: () => opLineSpacing(-3),
  removeEffects: () => opRemoveEffects(),
  pickFormat: () => opPickFormat(),
  applyFormat: () => opApplyFormat(),
  zeroMargins: () => opZeroMargins(false),
  zeroMarginsAll: () => opZeroMargins(true),
  centerH: () => opCenterOnSlide("h"),
  centerV: () => opCenterOnSlide("v"),
  centerBoth: () => opCenterOnSlide("both"),
  resetRotation: () => opResetRotation(),
  zFront: () => opZOrder("front"),
  zBack: () => opZOrder("back"),
  zForward: () => opZOrder("forward"),
  zBackward: () => opZOrder("backward"),
  lines2Shapes: () => opLines2Shapes(),
  auditSlide: async () => { const r = await opAudit("slide"); renderAuditReport(); return r; },
  auditAll: async () => { const r = await opAudit("all"); renderAuditReport(); return r; },
  repair: () => opRepair(),
  pickColor: async () => { const r = await opPickColor(); renderMyColors(); return r; },
  roadmapHeader: () => opRoadmapHeader(),
  sticky: () => opStickyNote(),
  dimLabel: () => opDimensionLabel(),
  toggleWrap: () => opToggleWrap(),
  insertTable: () => {
    const r = parseInt((document.getElementById("tblRows") as HTMLInputElement)?.value || "3", 10);
    const c = parseInt((document.getElementById("tblCols") as HTMLInputElement)?.value || "4", 10);
    return opInsertTable(r, c);
  },
  grayscale: () => opGrayscale(),
  saveScheme: () => {
    const n = (document.getElementById("schemeName") as HTMLInputElement)?.value.trim();
    if (!n) return Promise.resolve({ ok: false, op: "保存配色方案", detail: "请先在输入框填写方案名。" });
    return opSaveScheme(n).then((r) => { refreshSchemes(); return r; });
  },
  applyScheme: () => {
    const sel = document.getElementById("schemeSelect") as HTMLSelectElement;
    const n = sel?.value;
    if (!n) return Promise.resolve({ ok: false, op: "应用配色方案", detail: "请先在下拉选择已保存方案。" });
    return opApplyScheme(n);
  },
  refreshSchemes: async () => { refreshSchemes(); return { ok: true, op: "刷新方案", detail: "已列出 " + listSchemes().length + " 个方案。" }; },
  toggleSimilar: async () => {
    setSimilarMode(!getSimilarMode());
    return {
      ok: true,
      op: "目标切换",
      detail: getSimilarMode()
        ? "同类模式已开启：填充/删除文本将作用于当前页所有与首个选中形状同类的形状"
        : "已回到选区模式：操作仅作用于选中形状",
    };
  },
};

/** 刷新目标模式按钮文案（同类模式可能被「选择相似形状」自动开启） */
function refreshModeBtn(): void {
  const b = document.querySelector<HTMLButtonElement>("button[data-op='toggleSimilar']");
  if (b) b.textContent = getSimilarMode() ? "目标：同类形状 ✔（点击切回选区）" : "目标：选区";
}

/** 刷新已保存配色方案下拉 */
function refreshSchemes(): void {
  const sel = document.getElementById("schemeSelect") as HTMLSelectElement | null;
  if (!sel) return;
  const names = listSchemes();
  sel.innerHTML = names.length
    ? names.map((n) => `<option value="${n}">${n}</option>`).join("")
    : `<option value="">（无已存方案）</option>`;
}

/** 刷新使用频率表（E 元能力） */
function renderUsage(): void {
  const table = document.getElementById("usageTable");
  if (!table) return;
  const stats = getUsageStats();
  table.innerHTML = stats.length
    ? stats.map((s) => `<tr><td>${s.op}</td><td>${s.n}</td></tr>`).join("")
    : `<tr><td colspan="2" style="color:#888">暂无记录（用一次任意操作即统计）</td></tr>`;
}

/** 渲染排版体检报告（来自 localStorage 的 AuditReport） */
function renderAuditReport(): void {
  const el = document.getElementById("auditReport");
  if (!el) return;
  const r = getAuditReport();
  if (!r) { el.innerHTML = `<div class="hint">点击上方「扫描本页/全文档」生成规范体检报告。</div>`; return; }
  const lvlText: Record<string, string> = { warn: "需关注", info: "提示", ok: "合规" };
  const rows = r.issues.map((i) => `<li class="lvl-${i.level}">[${lvlText[i.level] || i.level}] ${i.rule}（${i.count}）：${i.detail}</li>`).join("");
  el.innerHTML = `<div class="audit-sum">${r.scope === "all" ? "全文档" : "本页"} ${r.slideCount} 页 / ${r.checkedShapes} 形状 · 字体 ${r.fontKinds.length} 种 · 非中性色 ${r.colorKinds} 种 · 内边距 ${r.marginShapes} · 超版心 ${r.offslide} · 缺标题 ${r.noActionTitle}</div><ul class="audit-list">${rows}</ul>`;
}

/** 渲染「我的色板」（取色器吸取的色，可点选填充选中形状） */
function renderMyColors(): void {
  const el = document.getElementById("myColors");
  if (!el) return;
  const hex = getPickedColor();
  if (!hex) { el.innerHTML = `<div class="hint" style="margin:0">点「取色器」吸取形状颜色后，这里出现可点击的色块。</div>`; return; }
  el.innerHTML = `<button class="swatch" data-fill="${hex}" style="background:#${hex}" title="#${hex} 点击填充选中形状"></button>`;
  el.querySelector("button")!.addEventListener("click", (e) => { void runOp(e.target as HTMLButtonElement); });
}

async function runOp(btn: HTMLButtonElement): Promise<void> {
  // 色板按钮：data-fill 存 hex；排版按钮：data-op 存操作名
  const fill = btn.dataset.fill;
  const opKey = btn.dataset.op;
  const op = fill ? () => opApplyFill(fill) : OPS[opKey!];
  if (!op) return;
  // 频率埋点（色板填充记为 fill）
  trackUsage(fill ? "填充#" + fill : opKey!);
  btn.disabled = true;
  try {
    appendLog(await op());
  } catch (e: any) {
    appendLog({ ok: false, op: btn.textContent || opKey!, detail: e?.message || String(e) });
  } finally {
    btn.disabled = false;
    refreshModeBtn();
    renderUsage();
  }
}

document.querySelectorAll<HTMLButtonElement>("button[data-op]").forEach((btn) => {
  btn.addEventListener("click", () => void runOp(btn));
});

Office.onReady(() => {
  // 1) 色板条
  const palette = document.getElementById("palette")!;
  for (const c of PALETTE) {
    const b = document.createElement("button");
    b.dataset.fill = c.hex;
    b.title = c.name + " #" + c.hex;
    b.style.background = "#" + c.hex;
    b.addEventListener("click", () => void runOp(b));
    palette.appendChild(b);
  }
  // 2) 能力探测表
  const table = document.getElementById("capTable")!;
  for (const cap of probeCapabilities()) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>${cap.name}</td><td>PowerPointApi ${cap.min}+</td>` +
      `<td class="${cap.ok ? "yes" : "no"}">${cap.ok ? "✔ 支持" : "✘ 不支持"}</td>`;
    table.appendChild(tr);
  }
  // 3) 配色方案下拉
  refreshSchemes();
  // 4) 使用频率表
  renderUsage();
  // 5) Ribbon 反馈桥：监听 localStorage 新消息（storage 事件 + 轮询兜底），
  //    打开面板时回放最近记录（最多 8 条）
  try {
    const raw = localStorage.getItem(NOTIFY_KEY);
    if (raw) {
      const list: NotifyMsg[] = JSON.parse(raw);
      for (const m of list.slice(-8)) showNotify(m);
    }
  } catch { /* ignore */ }
  window.addEventListener("storage", (e) => {
    if (e.key === NOTIFY_KEY) {
      try {
        const list: NotifyMsg[] = e.newValue ? JSON.parse(e.newValue) : [];
        for (const m of list.slice(-3)) showNotify(m);
      } catch { /* ignore */ }
    }
  });
  setInterval(pollNotify, 600);
});
