/**
 * Ribbon 功能文件：ExecuteFunction 按钮的入口。
 * 反馈策略（v0.7.1）：Ribbon 点击 → 直接作用于图元，永不弹窗。
 * 结果（成功+失败）写入 localStorage，由任务窗格「反馈与提示」区实时展示——
 * 面板未打开时先暂存，打开后回放最近记录。
 */
import {
  opSameSize, opAlignTouch, opTouchEdge, opRemoveText, opDistribute, opActionTitle, opUnifyPalette,
  opSubjectTitle, opFootnote, opMergeText, opCollapseSpaces, opAdjustRounding, opStraightenLines,
  opRemoveOutline, opUnifyFont, opCopyToAllSlides, opInsertTextbox, opRemoveNotes, opRemoveAnimations,
  opDuplicateSlide, opHarveyBall, opAgenda, opValueChain,
  opFontSize, opLineSpacing, opRemoveEffects, opPickFormat, opApplyFormat,
  opZeroMargins, opInsertTable, opGrayscale, type OpResult,
  opCenterOnSlide, opResetRotation, opZOrder,
  opInsertShape, opInsertLine, opInsertImageFromBase64,
} from "../core/ops";

/** 与 taskpane/main.ts 约定的通知通道 key 与消息结构 */
export const NOTIFY_KEY = "qslide.notify.v1";
type NotifyMsg = { t: number; ok: boolean; op: string; detail: string; api?: string };

declare global {
  interface Window {
    qslideSameWidth: () => Promise<void>;
    qslideSameHeight: () => Promise<void>;
    qslideTouchLeft: () => Promise<void>;
    qslideTouchUp: () => Promise<void>;
    qslideRemoveText: () => Promise<void>;
    qslideDistributeH: () => Promise<void>;
    qslideDistributeV: () => Promise<void>;
    qslideActionTitle: () => Promise<void>;
    qslideUnify: () => Promise<void>;
    qslideEdgeUp: () => Promise<void>;
    qslideEdgeDown: () => Promise<void>;
    qslideEdgeLeft: () => Promise<void>;
    qslideEdgeRight: () => Promise<void>;
    qslideSubjectTitle: () => Promise<void>;
    qslideFootnote: () => Promise<void>;
    qslideMergeText: () => Promise<void>;
    qslideCollapseSpaces: () => Promise<void>;
    qslideRoundUp: () => Promise<void>;
    qslideRoundDown: () => Promise<void>;
    qslideStraighten: () => Promise<void>;
    qslideRemoveOutline: () => Promise<void>;
    qslideUnifyFont: () => Promise<void>;
    qslideInsertTextbox: () => Promise<void>;
    qslideCopyToAll: () => Promise<void>;
    qslideRemoveNotes: () => Promise<void>;
    qslideRemoveAnimations: () => Promise<void>;
    qslideDuplicateSlide: () => Promise<void>;
    qslideHarvey: () => Promise<void>;
    qslideAgenda: () => Promise<void>;
    qslideValueChain: () => Promise<void>;
    qslideFontUp: () => Promise<void>;
    qslideFontDown: () => Promise<void>;
    qslideLineUp: () => Promise<void>;
    qslideLineDown: () => Promise<void>;
    qslideRemoveEffects: () => Promise<void>;
    qslidePickFormat: () => Promise<void>;
    qslideApplyFormat: () => Promise<void>;
    qslideZeroMargins: () => Promise<void>;
    qslideZeroMarginsAll: () => Promise<void>;
    qslideTable: () => Promise<void>;
    qslideGrayscale: () => Promise<void>;
    qslideCenterH: () => Promise<void>;
    qslideCenterV: () => Promise<void>;
    qslideCenterBoth: () => Promise<void>;
    qslideRotationZero: () => Promise<void>;
    qslideZFront: () => Promise<void>;
    qslideZBack: () => Promise<void>;
    qslideZForward: () => Promise<void>;
    qslideZBackward: () => Promise<void>;
    qslideShapeRect: () => Promise<void>;
    qslideShapeRoundRect: () => Promise<void>;
    qslideShapeEllipse: () => Promise<void>;
    qslideShapeTriangle: () => Promise<void>;
    qslideShapeDiamond: () => Promise<void>;
    qslideShapePentagon: () => Promise<void>;
    qslideShapeHexagon: () => Promise<void>;
    qslideShapeArrowRight: () => Promise<void>;
    qslideShapeStar5: () => Promise<void>;
    qslideShapeHeart: () => Promise<void>;
    qslideShapeCloud: () => Promise<void>;
    qslideShapeLine: () => Promise<void>;
    qslideInsertPicture: () => Promise<void>;
    qslideOpenShapePicker: () => Promise<void>;
  }
}

function report(r: OpResult): void {
  // 永不弹窗：写入 localStorage，任务窗格监听展示（图元变化本身即第一反馈）
  try {
    const raw = localStorage.getItem(NOTIFY_KEY);
    const list: NotifyMsg[] = raw ? JSON.parse(raw) : [];
    list.push({ t: Date.now(), ok: r.ok, op: r.op, detail: r.detail, api: r.api });
    localStorage.setItem(NOTIFY_KEY, JSON.stringify(list.slice(-30)));
  } catch {
    /* 存储不可用时静默，不影响操作本身 */
  }
}

/** Ribbon 函数命令包装：Office 要求每个函数命令最终调用 event.completed()，
 *  否则 Mac 宿主会弹「无 UI 外接程序命令时，请确保调用 event.completed()」错误框。
 *  Office 会把 Event 对象作为最后一个实参传入；这里统一兜底调用。 */
type CmdFn = (...args: any[]) => any;
const withEvent = (fn: CmdFn): CmdFn => async (...args: any[]) => {
  try {
    await fn();
  } finally {
    const ev: any = args[args.length - 1];
    if (ev && typeof ev.completed === "function") {
      try { ev.completed(); } catch { /* 宿主已回收 */ }
    }
  }
};

window.qslideSameWidth = withEvent(withEvent(async () => report(await opSameSize("width"))));
window.qslideSameHeight = withEvent(withEvent(async () => report(await opSameSize("height"))));
window.qslideTouchLeft = withEvent(withEvent(async () => report(await opAlignTouch("h"))));
window.qslideTouchUp = withEvent(withEvent(async () => report(await opAlignTouch("v"))));
window.qslideRemoveText = withEvent(withEvent(async () => report(await opRemoveText())));
window.qslideDistributeH = withEvent(async () => report(await opDistribute("h")));
window.qslideDistributeV = withEvent(async () => report(await opDistribute("v")));
window.qslideActionTitle = withEvent(async () => report(await opActionTitle()));
window.qslideUnify = withEvent(async () => report(await opUnifyPalette()));

window.qslideEdgeUp = withEvent(async () => report(await opTouchEdge("up")));
window.qslideEdgeDown = withEvent(async () => report(await opTouchEdge("down")));
window.qslideEdgeLeft = withEvent(async () => report(await opTouchEdge("left")));
window.qslideEdgeRight = withEvent(async () => report(await opTouchEdge("right")));
window.qslideSubjectTitle = withEvent(async () => report(await opSubjectTitle()));
window.qslideFootnote = withEvent(async () => report(await opFootnote()));
window.qslideMergeText = withEvent(async () => report(await opMergeText()));
window.qslideCollapseSpaces = withEvent(async () => report(await opCollapseSpaces()));
window.qslideRoundUp = withEvent(async () => report(await opAdjustRounding(0.1)));
window.qslideRoundDown = withEvent(async () => report(await opAdjustRounding(-0.1)));
window.qslideStraighten = withEvent(async () => report(await opStraightenLines()));
window.qslideRemoveOutline = withEvent(async () => report(await opRemoveOutline()));
window.qslideUnifyFont = withEvent(async () => report(await opUnifyFont()));
window.qslideInsertTextbox = withEvent(async () => report(await opInsertTextbox()));
window.qslideCopyToAll = withEvent(async () => report(await opCopyToAllSlides()));
window.qslideRemoveNotes = withEvent(async () => report(await opRemoveNotes()));
window.qslideRemoveAnimations = withEvent(async () => report(await opRemoveAnimations()));
window.qslideDuplicateSlide = withEvent(async () => report(await opDuplicateSlide()));
window.qslideHarvey = withEvent(async () => report(await opHarveyBall(0.5)));
window.qslideAgenda = withEvent(async () => report(await opAgenda()));
window.qslideValueChain = withEvent(async () => report(await opValueChain()));
window.qslideFontUp = withEvent(async () => report(await opFontSize(2)));
window.qslideFontDown = withEvent(async () => report(await opFontSize(-2)));
window.qslideLineUp = withEvent(async () => report(await opLineSpacing(3)));
window.qslideLineDown = withEvent(async () => report(await opLineSpacing(-3)));
window.qslideRemoveEffects = withEvent(async () => report(await opRemoveEffects()));
window.qslidePickFormat = withEvent(async () => report(await opPickFormat()));
window.qslideApplyFormat = withEvent(async () => report(await opApplyFormat()));
window.qslideZeroMargins = withEvent(async () => report(await opZeroMargins(false)));
window.qslideZeroMarginsAll = withEvent(async () => report(await opZeroMargins(true)));
window.qslideTable = withEvent(async () => report(await opInsertTable(3, 4)));
window.qslideGrayscale = withEvent(async () => report(await opGrayscale()));
window.qslideCenterH = withEvent(async () => report(await opCenterOnSlide("h")));
window.qslideCenterV = withEvent(async () => report(await opCenterOnSlide("v")));
window.qslideCenterBoth = withEvent(async () => report(await opCenterOnSlide("both")));
window.qslideRotationZero = withEvent(async () => report(await opResetRotation()));
window.qslideZFront = withEvent(async () => report(await opZOrder("front")));
window.qslideZBack = withEvent(async () => report(await opZOrder("back")));
window.qslideZForward = withEvent(async () => report(await opZOrder("forward")));
window.qslideZBackward = withEvent(async () => report(await opZOrder("backward")));

Office.onReady(() => {
  console.log("QSlide commands ready");
});

/* ===================== 插入组：形状 / 直线 / 图片 ===================== */

window.qslideShapeRect = withEvent(async () => report(await opInsertShape("rectangle", "矩形")));
window.qslideShapeRoundRect = withEvent(async () => report(await opInsertShape("roundedRectangle", "圆角矩形")));
window.qslideShapeEllipse = withEvent(async () => report(await opInsertShape("ellipse", "椭圆")));
window.qslideShapeTriangle = withEvent(async () => report(await opInsertShape("triangle", "三角形")));
window.qslideShapeDiamond = withEvent(async () => report(await opInsertShape("diamond", "菱形")));
window.qslideShapePentagon = withEvent(async () => report(await opInsertShape("regularPentagon", "五边形")));
window.qslideShapeHexagon = withEvent(async () => report(await opInsertShape("hexagon", "六边形")));
window.qslideShapeArrowRight = withEvent(async () => report(await opInsertShape("rightArrow", "箭头")));
window.qslideShapeStar5 = withEvent(async () => report(await opInsertShape("star5", "五角星")));
window.qslideShapeHeart = withEvent(async () => report(await opInsertShape("heart", "心形")));
window.qslideShapeCloud = withEvent(async () => report(await opInsertShape("cloud", "云形")));
window.qslideShapeLine = withEvent(async () => report(await opInsertLine()));

/** 插入图片：Ribbon 点击 → 打开本地图片选择小窗 → 选中后 base64 回传 → 插入当前页中心。
 *  选择文件本身需要系统文件对话框（PPT 原生「插入图片」同样如此），除此外不产生任何提示弹窗。 */
window.qslideInsertPicture = withEvent(async () => {
  const opName = "插入图片";
  try {
    const ui: any = (Office as any).context?.ui;
    if (!ui?.displayDialogAsync) {
      report({ ok: false, op: opName, detail: "宿主不支持文件选择窗口（Office.context.ui），可从「开始→插入→图片」操作。" });
      return;
    }
    const dlg: any = await new Promise((resolve, reject) => {
      ui.displayDialogAsync(
        window.location.origin + "/picker.html",
        { width: 32, height: 24, displayInIframe: true },
        (r: any) => {
          if (r.status === (Office as any).AsyncResultStatus.Succeeded) resolve(r.value);
          else reject(new Error(r.error?.message || "无法打开图片选择窗口"));
        }
      );
    });
    const msg: string = await new Promise((resolve) => {
      let done = false;
      const finish = (v: string) => { if (!done) { done = true; resolve(v); } };
      dlg.addEventHandler((Office as any).EventType.DialogMessageReceived, (arg: any) => finish(String(arg.message || "")));
      dlg.addEventHandler((Office as any).EventType.DialogEventReceived, () => finish("__closed__"));
      setTimeout(() => finish("__closed__"), 180000); // 3 分钟未选择视为取消
    });
    try { dlg.close(); } catch { /* 已关闭 */ }
    if (!msg || msg === "__closed__") {
      report({ ok: false, op: opName, detail: "已取消（未选择图片）。" });
      return;
    }
    const payload = JSON.parse(msg);
    report(await opInsertImageFromBase64(payload.b64, payload.w, payload.h));
  } catch (e: any) {
    report({ ok: false, op: opName, detail: e?.message || String(e) });
  }
});

/** 插入图形：Ribbon「插入图形」→ 打开图形选择小窗 → 选中后插入当前页中心。
 *  （PowerPoint 不支持 Ribbon 下拉菜单控件，故用对话框式画廊替代下拉。） */
window.qslideOpenShapePicker = withEvent(async () => {
  const opName = "插入图形";
  try {
    const ui: any = (Office as any).context?.ui;
    if (!ui?.displayDialogAsync) {
      report({ ok: false, op: opName, detail: "宿主不支持选择窗口（Office.context.ui）。" });
      return;
    }
    const dlg: any = await new Promise((resolve, reject) => {
      ui.displayDialogAsync(
        window.location.origin + "/shape-picker.html",
        { width: 40, height: 46, displayInIframe: true },
        (r: any) => {
          if (r.status === (Office as any).AsyncResultStatus.Succeeded) resolve(r.value);
          else reject(new Error(r.error?.message || "无法打开图形选择窗口"));
        }
      );
    });
    const msg: string = await new Promise((resolve) => {
      let done = false;
      const finish = (v: string) => { if (!done) { done = true; resolve(v); } };
      dlg.addEventHandler((Office as any).EventType.DialogMessageReceived, (arg: any) => finish(String(arg.message || "")));
      dlg.addEventHandler((Office as any).EventType.DialogEventReceived, () => finish("__closed__"));
      setTimeout(() => finish("__closed__"), 180000);
    });
    try { dlg.close(); } catch { /* 已关闭 */ }
    if (!msg || msg === "__closed__") return; // 取消，不提示
    const sel = JSON.parse(msg);
    if (sel.ast === "__line__") report(await opInsertLine());
    else report(await opInsertShape(sel.ast, sel.label));
  } catch (e: any) {
    report({ ok: false, op: opName, detail: e?.message || String(e) });
  }
});
