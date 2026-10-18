// 最小化 Office.js 环境声明（运行时经 CDN 加载真身）。
// Phase 0 采用宽松 any 策略：先探测能力，MVP 锁定后再逐步收紧类型。
declare const Office: any;
declare const PowerPoint: any;

interface OfficeExtensionErrorInfo {
  code?: string;
  message?: string;
}
