/**
 * Workbook surface of `@wae/open-excel`.
 *
 * The session, workbook summary, format registry and the native-backed create,
 * open, save, edit, inspect, export and SVG render operations live here. When
 * the addon is absent every engine call rejects with
 * `capability/not-implemented`; the capability matrix always states which
 * formats and checks this build really supports.
 */

export { OpenExcel } from './open-excel.js';
export { GuiError, closeGui, defaultGuiContentDir, isGuiOpen, openGui, pumpEvents, runGui } from './gui.js';
export type { GuiCloseResult, GuiOpenResult, GuiStatus, OpenGuiOptions } from './gui.js';
export { NotImplementedError, OpenExcelError, errorCodes, isOpenExcelError, renderErrorCopy, setErrorLocale } from './errors.js';
export type { OpenExcelErrorCode, OpenExcelErrorDetails, OpenExcelErrorOptions, OpenExcelErrorParams } from './errors.js';
export { defaultIntentFor, formatFromPath, formatInfo, formatRegistry, workbookFormats } from './formats.js';
export type {
    FormatFamily,
    FormatMarkup,
    FormatOpenKind,
    WorkbookFormat,
    WorkbookFormatInfo,
    WorkbookIntent,
} from './formats.js';
export type * from './types.js';
