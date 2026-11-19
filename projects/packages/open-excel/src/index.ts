/**
 * Declared workbook surface of `@wae/open-excel`.
 *
 * Everything here is a declared contract: the session, the workbook summary and
 * the format registry exist, while every capability rejects with
 * `capability/not-implemented` until the shared engine lands. Domain rules,
 * formula evaluation, pivot aggregation, rendering and spreadsheet exchange
 * belong upstream and are never re-implemented in this package.
 */

export { OpenExcel } from './open-excel.js';
export { GuiError, closeGui, defaultGuiContentDir, isGuiOpen, openGui, pumpEvents, runGui } from './gui.js';
export type { GuiCloseResult, GuiOpenResult, GuiStatus, OpenGuiOptions } from './gui.js';
export { NotImplementedError, OpenExcelError, errorCodes, isOpenExcelError } from './errors.js';
export type { OpenExcelErrorCode, OpenExcelErrorDetails, OpenExcelErrorOptions } from './errors.js';
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
