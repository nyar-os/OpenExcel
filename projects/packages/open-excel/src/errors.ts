/**
 * Stable error surface of `@wae/open-excel`.
 *
 * Codes are contract, not prose: callers compare them, localize around them and
 * never parse the message. A failure therefore keeps its code even when the
 * text is translated.
 */

/** Every stable code the workbook surface may report. */
export const errorCodes = [
    'session/disposed',
    'session/argument-invalid',
    'session/budget-exceeded',
    'workbook/not-found',
    'document/revision-conflict',
    'io/read',
    'io/write',
    'io/exists',
    'io/failure',
    // Frozen published code of the 0.0.1 addon; the structured contract emits
    // `io/exists` instead. Kept so the legacy adapter still decodes old bindings.
    'io/target-exists',
    'object/not-found',
    'object/unsupported',
    'format/unsupported',
    'format/invalid',
    'capability/not-implemented',
    'binding/failure',
] as const;

/** Union of {@link errorCodes}. */
export type OpenExcelErrorCode = (typeof errorCodes)[number];

/** Optional locators a failure can carry so callers can point at the object. */
export interface OpenExcelErrorDetails {
    readonly revision?: number;
    readonly sheetId?: string;
    /** A1-style cell address such as `B4`. */
    readonly cell?: string;
    /** Workbook format involved in the failure, without the leading dot. */
    readonly format?: string;
    /** Capability involved in the failure, e.g. `workbook.export`. */
    readonly feature?: string;
}

/** Machine-readable per-code parameters delivered by the binding contract. */
export type OpenExcelErrorParams = Readonly<Record<string, unknown>>;

/** Options accepted by {@link OpenExcelError}. */
export interface OpenExcelErrorOptions {
    readonly recoverable?: boolean;
    readonly details?: OpenExcelErrorDetails;
    readonly params?: OpenExcelErrorParams;
    readonly cause?: unknown;
}

/** Every failure thrown or rejected by this package. */
export class OpenExcelError extends Error {
    /** Stable, comparable code. */
    readonly code: OpenExcelErrorCode;
    /** Whether retrying with different input or state can succeed. */
    readonly recoverable: boolean;
    /** Optional locators (`revision`, `sheetId`, `cell`, `format`, `feature`). */
    readonly details: OpenExcelErrorDetails;
    /** Typed parameters of the code, e.g. `{ kind, id }` for `object/not-found`. */
    readonly params: OpenExcelErrorParams;

    constructor(code: OpenExcelErrorCode, message: string, options: OpenExcelErrorOptions = {}) {
        super(message, options.cause === undefined ? undefined : { cause: options.cause });
        this.name = 'OpenExcelError';
        this.code = code;
        this.recoverable = options.recoverable ?? false;
        this.details = Object.freeze({ ...options.details });
        this.params = Object.freeze({ ...options.params });
    }
}

/**
 * Reported by declared methods whose backing capability is not wired yet.
 *
 * The scaffold uses one code for every unimplemented entry point instead of
 * inventing per-method placeholders; each method drops this rejection as it is
 * implemented, and the code disappears with the last one.
 */
export class NotImplementedError extends OpenExcelError {
    constructor(feature: string, message = `\`${feature}\` is not implemented yet`) {
        super('capability/not-implemented', message, { recoverable: false, details: { feature } });
        this.name = 'NotImplementedError';
    }
}

/** Narrows an unknown thrown value to {@link OpenExcelError}. */
export function isOpenExcelError(value: unknown): value is OpenExcelError {
    return value instanceof OpenExcelError;
}

/**
 * Product copy for each stable code, rendered from its typed `params`.
 *
 * The binding's own text is only a fallback for codes this catalog does not
 * carry yet; the machine-readable `code` and `params` never depend on locale.
 */
const messageCatalog: Readonly<Record<string, Readonly<Record<string, string>>>> = {
    en: {
        'session/disposed': 'the OpenExcel session has been disposed; create a new instance',
        'session/argument-invalid': 'the request was rejected: {reason}',
        'session/budget-exceeded': 'the operation exceeded the session budget',
        'workbook/not-found': 'no workbook is open for this handle',
        'document/revision-conflict': 'the workbook changed at revision {actual}; the edit expected {expected}',
        'io/read': 'could not read `{path}` ({kind})',
        'io/write': 'could not publish `{path}`',
        'io/exists': '`{path}` already exists; pass overwrite to replace it',
        'io/failure': 'the file operation failed ({kind})',
        'object/not-found': 'the `{kind}` `{id}` was not found',
        'object/unsupported': 'the `{operation}` `{name}` is not supported by this build',
        'format/unsupported': '`{format}` is not an offered format',
        'format/invalid': 'the `{format}` input is invalid: {reason}',
        'capability/not-implemented': 'the `{feature}` capability is not implemented yet',
        'binding/failure': 'the native binding reported a failure without a stable code',
    },
    'zh-CN': {
        'session/disposed': 'OpenExcel 会话已释放；请创建新实例',
        'session/argument-invalid': '请求被拒绝：{reason}',
        'session/budget-exceeded': '操作超出会话预算',
        'workbook/not-found': '该句柄没有打开的工作簿',
        'document/revision-conflict': '工作簿已变更到修订 {actual}；本次编辑期望 {expected}',
        'io/read': '无法读取 `{path}`（{kind}）',
        'io/write': '无法发布 `{path}`',
        'io/exists': '`{path}` 已存在；传入 overwrite 以替换',
        'io/failure': '文件操作失败（{kind}）',
        'object/not-found': '未找到 `{kind}` `{id}`',
        'object/unsupported': '此构建不支持 `{operation}` `{name}`',
        'format/unsupported': '`{format}` 不是可提供的格式',
        'format/invalid': '`{format}` 输入无效：{reason}',
        'capability/not-implemented': '`{feature}` 能力尚未实现',
        'binding/failure': '原生绑定返回了没有稳定码的失败',
    },
};

let currentLocale = 'en';

/** Selects the locale used to render error copy; `zh*` maps to `zh-CN`. */
export function setErrorLocale(tag: string): void {
    currentLocale = tag.toLowerCase().startsWith('zh') ? 'zh-CN' : 'en';
}

/**
 * Renders the product copy for a code and its params.
 *
 * Unknown codes and missing translations fall back to the binding's own text,
 * so a replaced or removed native message never changes which branch a caller
 * takes: that branch is decided by `code` and `params`.
 */
export function renderErrorCopy(code: string, params: OpenExcelErrorParams, fallback: string): string {
    const template = messageCatalog[currentLocale]?.[code];
    if (template === undefined) {
        return fallback;
    }
    let complete = true;
    const rendered = template.replace(/\{(\w+)\}/g, (match, key: string) => {
        if (params[key] === undefined) {
            complete = false;
            return match;
        }
        return String(params[key]);
    });
    return complete ? rendered : fallback;
}
