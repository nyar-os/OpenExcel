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
    'workbook/revision-conflict',
    'object/not-found',
    'object/unsupported',
    'format/unsupported',
    'format/invalid',
    'capability/not-implemented',
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

/** Options accepted by {@link OpenExcelError}. */
export interface OpenExcelErrorOptions {
    readonly recoverable?: boolean;
    readonly details?: OpenExcelErrorDetails;
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

    constructor(code: OpenExcelErrorCode, message: string, options: OpenExcelErrorOptions = {}) {
        super(message, options.cause === undefined ? undefined : { cause: options.cause });
        this.name = 'OpenExcelError';
        this.code = code;
        this.recoverable = options.recoverable ?? false;
        this.details = Object.freeze({ ...options.details });
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
