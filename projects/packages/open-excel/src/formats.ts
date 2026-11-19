/**
 * Workbook format family.
 *
 * Excel extensions encode the container and the intended use: `x` = OOXML,
 * `m` = macro-enabled OOXML, no suffix = 97-2003 binary, `t` = template. The
 * registry below classifies extensions and fixes the default open intent.
 *
 * Classification is **not** a support claim. What the installed engine can read,
 * edit, inspect or write is reported by `OpenExcel.capabilities()`; an extension
 * listed here may still be `unsupported`.
 */

/** Intent applied when the caller does not override it. */
export type WorkbookIntent = 'edit' | 'data' | 'template';

/** Grouping used for capability reporting and user-facing wording. */
export type FormatFamily = 'workbook' | 'template' | 'legacy' | 'text' | 'document' | 'interchange';

/** Container encoding behind an extension. */
export type FormatMarkup = 'ooxml' | 'binary' | 'odf' | 'text' | 'json' | 'xml' | 'fixed';

/** What an extension can be consumed as. `none` means it never becomes a workbook. */
export type FormatOpenKind = 'workbook' | 'data' | 'none';

/** Every extension the product must classify. */
export const workbookFormats = ['xlsx', 'xlsm', 'xltx', 'xls', 'csv', 'tsv', 'ods', 'json', 'xml', 'pdf'] as const;

/** Union of {@link workbookFormats}. */
export type WorkbookFormat = (typeof workbookFormats)[number];

/** What a format is, independent of whether it is implemented. */
export interface WorkbookFormatInfo {
    readonly format: WorkbookFormat;
    readonly family: FormatFamily;
    readonly markup: FormatMarkup;
    /** Whether the container may hold VBA macros. Such parts are never executed. */
    readonly macroEnabled: boolean;
    readonly opensAs: FormatOpenKind;
    /** Intent applied to `open()` when the extension arrives without one. */
    readonly defaultIntent: WorkbookIntent | null;
}

type FormatRegistryEntry = Omit<WorkbookFormatInfo, 'format'>;

/** Extension to classification. Values describe the format, never a capability. */
export const formatRegistry: Readonly<Record<WorkbookFormat, FormatRegistryEntry>> = Object.freeze({
    xlsx: { family: 'workbook', markup: 'ooxml', macroEnabled: false, opensAs: 'workbook', defaultIntent: 'edit' },
    xlsm: { family: 'workbook', markup: 'ooxml', macroEnabled: true, opensAs: 'workbook', defaultIntent: 'edit' },
    xltx: { family: 'template', markup: 'ooxml', macroEnabled: false, opensAs: 'workbook', defaultIntent: 'template' },
    xls: { family: 'legacy', markup: 'binary', macroEnabled: true, opensAs: 'workbook', defaultIntent: 'edit' },
    csv: { family: 'text', markup: 'text', macroEnabled: false, opensAs: 'data', defaultIntent: 'data' },
    tsv: { family: 'text', markup: 'text', macroEnabled: false, opensAs: 'data', defaultIntent: 'data' },
    ods: { family: 'document', markup: 'odf', macroEnabled: false, opensAs: 'workbook', defaultIntent: 'edit' },
    json: { family: 'interchange', markup: 'json', macroEnabled: false, opensAs: 'data', defaultIntent: 'data' },
    xml: { family: 'interchange', markup: 'xml', macroEnabled: false, opensAs: 'workbook', defaultIntent: 'edit' },
    pdf: { family: 'document', markup: 'fixed', macroEnabled: false, opensAs: 'none', defaultIntent: null },
});

/** Returns the full classification of one format. */
export function formatInfo(format: WorkbookFormat): WorkbookFormatInfo {
    return { format, ...formatRegistry[format] };
}

/** Extracts a classified extension from a path, ignoring case. */
export function formatFromPath(path: string): WorkbookFormat | undefined {
    const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
    return (workbookFormats as readonly string[]).includes(extension) ? (extension as WorkbookFormat) : undefined;
}

/** Default open intent of a format, or `null` when it never opens as a workbook. */
export function defaultIntentFor(format: WorkbookFormat): WorkbookIntent | null {
    return formatRegistry[format].defaultIntent;
}
