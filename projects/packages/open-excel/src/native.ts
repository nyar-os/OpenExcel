/**
 * Native addon discovery.
 *
 * The addon only maps values: workbook semantics, inspection and export stay in
 * the shared engine. An absent addon is a declarative state, so calls report
 * `capability/not-implemented` instead of inventing a matrix.
 */

import { createRequire } from 'node:module';

const requireNative = createRequire(import.meta.url);

/** Orthogonal exports of the built addon. Every workbook call returns a JSON string. */
export interface NativeBinding {
    version(): string;
    capabilities(): string;
    workbook_create(requestJson: string): string;
    workbook_open(path: string): string;
    workbook_save(handle: number, path: string, overwrite: boolean): string;
    workbook_inspect(handle: number, checksJson?: string): string;
    workbook_summary(handle: number): string;
    workbook_edit(handle: number, requestJson: string): string;
    workbook_cell(handle: number, requestJson: string): string;
    workbook_range(handle: number, requestJson: string): string;
    workbook_recalc(handle: number): string;
    workbook_render(handle: number, requestJson?: string): string;
    workbook_export(handle: number, format: string, target: string, overwrite: boolean): string;
    workbook_close(handle: number): boolean;
    openGui(options?: { title?: string; contentDir?: string; width?: number; height?: number }): string;
    pumpEvents(): string;
    closeGui(): string;
    isGuiOpen(): boolean;
}

let cached: NativeBinding | null | undefined;

/** Loads the addon once, or returns `undefined` when it is not installed. */
export function nativeBinding(): NativeBinding | undefined {
    if (cached === undefined) {
        cached = load();
    }
    return cached ?? undefined;
}

function load(): NativeBinding | null {
    for (const candidate of candidates()) {
        try {
            return requireNative(candidate) as NativeBinding;
        } catch (error) {
            if (candidate === explicitPath()) {
                throw error;
            }
            if (!isModuleNotFound(error, candidate)) {
                throw error;
            }
        }
    }
    return null;
}

function candidates(): string[] {
    const explicit = explicitPath();
    if (explicit) {
        return [explicit];
    }
    return [`@wae/open-excel-${process.platform}-${process.arch}`];
}

function explicitPath(): string | undefined {
    const value = process.env.OPEN_EXCEL_NATIVE;
    return value !== undefined && value.length > 0 ? value : undefined;
}

function isModuleNotFound(error: unknown, candidate: string): boolean {
    if (typeof error !== 'object' || error === null) {
        return false;
    }
    const code = (error as { code?: unknown }).code;
    // Node prints the resolved path, which uses the platform separator on
    // Windows. Normalize it so a linked platform package without a built
    // addon is still treated as "not installed" instead of a fatal error.
    const message = (error as Error).message.replace(/\\/g, '/');
    return code === 'MODULE_NOT_FOUND' && message.includes(candidate);
}
