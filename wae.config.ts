/**
 * Build and release configuration for OpenExcel.
 *
 * Pure literals only: `wae` parses this file statically and never executes it.
 * Build order is fixed: native → ts → ui → skills → staging.
 *
 * The packages live under `projects/packages`; the native addon is built from
 * the `open-excel-napi` crate and staged into the per-platform packages.
 */
import { defineConfig } from '@wae/wae';

export default defineConfig({
    native: {
        crate: 'open-excel-napi',
        cdylib: 'open_excel_napi',
        targets: [
            {
                platform: 'win32',
                arch: 'x64',
                triple: 'x86_64-pc-windows-msvc',
                package: '@wae/open-excel-win32-x64',
                nodeFile: 'open-excel.win32-x64.node',
                dir: 'projects/packages/open-excel-win32-x64',
            },
            {
                platform: 'darwin',
                arch: 'x64',
                triple: 'x86_64-apple-darwin',
                package: '@wae/open-excel-darwin-x64',
                nodeFile: 'open-excel.darwin-x64.node',
                dir: 'projects/packages/open-excel-darwin-x64',
            },
            {
                platform: 'darwin',
                arch: 'arm64',
                triple: 'aarch64-apple-darwin',
                package: '@wae/open-excel-darwin-arm64',
                nodeFile: 'open-excel.darwin-arm64.node',
                dir: 'projects/packages/open-excel-darwin-arm64',
            },
            {
                platform: 'linux',
                arch: 'x64',
                triple: 'x86_64-unknown-linux-gnu',
                package: '@wae/open-excel-linux-x64',
                nodeFile: 'open-excel.linux-x64.node',
                dir: 'projects/packages/open-excel-linux-x64',
            },
            {
                platform: 'linux',
                arch: 'arm64',
                triple: 'aarch64-unknown-linux-gnu',
                package: '@wae/open-excel-linux-arm64',
                nodeFile: 'open-excel.linux-arm64.node',
                dir: 'projects/packages/open-excel-linux-arm64',
            },
        ],
    },
    packages: [
        {
            name: '@wae/open-excel-win32-x64',
            dir: 'projects/packages/open-excel-win32-x64',
            os: ['win32'],
            cpu: ['x64'],
            main: 'open-excel.win32-x64.node',
        },
        {
            name: '@wae/open-excel-darwin-x64',
            dir: 'projects/packages/open-excel-darwin-x64',
            os: ['darwin'],
            cpu: ['x64'],
            main: 'open-excel.darwin-x64.node',
        },
        {
            name: '@wae/open-excel-darwin-arm64',
            dir: 'projects/packages/open-excel-darwin-arm64',
            os: ['darwin'],
            cpu: ['arm64'],
            main: 'open-excel.darwin-arm64.node',
        },
        {
            name: '@wae/open-excel-linux-x64',
            dir: 'projects/packages/open-excel-linux-x64',
            os: ['linux'],
            cpu: ['x64'],
            main: 'open-excel.linux-x64.node',
        },
        {
            name: '@wae/open-excel-linux-arm64',
            dir: 'projects/packages/open-excel-linux-arm64',
            os: ['linux'],
            cpu: ['arm64'],
            main: 'open-excel.linux-arm64.node',
        },
    ],
    build: {
        ts: { packages: ['@wae/open-excel'] },
        ui: { package: '@wae/open-excel-ui' },
        stage: [
            {
                from: 'projects/packages/open-excel-ui/dist/static',
                to: 'projects/packages/open-excel/dist/gui',
                clean: true,
            },
        ],
        outputs: [
            'projects/packages/open-excel/dist/cli.js',
            'projects/packages/open-excel/dist/gui/index.html',
        ],
    },
    skills: {
        packages: [
            {
                dir: 'projects/packages/open-excel-skills/skills',
                packageJson: 'projects/packages/open-excel-skills/package.json',
                appPackageJson: 'projects/packages/open-excel/package.json',
                cliEntry: 'projects/packages/open-excel/dist/cli.js',
                commandName: 'open-excel',
            },
        ],
    },
    // `release` is intentionally unset while the app packages stay `private`;
    // the platform packages are the only publishable set and have no CI release
    // workflow yet.
    ci: {
        repo: 'nyar-os/OpenExcel',
        workflows: ['check-rust.yml', 'check-typescript.yml'],
        requiredJobs: ['build', 'check'],
        shaBinding: false,
    },
    install: { shortcutDefault: true },
    audit: { actor: { kind: 'flag', id: 'flag', source: 'flag' } },
    window: { mode: 'single', secondLaunch: 'focus' },
});
