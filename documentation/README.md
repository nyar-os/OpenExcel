# OpenExcel maintainer guide

Repository layout, build and test commands, dependency pins, CI and the actual state of the npm release path. User-facing documentation lives in [../readme.md](../readme.md) and [../projects/packages/open-excel/README.md](../projects/packages/open-excel/README.md); agent-facing documentation lives in [agents.md](agents.md).

## Project layout

| Path | What it is |
| --- | --- |
| `projects/crates/model` | Rust crate `open-excel-workbook`: project model, XLSX/CSV import, XLSX writer, inspection rules, formula evaluation, SVG and PDF projections, process-wide handle store. |
| `projects/crates/napi` | Rust crate `open-excel-napi`: `cdylib` Node-API and WebView binding. It maps JSON arguments onto model functions and hosts the Studio window; workbook logic stays in the model crate. |
| `projects/packages/open-excel` | TypeScript package `@wae/open-excel`: session API (`src/open-excel.ts`), CLI (`src/cli.ts`, bin `open-excel`), types, stable errors, format registry and `locales/en-US/cli.json` help text. |
| `projects/packages/open-excel-ui` | `@wae/open-excel-ui`: VMZ static shell for the Studio window. `vmz build . --profile static` emits `dist/static`, staged into `open-excel/dist/gui`. |
| `projects/packages/open-excel-skills` | `@wae/open-excel-skills`: skill documents under `skills/`. No runtime library and no workbook algorithms. |
| `projects/packages/open-excel-<platform>-<arch>` | Five prebuilt addon packages: `win32-x64`, `darwin-x64`, `darwin-arm64`, `linux-x64`, `linux-arm64`. Each is a thin manifest around a staged `.node` file. |
| `projects/sites/homepage` | `@wae/open-excel-homepage`: VMZ product site. |
| `projects/sub_projects` | Optional nested project slot used by sibling repositories in this family. **OpenExcel has no such directory today.** |
| `wae.config.ts` | Build order, native targets, platform package mapping, staging, skills registration, CI job names and window behaviour. |
| `Cargo.toml` | Cargo workspace; members are `projects/crates/*`, default member is `napi`. |
| `pnpm-workspace.yaml` | pnpm workspace covering `projects/packages/*` and `projects/sites/*`. |

The CLI, SDK and Studio all read the same model crate. A behavior change in `crates/model` reaches every surface; `capabilities()` is the single disclosure point.

## Build and test

Prerequisites: Node.js 18 or newer, pnpm 10.28.0 (from `package.json`), Rust 1.90 with edition 2024 (from `rust-toolchain.toml` / `Cargo.toml`).

```bash
pnpm install                  # workspace dependencies
pnpm build:napi               # wae native build (crate open-excel-napi)
pnpm build                    # wae build: native -> ts -> ui -> skills -> staging
pnpm typecheck                # pnpm -r --filter "./projects/packages/*" run typecheck
pnpm test                     # package tests
cargo test -p open-excel-workbook
pnpm check:rust               # fmt check + clippy -D warnings + tests for model and napi
```

Useful narrower commands:

```bash
cargo test -p open-excel-workbook --release
cargo test -p open-excel-napi
pnpm build:ts                 # tsc for @wae/open-excel
pnpm build:ui                 # static Studio shell
pnpm stage:native             # wae native stage into platform package directories
pnpm gui                      # build and launch the Studio window
```

`wae.config.ts` declares the outputs that a full build must produce: `projects/packages/open-excel/dist/cli.js` and `projects/packages/open-excel/dist/gui/index.html`.

TypeScript tests live in `projects/packages/*/tests/` and compile to `dist-test/` via `tsc -p tsconfig.test.json`; Rust integration tests live in `projects/crates/*/tests/`.

## Studio build

`@wae/open-excel-ui` is a private VMZ package. Its build is part of `wae build` (`native → ts → ui → skills → staging`); run it directly with:

```bash
pnpm --filter @wae/open-excel-ui run check   # vmz check
pnpm --filter @wae/open-excel-ui run build   # vmz build --profile static
pnpm --filter @wae/open-excel-ui run serve   # local preview
```

`vmz build` emits `dist/static`, which the staging step copies to `projects/packages/open-excel/dist/gui` for the native WebView host.

## Dependency pins

Cross-repository dependencies pin a full 40-character revision, not a branch or tag:

- `projects/crates/model/Cargo.toml` pins `boost-fs`, `data-text-csv`, `document-formats`, `document-session` and `document-spreadsheet-formula` to `0ded81d40572c38177040e0bfaff7a0f8c78769c` from `oovm/rust-boost`.
- The workspace `Cargo.toml` pins `wae` to `f92f818fde7174154d77a87b55a5e44b367034c9` from `doki-land/wae-framework`.

The formula engine lives in `document-spreadsheet-formula`. Its function list is surfaced through `capabilities().formulas.functions`, and unknown functions keep the cached source value (`unsupportedBehavior: "preserve-cached"`). Treat a boost revision bump as a behavior change: move every boost pin together, rerun `cargo test -p open-excel-workbook`, and re-read the capability matrix before updating user documentation. The `u` script in the root `package.json` runs `cargo upgrade --workspace`.

## Release and publishing

The publishable set is `@wae/open-excel` plus the five platform packages. `@wae/open-excel-skills`, `@wae/open-excel-ui` and the homepage remain `private: true`, and `wae.config.ts` leaves `release` unset.

The intended npm release contract is:

1. Build and stage the native addon so each platform package contains its `.node` file.
2. Publish from CI with npm Trusted Publisher / OIDC, triggered by a version tag: `permissions: id-token: write` and `environment: NPM_PUBLISH` in `.github/workflows/publish-npm.yml`. No long-lived npm token is stored in workflow YAML.
3. Keep local publishing to inspection only: `nifty publish --dry-run`, or `nifty publish --placeholder` to claim never-published `0.0.0` names. Do not run a real-version publish locally; push the tag and let CI do it.
4. Keep `.secrets/nifty.json` local. The whole `.secrets/` directory is git-ignored, and the file may hold token or OTP material that must never be committed.

### Verified state

- `.github/workflows/publish-npm.yml` builds the native addon for all five targets, stages each `.node` into its platform package, builds the TypeScript surface and GUI, and publishes through `nifty publish` on `v*` tags. It requests `id-token: write` and uses the `NPM_PUBLISH` environment.
- `nifty.config.ts` declares the Trusted Publisher target (`nyar-os/OpenExcel`, `publish-npm.yml`, `NPM_PUBLISH`).
- `@doki-land/nifty` is a root dev dependency, so `pnpm exec nifty` is available locally.
- `wae.config.ts` sets `ci.workflows` to `check-rust.yml` and `check-typescript.yml` only, with required jobs `build` and `check`.
- `nifty publish` aborts when a native platform package has no staged `.node`, so the native job is a hard prerequisite.

## CI workflows

| Workflow | Trigger | What it runs |
| --- | --- | --- |
| `check-rust.yml` | Push and pull request to `master` / `dev` | Matrix of Ubuntu, macOS and Windows; installs nightly with `rustfmt` and `clippy`; `cargo build --workspace --release`; `cargo test --workspace --release`. It does not run fmt or clippy checks; `pnpm check:rust` does. |
| `check-typescript.yml` | Push and pull request to `master` / `dev` | Ubuntu; pnpm and Node 22; `pnpm install --no-frozen-lockfile`; `pnpm typecheck`; `pnpm test`. |
| `publish-npm.yml` | Tag push `v*` or workflow dispatch | Builds and stages the five native addons, builds TS/UI, then publishes through npm Trusted Publisher. |

## Documentation map

| File | Audience |
| --- | --- |
| [../readme.md](../readme.md) | Product landing page and GitHub entry point. |
| [../projects/packages/open-excel/README.md](../projects/packages/open-excel/README.md) | npm registry page for `@wae/open-excel`. |
| [agents.md](agents.md) | Agents and automation that drive the CLI or SDK. |
| This file | Maintainers changing the repository, CI or release path. |

The project is MPL-2.0; see [../License.md](../License.md). The skills package declares Apache-2.0 in its own `package.json`.
