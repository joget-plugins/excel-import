# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single Joget DX 8 form-element plugin ("Excel Import"): drop an `.xlsx`/`.xls`/`.csv` into a
form, validate/preview it client-side, and on submit persist **each row as its own record** in a
target form/table. `README.md` is the authoritative functional reference (features, config tabs,
validation rules, storage model) — read it before changing behavior.

## Build

```bash
mvn clean install        # → target/excel-import-1.0.0.jar (OSGi bundle)
```

- Java 8 (`source`/`target` 1.8). `wflow-core` 8.0-SNAPSHOT is `provided` and pulled from the
  Joget Archiva repos — the build needs network access to them.
- **Tests are skipped** (`skipTests=true` in the POM). There is no test suite; verification is
  manual: build the JAR, upload via Joget **Settings → Manage Plugins**, exercise it in the Form
  Builder.

## Architecture

Three Java classes in `org.joget.plugin.melkart` plus front-end assets under
`src/main/resources/resources/`. The split exists because in Joget a form `Element` and a
`FormBinder` are **sibling types** — one class cannot be both.

- **`ExcelImport`** (element, ~1000 lines) — holds **all** configuration, renders the widget,
  builds the client config JSON (`buildClientConfig`), exposes parsed rows via `formatData()`,
  and mirrors the browser validation server-side in `selfValidate()`. Also contains the shared
  coercion logic (`coerce`/`coerceNumber`/`coerceDate`/`coerceBoolean`) that must stay
  behaviorally identical to the JS in `excel-import-lib.js`.
- **`ExcelImportBinder`** (real `FormBinder`) — `store()`/`load()` persist and reload the child
  rows. Holds **no config of its own**; reads everything from the host element.
- **`Activator`** — registers **only** `ExcelImport`. The binder is intentionally *not*
  registered so it never appears as a selectable binder in the form builder; the element hands
  it out via `getStoreBinder()`/`getLoadBinder()`.

Front-end (loaded by `templates/ExcelImport.ftl`, which lazy-loads libs once per page):
`resources/js/excel-import-lib.js` (parse/validate/preview/commit), `xlsx.full.min.js` (bundled
SheetJS — parsing is 100% client-side), plus the `excel-import*.css`/`excel-import-columns.js`
column-grid assets. Parsed rows round-trip to the server as a JSON array in a hidden `<input>`.

## Non-obvious invariants (get these wrong and it breaks silently)

- **Two-sided validation must match.** The same checks (required headers, required cells,
  per-column value rules — allowed values / min-max range / regex pattern, composite unique key)
  run in `excel-import-lib.js` and again in `ExcelImport.selfValidate()`. Coercion runs *before*
  validation on both sides, and the per-column value rules only fire on non-empty cells. The
  per-cell logic (`validateValue` in JS ↔ `ExcelImport.validateCell`) and its number/date parsing
  must stay behaviorally identical — note the regex uses partial-match semantics (`RegExp.test`
  ↔ `Matcher.find()`). Change one side → change the other, or the browser and server disagree.
- **Binder instance caching.** Joget keys the `formatData()` row set by the exact binder
  instance returned from `getStoreBinder()`, then looks it up again to call `store()`. The
  element must return a **cached** binder instance, not a fresh one each call.
- **Host-record suppression** relies on Joget running store binders depth-first with the root
  form's own binder **last**. `ExcelImportBinder` clears the root binder's row set before it
  runs so the default `WorkflowFormBinder` writes nothing. See `suppressHostRecordStore()`.
- **Parent value resolution order** (`resolveParentValue`): hidden-field carrier captured at
  render time → configured value with hash variables → current submission primary key. The
  carrier exists because request-param hash vars (e.g. `#requestParam.referent_id#`) are gone by
  submit time.
- **Duplicate render guard**: the form builder re-renders and *appends* markup on every
  property apply; the bootstrap script removes prior renders of the same element ID.

## Config & i18n

- Form-builder properties live in `src/main/resources/properties/ExcelImport.json`; every label
  and message resolves through `AppPluginUtil.getMessage(...)` against
  `src/main/resources/messages/ExcelImport.properties` (French defaults). Add a language via a
  locale-suffixed bundle (`ExcelImport_en.properties`).
- Every user-facing message (`msgRequired`, `msgDuplicate`, …) is overridable per element
  instance and falls back to the bundle. When adding a message, wire all three: JSON property,
  bundle key, and the `customMsg(...)` lookup.

## Commit convention

When committing, set **only the user's git author** — do **not** add a `Co-Authored-By: Claude`
trailer.
