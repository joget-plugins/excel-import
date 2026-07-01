# Excel Import — Joget Form Element Plugin

A self-contained **Excel import** form element for [Joget](https://www.joget.org/) (DX 8).
Drop an `.xlsx` / `.xls` / `.csv` file into a form, preview and validate it in the browser, and on
submit persist **every row as its own record** in a configurable target form or table — each
row optionally linked to the host submission by a foreign-key column.

Parsing happens entirely client-side with a bundled copy of [SheetJS](https://sheetjs.com/)
(`xlsx.full.min.js`), so no file is uploaded to the server for parsing. The validated rows
are round-tripped through a hidden field as a JSON array and re-validated server-side before
anything is written.

---

## Table of contents

- [Features](#features)
- [How it works](#how-it-works)
- [Architecture](#architecture)
- [Requirements](#requirements)
- [Build](#build)
- [Install](#install)
- [Usage](#usage)
- [Configuration reference](#configuration-reference)
  - [1. Configuration](#1-configuration)
  - [2. Data storage](#2-data-storage)
  - [3. Display options](#3-display-options)
  - [4. Error messages](#4-error-messages)
- [Storage model](#storage-model)
- [Validation](#validation)
- [Internationalisation](#internationalisation)
- [Project structure](#project-structure)
- [Developer notes](#developer-notes)

---

## Features

- **Client-side parsing** — `.xlsx` / `.xls` / `.csv` read in the browser with bundled SheetJS
  (the delimiter is auto-detected for CSV); nothing is uploaded just to be parsed.
- **Drag-and-drop widget** — drop zone, file bar, inline error panel and a live preview table,
  all rendered by the plugin (no external CSS/JS dependencies).
- **Header → field mapping** — map each Excel column header to a target field ID via a grid.
- **Per-column controls** — mark a column *required*, part of a *unique key*, or *hidden from
  the preview*; leave the field ID blank to show a column in the preview without storing it.
- **Type coercion & cleansing** — per column, normalize the value to *number* (`"1,234.50"` →
  `1234.5`), *date* (Excel serial or common date string → a chosen format, `yyyy-MM-dd` by
  default), or *boolean* (`oui`/`no`/`1`/`0`… → `true`/`false`), and supply a *default value* for
  empty cells. Applied consistently in the preview and server-side before validation and storage.
- **Two-sided validation** — required headers, required cells, and composite duplicate keys are
  checked in the browser (with a detailed per-row breakdown) **and** re-checked server-side,
  which blocks the submission.
- **Existing-duplicate check** — optionally reject rows whose unique key already exists in the
  target table.
- **Flexible storage target** — store into the current form, a different form, or a raw table.
- **Parent linking** — write a foreign-key value (supports hash variables) into each row so the
  imported rows belong to the host submission.
- **Replace strategies** — on re-submit, replace previous rows *by parent column* or *by unique
  key(s)* (upsert).
- **Edit-friendly** — on edit, the preview is rebuilt from the stored child records; an unchanged
  re-submit is a no-op.
- **Optional host-record suppression** — use a form purely as an import trigger without storing
  its own parent record.
- **Fully localisable** — all labels and messages come from a resource bundle, and every
  user-facing message is overridable per element instance.

---

## How it works

```
┌──────────────────────────── Browser ────────────────────────────┐
│  1. User drops an .xlsx / .xls / .csv into the drop zone         │
│  2. SheetJS parses it → array of row objects (keyed by header)   │
│  3. excel-import-lib.js validates: headers, required cells,      │
│     duplicate keys → renders the preview + summary               │
│  4. Valid rows are written to a hidden <input> as a JSON array   │
└──────────────────────────────┬───────────────────────────────────┘
                               │ form submit (JSON round-trips)
┌──────────────────────────────▼───────────────────────────────────┐
│                             Server                                │
│  5. ExcelImport.selfValidate() re-runs the same checks and       │
│     blocks the submit on any error                                │
│  6. ExcelImport.formatData() turns the JSON into a FormRowSet     │
│     (one FormRow per Excel row, keyed by target field ID)         │
│  7. ExcelImportBinder.store() writes each row to the target       │
│     form/table, stamping the parent FK and applying the           │
│     replace strategy                                              │
└───────────────────────────────────────────────────────────────────┘
```

On **edit**, `ExcelImportBinder.load()` reads the child rows back and the element rebuilds the
hidden-field JSON so the preview is restored and an unchanged re-submit re-stores identical data.

---

## Architecture

The plugin is deliberately split into two classes because in Joget a form `Element` and a
`FormBinder` are **sibling types** — a single class cannot be both.

| Class | Role |
|-------|------|
| `ExcelImport` | The form **element**. Renders the widget, builds the client config, exposes the parsed rows via `formatData()`, mirrors the front-end validation in `selfValidate()`, and holds **all** configuration. |
| `ExcelImportBinder` | The **store/load binder** (a real `FormBinder`). Supplied by the element through `getStoreBinder()` / `getLoadBinder()`. Persists and reloads the child rows. Carries no configuration of its own — it reads everything from the host element. |
| `Activator` | OSGi bundle activator. Registers **only** `ExcelImport`; the binder is instantiated directly by the element, so it never appears as a selectable binder in the form builder. |

Front-end assets:

| File | Role |
|------|------|
| `templates/ExcelImport.ftl` | Server-side FreeMarker template: widget markup, hidden fields, and the bootstrap script that lazy-loads the libraries once per page. |
| `resources/js/excel-import-lib.js` | The import logic: parsing, validation, preview, and committing rows to the hidden field. |
| `resources/js/xlsx.full.min.js` | Bundled SheetJS parser. |
| `resources/css/excel-import.css` | Widget styling. |

---

## Requirements

- **Joget DX 8** (`wflow-core` `8.0-SNAPSHOT`, provided at runtime).
- **Java 8** (source/target `1.8`).
- **Maven 3** with access to the Joget Archiva repositories (for `wflow-core`).

---

## Build

```bash
mvn clean install
```

The output is an OSGi bundle JAR under `target/`:

```
target/excel-import-1.0.0.jar
```

> Tests are skipped by the build configuration (`skipTests=true`).

---

## Install

1. In Joget, open **Settings → Manage Plugins**.
2. Upload `target/excel-import-1.0.0.jar`.
3. The **Excel Import** element now appears in the Form Builder palette under the
   **Custom** category (spreadsheet icon).

To uninstall, remove the bundle from **Manage Plugins**.

---

## Usage

1. In the **Form Builder**, drag **Excel Import** onto your form.
2. Open its properties and set an **ID** (letters, digits, underscore only).
3. Under **Columns / Mapping**, add one row per Excel column:
   - **Excel header** — the exact header text in the spreadsheet.
   - **Field ID (target)** — the field to store it into (leave blank to preview only).
   - **Type** — optional coercion: *Text* (default), *Number*, *Date*, or *Boolean*.
   - **Required / Unique key / Hide from preview** — per-column flags.
   - **Default value / Date format** — optional cleansing applied before validation and storage.
4. Configure the **Data storage** tab (target, parent link, replace strategy).
5. Save. At runtime, users drop a file, see a validated preview, and on submit each row is
   stored as a record in the target.

---

## Configuration reference

Properties are grouped into four tabs.

### 1. Configuration

| Property | Description |
|----------|-------------|
| **ID** | Element ID. Must match `^[a-zA-Z0-9_]+$`. |
| **Label** | Optional label shown above the widget. |
| **Required import** | Blocks submission if no valid file has been imported. |
| **Columns / Mapping** | Grid mapping each Excel header to a target field. Columns: *Excel header*, *Field ID (target)*, *Type*, *Required*, *Unique key*, *Default value*, *Date format*, *Hide from preview*. Blank field ID = preview-only (not stored). |
| **Case-sensitive headers** | When unchecked (default), Excel headers are matched case-insensitively. |

#### Column type coercion & cleansing

Each mapped column can normalize its cell values before validation and storage. Coercion runs
identically in the browser preview and server-side, and is **best-effort**: a value that cannot be
parsed for the requested type is left untouched (so the offending input stays visible).

| Option | Effect |
|--------|--------|
| **Type = Number** | Strips currency/spaces and resolves separators. When both `,` and `.` appear, the last is the decimal (`"1,234.50"` → `1234.5`). A lone `,` is grouped thousands for `1,000` / `1,234,567`, otherwise a decimal (`"12,5"` → `12.5`). |
| **Type = Date** | Parses an Excel serial number (1900 date system), an ISO `yyyy-MM-dd`, or a day-first `dd/MM/yyyy` date, and reformats it to **Date format** (tokens `yyyy MM dd HH mm ss`; `yyyy-MM-dd` by default). |
| **Type = Boolean** | Maps `true/1/yes/y/oui/o/vrai/x/on` → `true` and `false/0/no/n/non/faux/off` → `false` (case-insensitive). |
| **Default value** | Substituted when the cell is empty, then coerced like any other value. |

### 2. Data storage

| Property | Description |
|----------|-------------|
| **Do not store host form record** | Prevents the host form's own store binder from persisting its (parent) record; only the Excel rows are stored. Useful when the form is only an import trigger. |
| **Custom storage target** | By default rows are stored in the current form/table. Enable to choose another target. |
| **Target type** | *Form* or *Table name* (shown when custom target is on). |
| **Target form** | Form (table) to store each row into. |
| **Table name** | Raw target table. An `app_fd_` prefix is stripped automatically. |
| **Parent column (foreign key)** | Optional. Column in the target that receives the parent value. Leave blank for standalone rows (no link). |
| **Parent column value** | Optional. Value written to the parent column; hash variables are supported. Falls back to the current submission's primary key when blank. |
| **Replace existing rows** | On submit, delete existing rows before saving the new ones (per the chosen strategy). *On by default.* |
| **Replace strategy** | *By parent column* (delete rows linked to the parent value) or *By unique key(s)* (upsert: delete rows whose unique-key combination matches the imported rows). |
| **Check existing duplicates** | Reject rows whose unique key already exists in the target table (excluding the current submission's own rows). |

### 3. Display options

| Property | Description |
|----------|-------------|
| **Max size (MB)** | Maximum accepted file size. Default `5`. |
| **Hide preview** | Hides the preview table. Default off (preview shown full-width under the drop zone). |
| **Preview height (px)** | Max height of the scrollable preview. Default `400`. |
| **Drop zone text** | Custom prompt text for the drop zone. |
| **Full width** | Renders the drop zone across the full form width (`flex 0 0 100%`). Default off (behaves like a normal Joget field). |

### 4. Error messages

Every user-facing message can be overridden per element. Each falls back to its localised
default from the resource bundle when left blank:

`msgRequired`, `msgInvalidData`, `msgMissingHeaders`, `msgRequiredCell`, `msgDuplicate`,
`msgExistingDuplicate`, `msgEmptyFile`, `msgReadError`, `msgFileTooLarge`, `msgRowsValid`.

Messages support positional placeholders (`{0}`, `{1}`) — e.g. the missing-columns list, the
offending column/rows, or the valid/total row counts.

---

## Storage model

- Each Excel row becomes **one `FormRow`**, keyed by the mapped **target field IDs**.
- When a **parent column** and **value** resolve, that value is stamped onto every row so the
  imported rows belong to the host submission.
- The **parent value** resolves in this order:
  1. A value captured at render time and carried through the submission via a hidden field —
     essential for request-param hash variables (e.g. `#requestParam.referent_id#`), because the
     submit POST no longer carries the original load-URL query parameters.
  2. The configured **parent column value** with hash variables resolved.
  3. The current submission's primary key.
- Rows are written via `appService.storeFormData(...)`, which handles ID generation and the
  created/modified metadata.

---

## Validation

The same three checks run **in the browser** (with a detailed, per-row breakdown and highlighted
cells) and **on the server** (`ExcelImport.selfValidate`, which blocks the submit):

1. **Header structure** — every mapped Excel header must be present in the file.
2. **Required cells** — flagged columns must have a non-empty value in every row.
3. **Composite duplicate key** — the combination of *unique key* columns must be unique within
   the file.

> [Type coercion & cleansing](#column-type-coercion--cleansing) runs **before** these checks, so
> validation sees the cleansed values: a **default value** can satisfy a *required* cell, and the
> duplicate checks compare the coerced (e.g. normalized number/date) values that get stored.

An optional fourth check runs server-side when **Check existing duplicates** is enabled:

4. **Existing duplicates** — reject rows whose unique key already exists in the target table
   (excluding rows belonging to the current parent, which are being replaced).

> The server reports a **single generic** "invalid data" error for any failed check; the
> detailed breakdown (which columns, which rows) is surfaced client-side. The rows always
> round-trip through the hidden field — even when invalid — so the file and preview survive a
> validation reload.

---

## Internationalisation

All labels and messages live in `src/main/resources/messages/ExcelImport.properties` and are
resolved through `AppPluginUtil.getMessage(...)`. The bundle ships with **French** defaults.
To add another language, provide a locale-suffixed bundle (e.g. `ExcelImport_en.properties`)
following Joget's message-bundle conventions.

---

## Project structure

```
excel-import/
├── pom.xml
└── src/main/
    ├── java/org/joget/plugin/melkart/
    │   ├── Activator.java            # OSGi bundle activator (registers ExcelImport)
    │   ├── ExcelImport.java          # Form element: render, config, formatData, selfValidate
    │   └── ExcelImportBinder.java    # Store/load binder: persist & reload child rows
    └── resources/
        ├── messages/ExcelImport.properties   # i18n bundle (labels + messages)
        ├── properties/ExcelImport.json       # Form-builder property configuration
        ├── templates/ExcelImport.ftl         # Widget markup + bootstrap script
        └── resources/
            ├── css/excel-import.css
            └── js/
                ├── excel-import-lib.js        # Parse / validate / preview / commit
                └── xlsx.full.min.js           # Bundled SheetJS
```

---

## Developer notes

- **Why two classes?** `Element` and `FormBinder` are sibling types in Joget; a class can only
  be one. The element supplies a cached `ExcelImportBinder` through `getStoreBinder()` /
  `getLoadBinder()`. Caching matters: Joget keys the `formatData()` row set by the exact binder
  instance returned and later looks it up again to call `store()`.
- **Host-record suppression** relies on Joget executing store binders depth-first, with the root
  form's own binder running **last**. This (child) binder clears the root binder's row set before
  it runs, so the default `WorkflowFormBinder` short-circuits on the empty set and writes nothing.
- **Duplicate render guard** — the form builder re-renders and *appends* the element markup each
  time properties are applied. The bootstrap script removes earlier renders of the same element
  (matched by its stable ID) so drop zones don't stack.
- **Preview placement** — the preview table is rendered as a sibling of the widget (outside the
  field cell) so it can span the full form width regardless of the field's column width.
- The binder is intentionally **not** registered in the `Activator`, so it never shows up as a
  selectable binder in the form builder.

---

## License

Distributed as a Joget plugin. See your organisation's licensing terms.
