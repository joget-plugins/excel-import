/**
 * Excel Import Library (Joget ExcelImport plugin)
 *
 * Wires up the import widget whose markup is rendered server-side in ExcelImport.ftl
 * (drop zone, file bar, error/summary/preview containers) and styled by excel-import.css.
 * It parses the selected .xlsx/.xls client-side with the bundled SheetJS (XLSX),
 * validates (required headers, required cells across columns, composite duplicate
 * key across columns), previews the data, and writes the validated rows -- as a
 * JSON array keyed by Excel header -- into the element's hidden input.
 *
 * The same validation is enforced server-side by ExcelImport.selfValidate, which
 * blocks the submission on any error.
 *
 * Guarded so it is harmless if the script is included more than once on a page.
 */
(function () {
    if (window.initExcelImport) {
        return;
    }

    // Only Excel files are permitted.
    var ACCEPTED_EXT = [".xlsx", ".xls"];

    function ExcelImport(config) {
        this.config = config || {};
        this.headers = this.config.headers || [];
        // Columns shown in the preview table. A column can be present in `headers` (read from
        // the file, validated and/or stored) yet omitted here to keep it out of the preview.
        this.previewHeaders = this.config.previewHeaders || this.headers;
        this.requiredColumns = this.config.requiredColumns || [];
        this.uniqueColumns = this.config.uniqueColumns || [];
        this.caseSensitive = !!this.config.caseSensitive;
        this.messages = this.config.messages || {};
        this.container = document.getElementById(this.config.containerId);
        this.hiddenInput = document.querySelector('input.excel-import-input[name="' + this.config.fieldName + '"]');
        this.fileNameInput = this.config.fileNameField
            ? document.querySelector('input.excel-import-filename[name="' + this.config.fileNameField + '"]')
            : null;
        this.fileInput = null;
    }

    ExcelImport.prototype.msg = function (key, a0, a1) {
        var t = this.messages[key] || key;
        return t.replace("{0}", a0 == null ? "" : a0).replace("{1}", a1 == null ? "" : a1);
    };

    ExcelImport.prototype.init = function () {
        if (!this.container) {
            return;
        }
        this.cacheElements();
        if (!this.dropzone) {
            return;
        }
        this.bindEvents();
        // Restore from an existing value: an edit (data rebuilt from stored records) or a
        // validation reload (the submitted rows round-tripped in the hidden field).
        var existing = this.hiddenInput ? this.hiddenInput.value : "";
        if (existing && existing.trim()) {
            try {
                var rows = JSON.parse(existing);
                if (rows && rows.length) {
                    // Restore the file bar (name + Retirer) when the file name was carried through,
                    // since the browser cannot repopulate the file input itself.
                    var savedName = this.fileNameInput ? this.fileNameInput.value : "";
                    if (savedName) {
                        this.showFileBar(savedName);
                    }
                    this.process(rows, null);
                }
            } catch (e) { /* ignore malformed existing value */ }
        }
    };

    /** Switch the widget to the "file selected" state: hide the drop zone, show the file bar. */
    ExcelImport.prototype.showFileBar = function (name) {
        this.dropzone.classList.add("ei-hidden");
        this.filebar.classList.add("active");
        this.filebar.querySelector(".ei-fname").textContent = name;
    };

    /** Cache references to the widget elements rendered server-side in the FTL template. */
    ExcelImport.prototype.cacheElements = function () {
        this.dropzone = this.container.querySelector(".excel-import-dropzone");
        this.fileInput = this.container.querySelector('input[type="file"]');
        this.filebar = this.container.querySelector(".excel-import-filebar");
        this.errorBox = this.container.querySelector(".excel-import-error");
        this.summaryBox = this.container.querySelector(".excel-import-summary");
        // The preview is rendered as a sibling of the widget (outside this.container) so it can
        // span the full form width, and is omitted entirely when "hide preview" is enabled.
        this.previewBox = document.getElementById(this.config.containerId + "-preview");
    };

    ExcelImport.prototype.bindEvents = function () {
        var self = this;

        this.dropzone.addEventListener("click", function () {
            self.fileInput.click();
        });
        this.fileInput.addEventListener("change", function () {
            if (self.fileInput.files && self.fileInput.files.length) {
                self.onFile(self.fileInput.files[0]);
            }
        });
        ["dragenter", "dragover"].forEach(function (ev) {
            self.dropzone.addEventListener(ev, function (e) {
                e.preventDefault();
                e.stopPropagation();
                self.dropzone.classList.add("dragover");
            });
        });
        ["dragleave", "drop"].forEach(function (ev) {
            self.dropzone.addEventListener(ev, function (e) {
                e.preventDefault();
                e.stopPropagation();
                self.dropzone.classList.remove("dragover");
            });
        });
        this.dropzone.addEventListener("drop", function (e) {
            if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
                self.onFile(e.dataTransfer.files[0]);
            }
        });
        this.container.querySelector(".excel-import-remove").addEventListener("click", function () {
            self.clear();
        });
    };

    ExcelImport.prototype.clear = function () {
        if (this.hiddenInput) {
            this.hiddenInput.value = "";
        }
        if (this.fileNameInput) {
            this.fileNameInput.value = "";
        }
        if (this.fileInput) {
            this.fileInput.value = "";
        }
        this.filebar.classList.remove("active");
        this.filebar.querySelector(".ei-fname").textContent = "";
        this.dropzone.classList.remove("ei-hidden");
        if (this.previewBox) {
            this.previewBox.style.display = "none";
            this.previewBox.innerHTML = "";
        }
        this.summaryBox.style.display = "none";
        this.hideError();
    };

    ExcelImport.prototype.onFile = function (file) {
        var self = this;
        this.hideError();

        // Extension check (only Excel files are permitted).
        var name = file.name.toLowerCase();
        var okExt = ACCEPTED_EXT.some(function (ext) { return name.slice(-ext.length) === ext; });
        if (!okExt) {
            this.showError([this.msg("readError")]);
            return;
        }
        // Size check.
        var maxMB = parseFloat(this.config.maxFileSizeMB) || 0;
        if (maxMB > 0 && file.size > maxMB * 1024 * 1024) {
            this.showError([this.msg("fileTooLarge", maxMB)]);
            return;
        }

        this.showFileBar(file.name);
        // Persist the name so it round-trips and the file bar can be restored on a validation reload.
        if (this.fileNameInput) {
            this.fileNameInput.value = file.name;
        }

        var reader = new FileReader();
        reader.onload = function (e) {
            try {
                var wb = XLSX.read(e.target.result, { type: "array", cellDates: true, raw: false });
                var sheet = wb.Sheets[wb.SheetNames[0]];
                var rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
                if (!rows || !rows.length) {
                    self.showError([self.msg("emptyFile")]);
                    self.clearData();
                    return;
                }
                self.process(rows, file);
            } catch (err) {
                self.showError([self.msg("readError") + " " + (err && err.message ? err.message : "")]);
                self.clearData();
            }
        };
        reader.onerror = function () {
            self.showError([self.msg("readError")]);
            self.clearData();
        };
        reader.readAsArrayBuffer(file);
    };

    /** Read a cell from a parsed row by header, honouring case sensitivity. */
    ExcelImport.prototype.cell = function (row, header) {
        if (row == null) { return ""; }
        if (Object.prototype.hasOwnProperty.call(row, header)) {
            return row[header] == null ? "" : String(row[header]);
        }
        if (!this.caseSensitive) {
            var keys = Object.keys(row);
            for (var i = 0; i < keys.length; i++) {
                if (keys[i].toLowerCase() === String(header).toLowerCase()) {
                    return row[keys[i]] == null ? "" : String(row[keys[i]]);
                }
            }
        }
        return "";
    };

    ExcelImport.prototype.hasHeader = function (row, header) {
        if (row == null) { return false; }
        if (Object.prototype.hasOwnProperty.call(row, header)) { return true; }
        if (!this.caseSensitive) {
            return Object.keys(row).some(function (k) { return k.toLowerCase() === String(header).toLowerCase(); });
        }
        return false;
    };

    /** Validate, preview, and (when valid) commit the rows to the hidden field. */
    ExcelImport.prototype.process = function (rows, file) {
        var self = this;
        var errors = [];

        // 1. Header structure.
        var missing = this.headers.filter(function (h) { return !self.hasHeader(rows[0], h); });
        if (missing.length) {
            errors.push(this.msg("missingHeaders", missing.join(", ")));
            this.showError(errors);
            this.renderPreview(rows, {}, {});
            // Commit so the server also detects the missing headers (see commit() / process()).
            this.commit(rows);
            this.summary(0, rows.length);
            return;
        }

        var badCells = {};   // rowIndex -> { header: true }
        var rowBad = {};     // rowIndex -> true (has any required-cell error)
        var rowDup = {};     // rowIndex -> true (duplicate)

        // 2. Required cells across the flagged columns.
        this.requiredColumns.forEach(function (header) {
            var badRows = [];
            rows.forEach(function (row, i) {
                var v = self.cell(row, header);
                if (!v || !v.trim()) {
                    badRows.push(i + 1);
                    badCells[i] = badCells[i] || {};
                    badCells[i][header] = true;
                    rowBad[i] = true;
                }
            });
            if (badRows.length) {
                errors.push(self.msg("requiredCell", header, badRows.join(", ")));
            }
        });

        // 3. Composite duplicate key across the flagged columns (within file).
        if (this.uniqueColumns.length) {
            var seen = {};
            var dupRows = [];
            rows.forEach(function (row, i) {
                var key = self.uniqueColumns.map(function (h) {
                    var v = self.cell(row, h).trim();
                    return self.caseSensitive ? v : v.toLowerCase();
                }).join("");
                if (Object.prototype.hasOwnProperty.call(seen, key)) {
                    dupRows.push(i + 1);
                    rowDup[i] = true;
                    rowDup[seen[key]] = true; // also flag the first occurrence
                } else {
                    seen[key] = i;
                }
            });
            if (dupRows.length) {
                errors.push(this.msg("duplicate", this.uniqueColumns.join(" + "), dupRows.join(", ")));
            }
        }

        this.renderPreview(rows, badCells, rowDup, rowBad);

        // Always commit the parsed rows -- even when invalid -- so the data round-trips with the
        // submission. This keeps the file/preview from being lost on a validation reload and lets
        // the server-side selfValidate (which mirrors these checks) report the specific errors
        // (missing headers, required cells, duplicates) instead of a generic "required".
        this.commit(rows);
        if (errors.length) {
            this.showError(errors);
            // Count the rows that actually failed a per-row check (empty required cell or
            // duplicate) so the summary reflects how many rows are still valid, e.g. "17 / 19".
            var invalid = 0;
            for (var i = 0; i < rows.length; i++) {
                if (rowBad[i] || rowDup[i]) { invalid++; }
            }
            this.summary(rows.length - invalid, rows.length);
        } else {
            this.hideError();
            this.summary(rows.length, rows.length);
        }
    };

    /**
     * Write the parsed rows (keyed by Excel header) to the hidden field.
     *
     * Only headers actually present in the row are written: fabricating absent headers with an
     * empty default would hide a missing column from the server-side missing-header check.
     */
    ExcelImport.prototype.commit = function (rows) {
        if (!this.hiddenInput) { return; }
        var self = this;
        var payload = rows.map(function (row) {
            var obj = {};
            self.headers.forEach(function (h) {
                if (self.hasHeader(row, h)) {
                    obj[h] = self.cell(row, h);
                }
            });
            return obj;
        });
        this.hiddenInput.value = JSON.stringify(payload);
    };

    ExcelImport.prototype.clearData = function () {
        if (this.hiddenInput) {
            this.hiddenInput.value = "";
        }
    };

    ExcelImport.prototype.summary = function (valid, total) {
        if (!total) {
            this.summaryBox.style.display = "none";
            return;
        }
        this.summaryBox.style.display = "block";
        if (valid === total) {
            this.summaryBox.className = "excel-import-summary ok";
            this.summaryBox.textContent = this.msg("rowsValid", valid, total);
        } else {
            this.summaryBox.className = "excel-import-summary";
            this.summaryBox.textContent = this.msg("rowsValid", valid, total);
        }
    };

    ExcelImport.prototype.renderPreview = function (rows, badCells, rowDup, rowBad) {
        if (!this.previewBox) { return; } // preview hidden via config
        badCells = badCells || {};
        rowDup = rowDup || {};
        rowBad = rowBad || {};
        var self = this;
        var maxH = parseInt(this.config.previewHeight, 10) || 400;

        var thead = "<tr><th class='excel-import-rownum'>#</th>" +
            this.previewHeaders.map(function (h) { return "<th>" + escapeHtml(h) + "</th>"; }).join("") + "</tr>";

        var tbody = rows.map(function (row, i) {
            var cls = "";
            if (rowBad[i]) { cls = "ei-row-bad"; }
            else if (rowDup[i]) { cls = "ei-row-dup"; }
            var tds = self.previewHeaders.map(function (h, idx) {
                var cellCls = (badCells[i] && badCells[i][h]) ? " class='ei-cell-bad'" : "";
                var val = escapeHtml(self.cell(row, h));
                var chip = "";
                if (idx === 0 && rowDup[i]) { chip = " <span class='excel-import-chip dup'>doublon</span>"; }
                if (idx === 0 && rowBad[i]) { chip += " <span class='excel-import-chip bad'>requis</span>"; }
                return "<td" + cellCls + ">" + val + chip + "</td>";
            }).join("");
            return "<tr class='" + cls + "'><td class='excel-import-rownum'>" + (i + 1) + "</td>" + tds + "</tr>";
        }).join("");

        this.previewBox.style.display = "block";
        this.previewBox.innerHTML =
            "<div class='ei-scroll' style='max-height:" + maxH + "px'>" +
            "<table><thead>" + thead + "</thead><tbody>" + tbody + "</tbody></table></div>";
    };

    ExcelImport.prototype.showError = function (errors) {
        var html = "";
        if (errors.length === 1) {
            html = escapeHtml(errors[0]);
        } else {
            html = "<ul>" +
                errors.map(function (e) { return "<li>" + escapeHtml(e) + "</li>"; }).join("") + "</ul>";
        }
        this.errorBox.innerHTML = html;
        this.errorBox.style.display = "block";
    };

    ExcelImport.prototype.hideError = function () {
        this.errorBox.style.display = "none";
        this.errorBox.innerHTML = "";
    };

    function escapeHtml(s) {
        if (s == null) { return ""; }
        return String(s)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }

    window.initExcelImport = function (config) {
        var instance = new ExcelImport(config);
        instance.init();
        return instance;
    };
})();
