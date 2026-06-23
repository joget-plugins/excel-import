/**
 * Excel Import Library (Joget ExcelParser plugin)
 *
 * Self-contained: renders its own drop zone (no dependency on Joget's Dropzone),
 * parses the selected .xlsx/.xls client-side with the bundled SheetJS (XLSX),
 * validates (required headers, required cells across columns, composite duplicate
 * key across columns), previews the data, and writes the validated rows -- as a
 * JSON array keyed by Excel header -- into the element's hidden input.
 *
 * The same validation is enforced server-side by ExcelParser.selfValidate, which
 * blocks the submission on any error.
 *
 * Guarded so it is harmless if the script is included more than once on a page.
 */
(function () {
    if (window.initExcelImport) {
        return;
    }

    var STYLE_ID = "excel-import-styles";

    function injectCSS() {
        if (document.getElementById(STYLE_ID)) {
            return;
        }
        var style = document.createElement("style");
        style.id = STYLE_ID;
        style.textContent = [
            ".excel-import-widget{width:100%;font-size:14px;}",
            ".excel-import-dropzone{border:2px dashed #b6c0e0;border-radius:10px;padding:30px 20px;text-align:center;color:#5a6b9c;background:#f7f9ff;cursor:pointer;transition:all .2s ease;}",
            ".excel-import-dropzone:hover,.excel-import-dropzone.dragover{border-color:#667eea;background:#eef2ff;color:#3d4f8a;}",
            ".excel-import-dropzone .ei-icon{font-size:2.4em;display:block;margin-bottom:8px;}",
            ".excel-import-dropzone .ei-hint{font-size:.85em;opacity:.7;margin-top:6px;}",
            ".excel-import-filebar{display:none;align-items:center;justify-content:space-between;gap:12px;padding:10px 14px;margin-top:10px;background:#eef2ff;border:1px solid #d3dbf5;border-radius:8px;}",
            ".excel-import-filebar.active{display:flex;}",
            ".excel-import-filebar .ei-fname{font-weight:600;color:#3d4f8a;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}",
            ".excel-import-remove{color:#dc3545;cursor:pointer;font-weight:600;white-space:nowrap;}",
            ".excel-import-summary{margin-top:12px;font-weight:600;}",
            ".excel-import-summary.ok{color:#198754;}",
            ".excel-import-error{color:#842029;background:#f8d7da;border:1px solid #f5c2c7;border-radius:8px;padding:10px 14px;margin-top:12px;}",
            ".excel-import-error ul{margin:6px 0 0 0;padding-left:18px;}",
            ".excel-import-preview{margin-top:14px;border-radius:10px;overflow:hidden;box-shadow:0 6px 18px rgba(0,0,0,.08);}",
            ".excel-import-preview .ei-scroll{overflow:auto;}",
            ".excel-import-preview table{width:100%;border-collapse:separate;border-spacing:0;margin:0;}",
            ".excel-import-preview thead th{position:sticky;top:0;z-index:5;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:#fff;padding:12px 14px;text-align:left;font-weight:600;white-space:nowrap;}",
            ".excel-import-preview tbody td{padding:10px 14px;border-bottom:1px solid #eceff5;vertical-align:middle;}",
            ".excel-import-preview tbody tr:nth-child(even){background:#f8f9fc;}",
            ".excel-import-preview tbody tr.ei-row-bad{background:#fff3f4 !important;}",
            ".excel-import-preview tbody tr.ei-row-dup{background:#fff8e6 !important;}",
            ".excel-import-preview td.ei-cell-bad{box-shadow:inset 0 0 0 2px #dc3545;border-radius:4px;}",
            ".excel-import-rownum{color:#9aa3bd;font-variant-numeric:tabular-nums;width:1%;}",
            ".excel-import-chip{display:inline-block;padding:2px 8px;border-radius:12px;font-size:.75em;margin-left:6px;color:#fff;}",
            ".excel-import-chip.dup{background:#ffc107;color:#000;}",
            ".excel-import-chip.bad{background:#dc3545;}"
        ].join("\n");
        document.head.appendChild(style);
    }

    function ExcelImport(config) {
        this.config = config || {};
        this.headers = this.config.headers || [];
        this.requiredColumns = this.config.requiredColumns || [];
        this.uniqueColumns = this.config.uniqueColumns || [];
        this.caseSensitive = !!this.config.caseSensitive;
        this.messages = this.config.messages || {};
        this.container = document.getElementById(this.config.containerId);
        this.hiddenInput = document.querySelector('input.excel-import-input[name="' + this.config.fieldName + '"]');
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
        injectCSS();
        this.render();
        this.bindEvents();
        // Restore preview from an existing value (edit mode).
        var existing = this.hiddenInput ? this.hiddenInput.value : "";
        if (existing && existing.trim()) {
            try {
                var rows = JSON.parse(existing);
                if (rows && rows.length) {
                    this.process(rows, null);
                }
            } catch (e) { /* ignore malformed existing value */ }
        }
    };

    ExcelImport.prototype.render = function () {
        this.container.innerHTML =
            '<div class="excel-import-dropzone">' +
            '  <span class="ei-icon">&#128228;</span>' +
            '  <span class="ei-text">' + escapeHtml(this.config.dropzoneText || "Choisir / déposer un fichier Excel") + '</span>' +
            '  <span class="ei-hint">' + escapeHtml(this.config.acceptedFiles || ".xlsx,.xls") + '</span>' +
            '  <input type="file" accept="' + escapeHtml(this.config.acceptedFiles || ".xlsx,.xls") + '" style="display:none" />' +
            '</div>' +
            '<div class="excel-import-filebar"><span class="ei-fname"></span><span class="excel-import-remove">&#10005; ' + "Retirer" + '</span></div>' +
            '<div class="excel-import-error" style="display:none"></div>' +
            '<div class="excel-import-summary" style="display:none"></div>' +
            '<div class="excel-import-preview" style="display:none"></div>';
        this.dropzone = this.container.querySelector(".excel-import-dropzone");
        this.fileInput = this.container.querySelector('input[type="file"]');
        this.filebar = this.container.querySelector(".excel-import-filebar");
        this.errorBox = this.container.querySelector(".excel-import-error");
        this.summaryBox = this.container.querySelector(".excel-import-summary");
        this.previewBox = this.container.querySelector(".excel-import-preview");
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
        if (this.fileInput) {
            this.fileInput.value = "";
        }
        this.filebar.classList.remove("active");
        this.filebar.querySelector(".ei-fname").textContent = "";
        this.previewBox.style.display = "none";
        this.previewBox.innerHTML = "";
        this.summaryBox.style.display = "none";
        this.hideError();
    };

    ExcelImport.prototype.onFile = function (file) {
        var self = this;
        this.hideError();

        // Extension check.
        var accepted = (this.config.acceptedFiles || ".xlsx,.xls").split(",").map(function (s) { return s.trim().toLowerCase(); });
        var name = file.name.toLowerCase();
        var okExt = accepted.some(function (ext) { return ext && name.slice(-ext.length) === ext; });
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

        this.filebar.classList.add("active");
        this.filebar.querySelector(".ei-fname").textContent = file.name;

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
            this.clearData();
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

        if (errors.length) {
            this.showError(errors);
            this.clearData();
            this.summary(0, rows.length);
        } else {
            this.hideError();
            this.commit(rows);
            this.summary(rows.length, rows.length);
        }
    };

    /** Write the validated rows (keyed by Excel header) to the hidden field. */
    ExcelImport.prototype.commit = function (rows) {
        if (!this.hiddenInput) { return; }
        var self = this;
        var payload = rows.map(function (row) {
            var obj = {};
            self.headers.forEach(function (h) {
                obj[h] = self.cell(row, h);
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
        badCells = badCells || {};
        rowDup = rowDup || {};
        rowBad = rowBad || {};
        var self = this;
        var maxH = parseInt(this.config.previewHeight, 10) || 400;

        var thead = "<tr><th class='excel-import-rownum'>#</th>" +
            this.headers.map(function (h) { return "<th>" + escapeHtml(h) + "</th>"; }).join("") + "</tr>";

        var tbody = rows.map(function (row, i) {
            var cls = "";
            if (rowBad[i]) { cls = "ei-row-bad"; }
            else if (rowDup[i]) { cls = "ei-row-dup"; }
            var tds = self.headers.map(function (h, idx) {
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
            "<table class='table'><thead>" + thead + "</thead><tbody>" + tbody + "</tbody></table></div>";
    };

    ExcelImport.prototype.showError = function (errors) {
        var html = "";
        if (errors.length === 1) {
            html = escapeHtml(errors[0]);
        } else {
            html = "<strong>" + errors.length + "</strong><ul>" +
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
