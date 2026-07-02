/*
 * Reusable "grid with per-row options popup" custom property-editor type for Joget.
 *
 * Behaves like Joget's native `type:"grid"` (config-driven columns, same stored value shape: an
 * array of row objects keyed by column key) but adds two things the native grid lacks:
 *   1. real checkbox cells (the native grid silently drops `type:"checkbox"` columns), and
 *   2. columns flagged `"popup":"true"`, which are moved out of the inline table and edited in a
 *      per-row modal opened by an Options button — keeping wide grids readable.
 *
 * REUSE IN OTHER PLUGINS: this file is generic (nothing here is excel-import specific). Copy it into
 * your plugin's resources and point a `type:"custom"` property's `script_url` at it, then declare
 * your columns in the property JSON. The property editor fetches this file as TEXT and runs
 * `eval("[" + fileContents + "]")[0]`, so it MUST stay a single bare object literal: no var/const
 * wrapper, no leading statement, and NO trailing semicolon.
 *
 * PROPERTY JSON CONFIG (all resolved server-side, so @@i18n@@ works):
 *   {
 *     "name": "myProp",
 *     "type": "custom",
 *     "script_url": "[CONTEXT_PATH]/plugin/<ClassName>/js/excel-import-columns.js",
 *     "columns": [
 *       { "key": "field",  "label": "Field",  "type": "textfield", "required": "true" },
 *       { "key": "lookup", "label": "Lookup", "type": "autocomplete", "options_ajax": "[CONTEXT_PATH]/..." },
 *       { "key": "op",     "label": "Op",     "type": "select", "options": [ {"value":"eq","label":"="} ] },
 *       { "key": "active", "label": "Active", "type": "checkbox", "true_value": "true", "false_value": "" },
 *       { "key": "note",   "label": "Note",   "type": "textfield", "popup": "true" },
 *       { "key": "id",     "label": "Id",     "type": "readonly" }
 *     ],
 *     "labels": { "options": "...", "optionsTitle": "...", "addRow": "...", "done": "...",
 *                 "moveUp": "...", "moveDown": "...", "delete": "...", "noRows": "..." }
 *   }
 *
 * Column attributes: key (required), label, type, required, readonly, popup, options,
 * options_ajax, true_value, false_value, placeholder.
 * Supported types: textfield/text (default), number, password, readonly, select/selectbox,
 * checkbox/truefalse, autocomplete. A column with `options`/`options_ajax` and no explicit type
 * is treated as a select.
 *
 * DEPENDENT COLUMNS (Joget's dependency-field attribute syntax, evaluated per row against a
 * sibling column in the SAME row):
 *   control_field     : key of another column in the row that controls this one's visibility.
 *   control_value     : value (or regex) the controlling cell must match for this column to show.
 *   control_use_regex : "true" to match control_value as a regex; otherwise exact-equality.
 * A column hidden by a non-matching control is removed from the row (its value is cleared on
 * save and its "required" flag is skipped), mirroring Joget's native dependent-field behaviour.
 * Works for inline cells and popup fields, including a popup field controlled by an inline cell
 * (or vice-versa).
 *
 * Scope available on `this`: id, value, properties (incl. properties.columns / .labels),
 * options.contextPath, properties.appPath. Globals: jQuery, jQuery UI (autocomplete),
 * PropertyEditor.Util (escapeHtmlTag, replaceContextPath).
 */
{
    /* CSS is injected once here; the runtime <link> in ExcelImport.ftl is not present in the builder. */
    initialize: function () {
        if (!document.querySelector('link[data-ei-cols-css="1"]')) {
            var href = this.options.contextPath
                + "/plugin/org.joget.plugin.melkart.ExcelImport/css/excel-import-columns.css";
            var link = document.createElement("link");
            link.rel = "stylesheet";
            link.type = "text/css";
            link.href = href;
            link.setAttribute("data-ei-cols-css", "1");
            document.head.appendChild(link);
        }
    },

    _esc: function (v) {
        return PropertyEditor.Util.escapeHtmlTag(v === undefined || v === null ? "" : String(v));
    },

    /* Chrome labels (widget buttons / modal), server-resolved with English fallbacks. */
    _lbl: function () {
        var L = this.properties.labels || {};
        var d = function (v, f) { return (v === undefined || v === null || v === "") ? f : v; };
        return {
            options: d(L.options, "Options"),
            optionsTitle: d(L.optionsTitle, "Options"),
            addRow: d(L.addRow, "Add a row"),
            done: d(L.done, "Done"),
            moveUp: d(L.moveUp, "Move up"),
            moveDown: d(L.moveDown, "Move down"),
            "delete": d(L["delete"], "Delete"),
            noRows: d(L.noRows, "No row yet. Click “" + d(L.addRow, "Add a row") + "” to begin."),
            requiredSymbol: d(L.requiredSymbol, "*")
        };
    },

    /* Normalised column list from the property config, memoised. */
    _cols: function () {
        if (!this.__cols) {
            var raw = $.isArray(this.properties.columns) ? this.properties.columns : [];
            this.__cols = [];
            this.__colByKey = {};
            for (var i = 0; i < raw.length; i++) {
                if (raw[i] && raw[i].key) {
                    this.__cols.push(raw[i]);
                    this.__colByKey[raw[i].key] = raw[i];
                }
            }
        }
        return this.__cols;
    },
    _colByKey: function () { this._cols(); return this.__colByKey; },
    _inlineCols: function () { return $.grep(this._cols(), function (c) { return String(c.popup) !== "true"; }); },
    _popupCols: function () { return $.grep(this._cols(), function (c) { return String(c.popup) === "true"; }); },

    _type: function (col) {
        var t = (col.type || "").toLowerCase();
        if (t === "checkbox" || t === "truefalse") { return "checkbox"; }
        if (t === "autocomplete") { return "autocomplete"; }
        if (t === "number") { return "number"; }
        if (t === "password") { return "password"; }
        if (t === "select" || t === "selectbox") { return "select"; }
        if (t === "readonly") { return "readonly"; }
        if (col.options !== undefined || col.options_ajax !== undefined) { return "select"; }
        return "text";
    },
    _trueVal: function (col) { return col.true_value !== undefined ? col.true_value : "true"; },
    _falseVal: function (col) { return col.false_value !== undefined ? col.false_value : ""; },

    /* True when `col` has no controlling field, or the controlling cell's value matches. */
    _matchesControl: function (col, controllingValue) {
        var cf = col.control_field;
        if (cf === undefined || cf === null || cf === "") { return true; }
        var target = (col.control_value === undefined || col.control_value === null) ? "" : String(col.control_value);
        var val = (controllingValue === undefined || controllingValue === null) ? "" : String(controllingValue);
        if (String(col.control_use_regex) === "true") {
            try { return new RegExp(target).test(val); } catch (e) { return true; }
        }
        return val === target;
    },

    /* Reads every column's current value from a row's <tr> (inline controls + popup hidden inputs). */
    _rowObj: function ($tr) {
        var self = this, obj = {};
        $.each(this._cols(), function (i, col) { obj[col.key] = self._readCell($tr, col); });
        return obj;
    },

    /* Tolerant read of the stored value (native array, or JSON string as insurance). */
    _rows: function () {
        var v = this.value;
        if (typeof v === "string") { try { v = JSON.parse(v); } catch (e) { v = []; } }
        return $.isArray(v) ? v : [];
    },

    /* Renders a single cell/field control for a column, named by its key. */
    renderControl: function (col, value) {
        var esc = this._esc, t = this._type(col);
        var ro = String(col.readonly) === "true";
        var key = esc(col.key);
        value = (value === undefined || value === null) ? "" : value;

        if (t === "checkbox") {
            var tv = this._trueVal(col);
            var checked = (String(value) === String(tv)) ? " checked" : "";
            var dis = ro ? " disabled" : "";
            return '<label class="cg-check"><input type="checkbox" name="' + key + '" value="'
                + esc(tv) + '"' + checked + dis + '/><span></span></label>';
        }

        if (t === "select") {
            var dis2 = ro ? " disabled" : "";
            var ajax = col.options_ajax ? ' data-ajax="' + esc(col.options_ajax) + '"' : '';
            var html = '<select name="' + key + '" data-value="' + esc(value) + '"' + dis2 + ajax + '>';
            var opts = $.isArray(col.options) ? col.options : [];
            var found = false, inner = '';
            $.each(opts, function (i, o) {
                var sel = (String(value) === String(o.value)) ? ' selected' : '';
                if (sel) { found = true; }
                inner += '<option value="' + esc(o.value) + '"' + sel + '>' + esc(o.label) + '</option>';
            });
            // Keep an unknown current value selectable until an options_ajax fetch replaces the list.
            if (!found && value !== "") {
                inner = '<option value="' + esc(value) + '" selected>' + esc(value) + '</option>' + inner;
            }
            return html + inner + '</select>';
        }

        if (t === "autocomplete") {
            var ajax2 = col.options_ajax ? ' data-ajax="' + esc(col.options_ajax) + '"' : '';
            var ph = col.placeholder ? ' placeholder="' + esc(col.placeholder) + '"' : '';
            return '<input type="text" name="' + key + '" class="cg-autocomplete" value="'
                + esc(value) + '"' + ajax2 + ph + (ro ? ' readonly' : '') + '/>';
        }

        var inputType = (t === "number") ? "number" : (t === "password") ? "password" : "text";
        var ph2 = col.placeholder ? ' placeholder="' + esc(col.placeholder) + '"' : '';
        var roAttr = (ro || t === "readonly") ? ' readonly' : '';
        return '<input type="' + inputType + '" name="' + key + '" value="' + esc(value) + '"' + ph2 + roAttr + '/>';
    },

    _labelHtml: function (col) {
        var L = this._lbl();
        var req = (String(col.required) === "true")
            ? ' <span class="cg-required">' + this._esc(L.requiredSymbol) + '</span>' : '';
        return this._esc(col.label || col.key) + req;
    },

    renderRow: function (row) {
        var self = this, esc = this._esc;
        row = row || {};
        var inline = this._inlineCols(), popup = this._popupCols();
        var html = '<tr class="ei-cols-row">';
        $.each(inline, function (i, col) {
            var cls = (self._type(col) === "checkbox") ? ' ei-cols-flag' : '';
            html += '<td class="ei-cols-cell' + cls + '" data-col="' + esc(col.key)
                + '" data-th="' + esc(col.label || col.key) + '">'
                + self.renderControl(col, row[col.key]) + '</td>';
        });
        // Popup columns ride along as hidden inputs so getData reads every column uniformly.
        $.each(popup, function (i, col) {
            html += '<input type="hidden" name="' + esc(col.key) + '" value="'
                + esc(row[col.key] === undefined ? "" : row[col.key]) + '"/>';
        });
        html += '<td class="ei-cols-actions">';
        if (popup.length) {
            html += '<a href="#" class="ei-cols-opts" title="' + esc(this._lbl().options) + '"><i class="fas fa-sliders-h"></i></a>';
        }
        html += '<a href="#" class="ei-cols-up" title="' + esc(this._lbl().moveUp) + '"><i class="fas fa-chevron-up"></i></a>';
        html += '<a href="#" class="ei-cols-down" title="' + esc(this._lbl().moveDown) + '"><i class="fas fa-chevron-down"></i></a>';
        html += '<a href="#" class="ei-cols-del" title="' + esc(this._lbl()["delete"]) + '"><i class="fas fa-trash-alt"></i></a>';
        html += '</td></tr>';
        return html;
    },

    renderField: function () {
        var self = this, L = this._lbl(), rows = this._rows();
        var inline = this._inlineCols();
        var html = '<div class="ei-cols-widget" id="' + this.id + '_widget">';
        html += '<table class="ei-cols-table" id="' + this.id + '"><thead><tr>';
        $.each(inline, function (i, col) {
            var cls = (self._type(col) === "checkbox") ? ' class="ei-cols-flag"' : '';
            html += '<th' + cls + '>' + self._labelHtml(col) + '</th>';
        });
        html += '<th class="ei-cols-actions"></th></tr></thead><tbody>';
        for (var i = 0; i < rows.length; i++) { html += this.renderRow(rows[i]); }
        html += '</tbody></table>';
        html += '<div class="ei-cols-empty"' + (rows.length ? ' style="display:none"' : '') + '>' + this._esc(L.noRows) + '</div>';
        html += '<a href="#" class="ei-cols-add"><i class="fas fa-plus-circle"></i> ' + this._esc(L.addRow) + '</a>';
        html += '</div>';
        return html;
    },

    /* Reads one row's value for a column, honouring checkbox checked-state. */
    _readCell: function ($scope, col) {
        var $el = $scope.find('[name="' + col.key + '"]');
        if (this._type(col) === "checkbox" && $el.is(":checkbox")) {
            return $el.is(":checked") ? this._trueVal(col) : this._falseVal(col);
        }
        var v = $el.val();
        return (v === null || v === undefined) ? "" : ("" + v).trim();
    },

    getData: function (useDefault) {
        var self = this, data = {};
        if (this.isDataReady && $('#' + this.id).length > 0) {
            var cols = this._cols(), arr = [];
            $('#' + this.id + ' tbody tr.ei-cols-row').each(function () {
                var $tr = $(this), obj = {};
                $.each(cols, function (i, col) { obj[col.key] = self._readCell($tr, col); });
                // A column hidden by its controlling field is ignored on save (Joget behaviour).
                $.each(cols, function (i, col) {
                    if (col.control_field && !self._matchesControl(col, obj[col.control_field])) { obj[col.key] = ""; }
                });
                arr.push(obj);
            });
            if (arr.length === 0 && useDefault && this.defaultValue !== null && this.defaultValue !== undefined) {
                arr = this.defaultValue;
            }
            data[this.properties.name] = arr;
        } else {
            data[this.properties.name] = this.value;
        }
        return data;
    },

    /* Marks required inline cells that are empty, mirroring the native grid. */
    addOnValidation: function (data, errors, checkEncryption) {
        var self = this, wrapper = $('#' + this.id + '_input');
        var table = $('#' + this.id);
        table.find('td').removeClass('cg-error');
        var value = data[this.properties.name];
        var hasError = false;
        if ($.isArray(value)) {
            var inline = this._inlineCols();
            $.each(value, function (i, row) {
                $.each(inline, function (j, col) {
                    if (String(col.required) === "true"
                            && self._matchesControl(col, row[col.control_field])) {
                        var v = row[col.key];
                        if (v === undefined || v === null || v === "") {
                            table.find('tr.ei-cols-row:eq(' + i + ') td:eq(' + j + ')').addClass('cg-error');
                            hasError = true;
                        }
                    }
                });
            });
        }
        if (hasError) {
            var obj = { field: this.properties.name, fieldName: this.properties.label, message: this.options.mandatoryMessage };
            errors.push(obj);
            wrapper.append('<div class="property-input-error">' + obj.message + '</div>');
        }
    },

    initScripting: function () {
        var self = this;
        var $widget = $('#' + this.id + '_widget');
        var change = function () { $('#' + self.id).trigger("change"); };
        var refreshEmpty = function () {
            $widget.find('.ei-cols-empty').toggle($('#' + self.id + ' tbody tr.ei-cols-row').length === 0);
        };

        $widget.off("click.eicols");

        $widget.on("click.eicols", ".ei-cols-add", function (e) {
            e.preventDefault();
            var $row = $(self.renderRow({}));
            $('#' + self.id + ' tbody').append($row);
            self._initFieldSources($row);
            self._applyRowDeps($row);
            refreshEmpty();
            change();
        });

        $widget.on("click.eicols", ".ei-cols-del", function (e) {
            e.preventDefault();
            $(this).closest("tr.ei-cols-row").remove();
            refreshEmpty();
            change();
        });

        $widget.on("click.eicols", ".ei-cols-up", function (e) {
            e.preventDefault();
            var $tr = $(this).closest("tr.ei-cols-row"), $prev = $tr.prev("tr.ei-cols-row");
            if ($prev.length) { $tr.insertBefore($prev); change(); }
        });

        $widget.on("click.eicols", ".ei-cols-down", function (e) {
            e.preventDefault();
            var $tr = $(this).closest("tr.ei-cols-row"), $next = $tr.next("tr.ei-cols-row");
            if ($next.length) { $tr.insertAfter($next); change(); }
        });

        $widget.on("change.eicols", "input, select", function () {
            var $tr = $(this).closest("tr.ei-cols-row");
            if ($tr.length) { self._applyRowDeps($tr); }
            change();
        });

        $widget.on("click.eicols", ".ei-cols-opts", function (e) {
            e.preventDefault();
            self.openOptions($(this).closest("tr.ei-cols-row"));
        });

        this._initFieldSources($widget);
        $('#' + this.id + ' tbody tr.ei-cols-row').each(function () { self._applyRowDeps($(this)); });
    },

    /* Wires up options_ajax selects and (static or ajax) autocomplete inputs within a scope. */
    _initFieldSources: function (scope) {
        var self = this, esc = this._esc, byKey = this._colByKey();

        $(scope).find('select[data-ajax]').each(function () {
            var $sel = $(this), url = $sel.attr('data-ajax'), cur = $sel.attr('data-value') || $sel.val();
            self._fetchOptions(url).done(function (opts) {
                var html = '', found = false;
                $.each(opts, function (i, o) {
                    var sel = (String(cur) === String(o.value)) ? ' selected' : '';
                    if (sel) { found = true; }
                    html += '<option value="' + esc(o.value) + '"' + sel + '>' + esc(o.label) + '</option>';
                });
                if (!found && cur) { html = '<option value="' + esc(cur) + '" selected>' + esc(cur) + '</option>' + html; }
                $sel.html(html);
            });
        });

        $(scope).find('input.cg-autocomplete').each(function () {
            var $inp = $(this);
            if (!$inp.autocomplete) { return; }
            var ajaxUrl = $inp.attr('data-ajax');
            var opts = { minLength: 0, open: function () { $(this).autocomplete('widget').css('z-index', 100001); } };
            if (ajaxUrl) {
                self._fetchOptions(ajaxUrl).done(function (list) { $inp.autocomplete($.extend({ source: list }, opts)); });
            } else {
                var col = byKey[$inp.attr('name')];
                opts.source = (col && $.isArray(col.options)) ? col.options : [];
                $inp.autocomplete(opts);
            }
        });
    },

    /* Fetches an options endpoint once per URL (cached), returning a promise of [{value,label}]. */
    _fetchOptions: function (url) {
        this.__ajaxCache = this.__ajaxCache || {};
        if (this.__ajaxCache[url]) { return this.__ajaxCache[url]; }
        var real = PropertyEditor.Util.replaceContextPath(url, this.options.contextPath)
            .replace(/\[APP_PATH\]/g, this.properties.appPath || "");
        var p = $.ajax({ url: real, dataType: "json" }).then(function (resp) {
            if ($.isArray(resp)) { return resp; }
            if (resp && $.isArray(resp.data)) { return resp.data; }
            return [];
        }, function () { return []; });
        this.__ajaxCache[url] = p;
        return p;
    },

    /* Shows/hides inline dependent cells in a row based on their controlling sibling cell. */
    _applyRowDeps: function ($tr) {
        var self = this, obj = this._rowObj($tr);
        $.each(this._inlineCols(), function (i, col) {
            if (col.control_field) {
                var vis = self._matchesControl(col, obj[col.control_field]);
                $tr.find('td.ei-cols-cell[data-col="' + col.key + '"]').toggleClass("ei-cols-dep-hidden", !vis);
            }
        });
    },

    /* Shows/hides popup dependent fields in an open modal. Controlling values come from the modal
       for popup columns and from the row's <tr> for inline columns (cross-scope dependencies). */
    _applyModalDeps: function ($modal, $tr) {
        var self = this, obj = {};
        $.each(this._cols(), function (i, col) {
            obj[col.key] = self._readCell(String(col.popup) === "true" ? $modal : $tr, col);
        });
        $.each(this._popupCols(), function (i, col) {
            if (col.control_field) {
                $modal.find('.ei-modal-field[data-col="' + col.key + '"]')
                    .toggle(self._matchesControl(col, obj[col.control_field]));
            }
        });
    },

    openOptions: function ($tr) {
        var self = this, esc = this._esc, L = this._lbl();
        var popup = this._popupCols();
        if (!popup.length) { return; }

        var body = '';
        $.each(popup, function (i, col) {
            var cur = $tr.find('[name="' + col.key + '"]').val() || "";
            body += '<div class="ei-modal-field" data-col="' + esc(col.key) + '"><label>' + self._labelHtml(col) + '</label>'
                + self.renderControl(col, cur) + '</div>';
        });

        var firstInline = this._inlineCols()[0];
        var header = firstInline ? ($tr.find('[name="' + firstInline.key + '"]').val() || "") : "";
        var title = esc(L.optionsTitle) + (header ? ' — ' + esc(header) : '');

        var $overlay = $('<div class="ei-modal-overlay"></div>');
        var $modal = $(
            '<div class="ei-modal" role="dialog" aria-modal="true">'
            + '<div class="ei-modal-header"><span class="ei-modal-title">' + title + '</span>'
            + '<a href="#" class="ei-modal-close" aria-label="close">&times;</a></div>'
            + '<div class="ei-modal-body">' + body + '</div>'
            + '<div class="ei-modal-footer"><a href="#" class="ei-modal-done">' + esc(L.done) + '</a></div>'
            + '</div>'
        );
        $overlay.append($modal);
        $('body').append($overlay);
        this._initFieldSources($modal);
        this._applyModalDeps($modal, $tr);

        var close = function () { $(document).off("keydown.eimodal"); $overlay.remove(); };
        var commit = function () {
            $.each(popup, function (i, col) {
                $tr.find('[name="' + col.key + '"]').val(self._readCell($modal, col));
            });
            self._applyRowDeps($tr);      // a popup control may drive an inline dependent
            $('#' + self.id).trigger("change");
            close();
        };

        $modal.on("change", "input, select", function () { self._applyModalDeps($modal, $tr); });
        $overlay.on("click", function (e) { if (e.target === $overlay[0]) { close(); } });
        $modal.find(".ei-modal-close").on("click", function (e) { e.preventDefault(); close(); });
        $modal.find(".ei-modal-done").on("click", function (e) { e.preventDefault(); commit(); });
        $(document).on("keydown.eimodal", function (e) { if (e.key === "Escape") { close(); } });
        $modal.find("input, select").first().trigger("focus");
    }
}
