<#assign base = "${request.contextPath}/plugin/org.joget.plugin.melkart.ExcelParser/">
<link rel="stylesheet" type="text/css" href="${base}css/excel-import.css" />

<div class="form-cell<#if element.properties.fullWidth! == 'true'> excel-import-fullwidth</#if>" ${elementMetaData!}>
    <#if element.properties.label?? && element.properties.label?has_content>
        <label class="label" field-tooltip="${elementParamName!}">${element.properties.label} <span class="form-cell-validator">${decoration!}</span><#if error??> <span class="form-error-message">${error}</span></#if></label>
    <#elseif error??>
        <label class="label"><span class="form-error-message">${error}</span></label>
    </#if>

    <#if element.properties.readonly! != 'true'>
        <div class="excel-import-widget form-fileupload" id="excel-import-${element.properties.id!}-${element.properties.elementUniqueKey!}">
            <div class="excel-import-dropzone">
                <span class="ei-icon">&#128228;</span>
                <span class="ei-text">${dropzoneText!"Choisir ou déposer un fichier Excel ici"}</span>
                <span class="ei-hint">.xlsx, .xls</span>
                <input type="file" accept=".xlsx,.xls" style="display:none" />
            </div>
            <div class="excel-import-filebar form-fileupload">
                <span class="ei-fname"></span>
                <span class="excel-import-remove">&#10005; Retirer</span>
            </div>
            <div class="excel-import-error" style="display:none"></div>
            <div class="excel-import-summary" style="display:none"></div>
            <div class="excel-import-preview" style="display:none"></div>
        </div>
    </#if>

    <input type="hidden"
        class="excel-import-input"
        value="${value!?html}"
        name="${elementParamName!}"
        <#if element.properties.id??>data-element-id="${element.properties.id!}"</#if> />

    <#-- Carries the render-time resolved parent value through the submission, so store() does not
         depend on the original form-load URL request parameters (which the submit POST loses). -->
    <input type="hidden"
        class="excel-import-parent"
        value="${resolvedParentValue!?html}"
        name="${parentCarrierName!}" />

    <#if element.properties.readonly! != 'true'>
    <script type="text/javascript">
        (function () {
            var base = "${base}js/";

            // Load a script once, even when several Excel Parser elements are on the page.
            function loadOnce(src, cb) {
                var sel = 'script[data-ei-src="' + src + '"]';
                var existing = document.querySelector(sel);
                if (existing) {
                    if (existing.getAttribute("data-ei-loaded") === "1") { cb(); }
                    else { existing.addEventListener("load", cb); }
                    return;
                }
                var s = document.createElement("script");
                s.src = src;
                s.setAttribute("data-ei-src", src);
                s.onload = function () { s.setAttribute("data-ei-loaded", "1"); cb(); };
                document.head.appendChild(s);
            }

            function ensureLibs(cb) {
                loadOnce(base + "xlsx.full.min.js", function () {
                    loadOnce(base + "excel-import-lib.js", cb);
                });
            }

            var config = ${jsConfig!"{}"};
            config.fieldName = "${elementParamName!}";
            config.containerId = "excel-import-${element.properties.id!}-${element.properties.elementUniqueKey!}";

            var start = function () {
                ensureLibs(function () {
                    if (typeof window.initExcelImport === "function") {
                        window.initExcelImport(config);
                    }
                });
            };

            if (window.jQuery) { jQuery(start); }
            else if (document.readyState !== "loading") { start(); }
            else { document.addEventListener("DOMContentLoaded", start); }
        })();
    </script>
    </#if>
</div>
