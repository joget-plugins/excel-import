<div class="form-cell" ${elementMetaData!}>
    <div class="form-cell-value">
        <#if element.properties.label?? && element.properties.label?has_content>
            <label class="label" field-tooltip="${elementParamName!}">${element.properties.label} <span class="form-cell-validator">${decoration!}</span><#if error??> <span class="form-error-message">${error}</span></#if></label>
        <#elseif error??>
            <label class="label"><span class="form-error-message">${error}</span></label>
        </#if>

        <#if element.properties.readonly! != 'true'>
            <div class="excel-import-widget" id="excel-import-${element.properties.id!}-${element.properties.elementUniqueKey!}"></div>
        </#if>

        <input type="hidden"
               class="excel-import-input"
               name="${elementParamName!}"
               value="${value!?html}"
               <#if element.properties.id??>data-element-id="${element.properties.id!}"</#if> />
    </div>

    <#if element.properties.readonly! != 'true'>
    <script type="text/javascript">
        (function () {
            var base = "${request.contextPath}/plugin/org.joget.plugin.melkart.ExcelParser/js/";

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
