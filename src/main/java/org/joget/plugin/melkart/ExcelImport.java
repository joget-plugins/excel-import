package org.joget.plugin.melkart;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;

import org.joget.apps.app.model.AppDefinition;
import org.joget.apps.app.service.AppPluginUtil;
import org.joget.apps.app.service.AppService;
import org.joget.apps.app.service.AppUtil;
import org.joget.apps.form.dao.FormDataDao;
import org.joget.apps.form.model.Element;
import org.joget.apps.form.model.Form;
import org.joget.apps.form.model.FormBinder;
import org.joget.apps.form.model.FormBuilderPalette;
import org.joget.apps.form.model.FormBuilderPaletteElement;
import org.joget.apps.form.model.FormContainer;
import org.joget.apps.form.model.FormData;
import org.joget.apps.form.model.FormLoadBinder;
import org.joget.apps.form.model.FormRow;
import org.joget.apps.form.model.FormRowSet;
import org.joget.apps.form.model.FormStoreBinder;
import org.joget.apps.form.service.FormUtil;
import org.joget.commons.util.LogUtil;
import org.json.JSONArray;
import org.json.JSONObject;

/**
 * Excel Import form element.
 *
 * <p>Renders a self-contained Excel import widget (own drop zone + client side preview,
 * parsed with a bundled copy of SheetJS/xlsx). The validated rows are written, keyed by
 * Excel header, into a hidden field as a JSON array.</p>
 *
 * <p>Persistence is delegated to a companion {@link ExcelImportBinder} (a real
 * {@link FormBinder}) which this element supplies through {@link #getStoreBinder()} /
 * {@link #getLoadBinder()}: on submit every Excel row is stored as one record in a configurable
 * target form/table (multi-row), each linked to the current submission via a configurable parent
 * foreign-key column. The binder is a separate class because Joget's {@code Element} and
 * {@code FormBinder} are distinct (sibling) types and a single class cannot be both. The
 * front-end validation (required headers, required cells across columns, composite duplicate
 * key across columns) is mirrored server-side in {@link #selfValidate(FormData)} and blocks
 * the submission on any error.</p>
 */
public class ExcelImport extends Element implements FormBuilderPaletteElement, FormContainer {

    public static final String MESSAGES_PATH = "messages/ExcelImport";

    @Override
    public String getClassName() {
        return getClass().getName();
    }

    @Override
    public String getName() {
        return "Excel Import";
    }

    @Override
    public String getLabel() {
        return AppPluginUtil.getMessage("ExcelImport.pluginLabel", getClassName(), MESSAGES_PATH);
    }

    @Override
    public String getDescription() {
        return AppPluginUtil.getMessage("ExcelImport.pluginDesc", getClassName(), MESSAGES_PATH);
    }

    @Override
    public String getVersion() {
        try {
            Properties props = new Properties();
            try (InputStream is = getClass().getResourceAsStream(
                    "/META-INF/maven/org.joget.plugin.melkart/excel-import/pom.properties")) {
                if (is != null) {
                    props.load(is);
                    return props.getProperty("version");
                }
            }
        } catch (IOException e) {
            LogUtil.warn(getClassName(), "Error getting version from pom.properties file : " + e.getMessage());
        }
        return "1.0-SNAPSHOT";
    }

    @Override
    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(), "/properties/ExcelImport.json", null, true, MESSAGES_PATH);
    }

    //
    // ---- Form builder palette ----
    //

    @Override
    public String getFormBuilderTemplate() {
        return "<div style='border: 2px dashed #667eea;padding: 20px;text-align: center;border-radius: 10px;color:#667eea;'>"
                + "<i class=\"fas fa-file-excel\"></i> Excel Import</div>";
    }

    @Override
    public String getFormBuilderCategory() {
        return FormBuilderPalette.CATEGORY_CUSTOM;
    }

    @Override
    public String getFormBuilderIcon() {
        return "<i class=\"fas fa-file-excel\"></i>";
    }

    @Override
    public int getFormBuilderPosition() {
        return 100;
    }

    //
    // ---- Supply the companion binder (a real FormBinder) ----
    //
    // The element cannot be its own binder: Element and FormBinder are sibling types, so the
    // runtime cast "(FormBinder) element.getStoreBinder()" would throw ClassCastException.
    // Instead we hand back a cached ExcelImportBinder instance. It is cached (via the inherited
    // store/load binder fields) because Joget keys the formatData() row set by the exact binder
    // instance returned here and later looks it up again to invoke store().

    @Override
    public FormStoreBinder getStoreBinder() {
        FormStoreBinder binder = super.getStoreBinder();
        if (binder == null) {
            ExcelImportBinder excelBinder = new ExcelImportBinder();
            excelBinder.setElement(this);
            setStoreBinder(excelBinder);
            binder = excelBinder;
        }
        return binder;
    }

    @Override
    public FormLoadBinder getLoadBinder() {
        FormLoadBinder binder = super.getLoadBinder();
        if (binder == null) {
            ExcelImportBinder excelBinder = new ExcelImportBinder();
            excelBinder.setElement(this);
            setLoadBinder(excelBinder);
            binder = excelBinder;
        }
        return binder;
    }

    //
    // ---- Rendering ----
    //

    @Override
    public String renderTemplate(FormData formData, Map dataModel) {
        String template = "ExcelImport.ftl";

        String parentValue = resolveParentValue(formData);

        // Current value: prefer a submitted (or failed-validation reload) value, otherwise
        // rebuild it from the existing child records so an edit/no-change re-submit is a no-op.
        String value = FormUtil.getElementPropertyValue(this, formData);
        if ((value == null || value.trim().isEmpty())) {
            String primaryKey = formData != null ? formData.getPrimaryKeyValue() : null;
            if (primaryKey != null && !primaryKey.isEmpty()) {
                value = loadExistingAsJson(primaryKey, formData);
            }
        }
        if (value == null) {
            value = "";
        }

        dataModel.put("value", value);
        dataModel.put("parentCarrierName", getParentCarrierName());
        dataModel.put("resolvedParentValue", parentValue != null ? parentValue : "");
        dataModel.put("fileNameCarrierName", getFileNameCarrierName());
        dataModel.put("fileName", resolveFileName(formData));
        dataModel.put("jsConfig", buildClientConfig());
        dataModel.put("dropzoneText", defaultStr(getPropertyString("dropzoneText"),
                AppPluginUtil.getMessage("ExcelImport.dropzoneTextDefault", getClassName(), MESSAGES_PATH)));

        return FormUtil.generateElementHtml(this, formData, template, dataModel);
    }

    /**
     * Builds the JSON config consumed by excel-import-lib.js.
     */
    protected String buildClientConfig() {
        JSONObject cfg = new JSONObject();
        try {
            List<Map<String, String>> columns = getColumns();

            JSONArray headers = new JSONArray();
            JSONArray previewHeaders = new JSONArray();
            JSONObject mapping = new JSONObject();
            JSONArray requiredColumns = new JSONArray();
            JSONArray uniqueColumns = new JSONArray();
            JSONObject coercion = new JSONObject();
            JSONObject validation = new JSONObject();
            for (Map<String, String> col : columns) {
                String header = col.get("excelHeader");
                headers.put(header);
                if (!"true".equalsIgnoreCase(col.get("hideInPreview"))) {
                    previewHeaders.put(header);
                }
                mapping.put(header, col.get("fieldId"));
                if ("true".equalsIgnoreCase(col.get("required"))) {
                    requiredColumns.put(header);
                }
                if ("true".equalsIgnoreCase(col.get("uniqueKey"))) {
                    uniqueColumns.put(header);
                }
                JSONObject rule = coercionRule(col);
                if (rule != null) {
                    coercion.put(header, rule);
                }
                JSONObject vrule = validationRule(col);
                if (vrule != null) {
                    validation.put(header, vrule);
                }
            }

            cfg.put("elementId", getPropertyString("id"));
            cfg.put("headers", headers);
            cfg.put("previewHeaders", previewHeaders);
            cfg.put("mapping", mapping);
            cfg.put("requiredColumns", requiredColumns);
            cfg.put("uniqueColumns", uniqueColumns);
            cfg.put("coercion", coercion);
            cfg.put("validation", validation);
            cfg.put("caseSensitive", "true".equalsIgnoreCase(getPropertyString("caseSensitiveHeaders")));
            // CSV parsing overrides (ignored for .xlsx/.xls). Empty delimiter = auto-detect;
            // empty quote = the default double-quote.
            cfg.put("csvDelimiter", getPropertyStringOrDefault("csvDelimiter", ""));
            cfg.put("csvQuote", getPropertyStringOrDefault("csvQuote", ""));
            cfg.put("required", "true".equalsIgnoreCase(getPropertyString("required")));
            cfg.put("maxFileSizeMB", parseDouble(getPropertyString("maxFileSizeMB"), 5));
            cfg.put("previewHeight", defaultStr(getPropertyString("previewHeight"), "400"));
            cfg.put("showPreview", !"true".equalsIgnoreCase(getPropertyString("hidePreview")));

            // Localised messages shared with the client (overridable per element).
            JSONObject messages = new JSONObject();
            messages.put("missingHeaders", customMsg("msgMissingHeaders", "ExcelImport.err.missingHeaders"));
            messages.put("requiredCell", customMsg("msgRequiredCell", "ExcelImport.err.requiredCell"));
            messages.put("duplicate", customMsg("msgDuplicate", "ExcelImport.err.duplicate"));
            messages.put("pattern", customMsg("msgPattern", "ExcelImport.err.pattern"));
            messages.put("range", customMsg("msgRange", "ExcelImport.err.range"));
            messages.put("notAllowed", customMsg("msgNotAllowed", "ExcelImport.err.notAllowed"));
            messages.put("emptyFile", customMsg("msgEmptyFile", "ExcelImport.err.emptyFile"));
            messages.put("readError", customMsg("msgReadError", "ExcelImport.err.readError"));
            messages.put("fileTooLarge", customMsg("msgFileTooLarge", "ExcelImport.err.fileTooLarge"));
            messages.put("rowsValid", customMsg("msgRowsValid", "ExcelImport.info.rowsValid"));
            cfg.put("messages", messages);
        } catch (Exception e) {
            LogUtil.error(getClassName(), e, "Error building client config");
        }
        return cfg.toString();
    }

    //
    // ---- Element data (rows handed to the store binder) ----
    //

    @Override
    public FormRowSet formatData(FormData formData) {
        FormRowSet rowSet = new FormRowSet();
        rowSet.setMultiRow(true);

        String value = FormUtil.getElementPropertyValue(this, formData);
        if (value == null || value.trim().isEmpty()) {
            return rowSet;
        }

        List<Map<String, String>> columns = getColumns();
        boolean caseSensitive = "true".equalsIgnoreCase(getPropertyString("caseSensitiveHeaders"));

        try {
            JSONArray arr = new JSONArray(value);
            for (int i = 0; i < arr.length(); i++) {
                JSONObject obj = arr.getJSONObject(i);
                FormRow row = new FormRow();
                for (Map<String, String> col : columns) {
                    String header = col.get("excelHeader");
                    String fieldId = col.get("fieldId");
                    if (fieldId == null || fieldId.isEmpty()) {
                        continue;
                    }
                    row.setProperty(fieldId, coerce(getCell(obj, header, caseSensitive), col));
                }
                rowSet.add(row);
            }
        } catch (Exception e) {
            LogUtil.error(getClassName(), e, "Error parsing imported Excel JSON");
        }
        return rowSet;
    }

    //
    // ---- Child-record helpers (shared by the binder, validation and rendering) ----
    //

    /**
     * Rebuilds the hidden-field JSON (keyed by Excel header) from the stored child records,
     * so the preview is restored on edit and an unchanged re-submit re-stores the same data.
     */
    protected String loadExistingAsJson(String primaryKey, FormData formData) {
        try {
            FormRowSet existing = findChildRows(primaryKey, formData);
            if (existing == null || existing.isEmpty()) {
                return "";
            }
            List<Map<String, String>> columns = getColumns();
            JSONArray arr = new JSONArray();
            for (FormRow row : existing) {
                JSONObject obj = new JSONObject();
                for (Map<String, String> col : columns) {
                    String header = col.get("excelHeader");
                    String fieldId = col.get("fieldId");
                    String cell = row.getProperty(fieldId);
                    obj.put(header, cell != null ? cell : "");
                }
                arr.put(obj);
            }
            return arr.toString();
        } catch (Exception e) {
            LogUtil.error(getClassName(), e, "Error rebuilding existing Excel rows");
            return "";
        }
    }

    /**
     * Finds the rows belonging to the current submission, scoped by the configured parent
     * column/value. Returns {@code null} when no parent link is configured (the preview cannot
     * be scoped to a single submission without one).
     */
    protected FormRowSet findChildRows(String primaryKey, FormData formData) {
        StorageTarget target = resolveStorageTarget();
        if (target == null) {
            return null;
        }
        String parentColumn = getParentColumn();
        if (parentColumn == null) {
            return null;
        }
        String parentValue = resolveParentValue(formData);
        if (parentValue == null || parentValue.isEmpty()) {
            parentValue = primaryKey;
        }
        if (parentValue == null || parentValue.isEmpty()) {
            return null;
        }
        FormDataDao formDataDao = (FormDataDao) AppUtil.getApplicationContext().getBean("formDataDao");
        return formDataDao.find(target.formDefId, target.tableName,
                " WHERE e.customProperties." + parentColumn + " = ?",
                new Object[]{parentValue}, null, null, null, null);
    }

    //
    // ---- Server-side validation (mirrors the front-end, blocks submit) ----
    //

    @Override
    public Boolean selfValidate(FormData formData) {
        String fieldId = getPropertyString("id");
        String value = FormUtil.getElementPropertyValue(this, formData);
        boolean required = "true".equalsIgnoreCase(getPropertyString("required"));

        boolean empty = (value == null || value.trim().isEmpty());
        if (empty) {
            if (required) {
                formData.addFormError(fieldId, customMsg("msgRequired", "ExcelImport.err.required"));
                return false;
            }
            return true;
        }

        // Beyond "required", the back-end reports a single generic "invalid data" error for any
        // failed check (bad JSON, missing headers, empty required cells, duplicates). The detailed,
        // per-row breakdown is surfaced client-side by excel-import-lib.js.
        String invalidMsg = customMsg("msgInvalidData", "ExcelImport.err.invalidData");

        List<Map<String, String>> columns = getColumns();
        boolean caseSensitive = "true".equalsIgnoreCase(getPropertyString("caseSensitiveHeaders"));

        JSONArray arr;
        try {
            arr = new JSONArray(value);
        } catch (Exception e) {
            formData.addFormError(fieldId, invalidMsg);
            return false;
        }

        if (arr.length() == 0) {
            if (required) {
                formData.addFormError(fieldId, customMsg("msgRequired", "ExcelImport.err.required"));
                return false;
            }
            return true;
        }

        // 1. Header structure.
        JSONObject first = arr.optJSONObject(0);
        for (Map<String, String> col : columns) {
            if (!hasKey(first, col.get("excelHeader"), caseSensitive)) {
                formData.addFormError(fieldId, invalidMsg);
                return false;
            }
        }

        // 2. Required cells across the flagged columns.
        for (Map<String, String> col : columns) {
            if (!"true".equalsIgnoreCase(col.get("required"))) {
                continue;
            }
            String header = col.get("excelHeader");
            for (int i = 0; i < arr.length(); i++) {
                String cell = coerce(getCell(arr.optJSONObject(i), header, caseSensitive), col);
                if (cell == null || cell.trim().isEmpty()) {
                    formData.addFormError(fieldId, invalidMsg);
                    return false;
                }
            }
        }

        // 2b. Per-column value rules: allowed values, min/max range, and regex pattern. Only
        // non-empty cells are checked (emptiness is governed by the required check above); mirrors
        // the client-side rules in excel-import-lib.js.
        for (Map<String, String> col : columns) {
            if (!hasCellRules(col)) {
                continue;
            }
            String header = col.get("excelHeader");
            for (int i = 0; i < arr.length(); i++) {
                String cell = coerce(getCell(arr.optJSONObject(i), header, caseSensitive), col);
                if (cell == null || cell.trim().isEmpty()) {
                    continue;
                }
                if (validateCell(cell, col, caseSensitive) != null) {
                    formData.addFormError(fieldId, invalidMsg);
                    return false;
                }
            }
        }

        // 3. Composite duplicate key across the flagged columns (within the file).
        List<Map<String, String>> uniqueCols = new ArrayList<Map<String, String>>();
        for (Map<String, String> col : columns) {
            if ("true".equalsIgnoreCase(col.get("uniqueKey"))) {
                uniqueCols.add(col);
            }
        }
        if (!uniqueCols.isEmpty()) {
            Set<String> seen = new HashSet<String>();
            for (int i = 0; i < arr.length(); i++) {
                String key = compositeKey(arr.optJSONObject(i), uniqueCols, caseSensitive);
                if (!seen.add(key)) {
                    formData.addFormError(fieldId, invalidMsg);
                    return false;
                }
            }

            // 4. Duplicate against existing records in the target table (optional).
            if ("true".equalsIgnoreCase(getPropertyString("checkExistingDuplicates"))) {
                List<String> existingDupRows = findExistingDuplicates(arr, columns, caseSensitive, formData);
                if (existingDupRows == null || !existingDupRows.isEmpty()) {
                    formData.addFormError(fieldId, invalidMsg);
                    return false;
                }
            }
        }

        return true;
    }

    /**
     * @return list of 1-based row numbers that already exist in the target table, or null on error.
     */
    protected List<String> findExistingDuplicates(JSONArray arr,
            List<Map<String, String>> columns, boolean caseSensitive, FormData formData) {
        StorageTarget target = resolveStorageTarget();
        if (target == null) {
            return new ArrayList<String>();
        }
        try {
            FormDataDao formDataDao = (FormDataDao) AppUtil.getApplicationContext().getBean("formDataDao");
            String parentColumn = getParentColumn();
            String parentValue = resolveParentValue(formData);
            boolean excludeParent = parentColumn != null && parentValue != null && !parentValue.isEmpty();

            // Map unique Excel headers -> target field ids. A unique column that is not stored
            // (no field id, e.g. a preview-only column) cannot be queried against the table, so
            // skip it here; the in-file duplicate check above still covers it.
            Map<String, String> headerToField = new LinkedHashMap<String, String>();
            Map<String, Map<String, String>> headerToCol = new LinkedHashMap<String, Map<String, String>>();
            List<String> storedUniqueHeaders = new ArrayList<String>();
            for (Map<String, String> col : columns) {
                if ("true".equalsIgnoreCase(col.get("uniqueKey"))) {
                    String fieldId = col.get("fieldId");
                    if (fieldId != null && !fieldId.isEmpty()) {
                        headerToField.put(col.get("excelHeader"), fieldId);
                        headerToCol.put(col.get("excelHeader"), col);
                        storedUniqueHeaders.add(col.get("excelHeader"));
                    }
                }
            }
            // No stored unique column to query against: nothing can be an existing duplicate.
            if (storedUniqueHeaders.isEmpty()) {
                return new ArrayList<String>();
            }

            List<String> dupRows = new ArrayList<String>();
            for (int i = 0; i < arr.length(); i++) {
                JSONObject obj = arr.optJSONObject(i);
                StringBuilder condition = new StringBuilder(" WHERE ");
                List<Object> params = new ArrayList<Object>();
                boolean firstCond = true;
                for (String header : storedUniqueHeaders) {
                    if (!firstCond) {
                        condition.append(" AND ");
                    }
                    condition.append("e.customProperties.").append(headerToField.get(header)).append(" = ?");
                    params.add(coerce(getCell(obj, header, caseSensitive), headerToCol.get(header)));
                    firstCond = false;
                }
                // Ignore the rows belonging to the current parent (they are being replaced).
                if (excludeParent) {
                    condition.append(" AND e.customProperties.").append(parentColumn).append(" <> ?");
                    params.add(parentValue);
                }
                Long count = formDataDao.count(target.formDefId, target.tableName, condition.toString(), params.toArray());
                if (count != null && count > 0) {
                    dupRows.add(String.valueOf(i + 1));
                }
            }
            return dupRows;
        } catch (Exception e) {
            LogUtil.error(getClassName(), e, "Error checking duplicates against existing records");
            return null;
        }
    }

    @Override
    public FormData formatDataForValidation(FormData formData) {
        return formData;
    }

    //
    // ---- Helpers ----
    //

    @SuppressWarnings("unchecked")
    protected List<Map<String, String>> getColumns() {
        List<Map<String, String>> result = new ArrayList<Map<String, String>>();
        Object prop = getProperty("columns");
        if (prop instanceof Object[]) {
            for (Object o : (Object[]) prop) {
                if (o instanceof Map) {
                    Map<String, Object> raw = (Map<String, Object>) o;
                    String header = raw.get("excelHeader") != null ? raw.get("excelHeader").toString().trim() : "";
                    if (header.isEmpty()) {
                        continue;
                    }
                    Map<String, String> col = new LinkedHashMap<String, String>();
                    col.put("excelHeader", header);
                    col.put("fieldId", raw.get("fieldId") != null ? raw.get("fieldId").toString().trim() : "");
                    col.put("required", raw.get("required") != null ? raw.get("required").toString() : "");
                    col.put("uniqueKey", raw.get("uniqueKey") != null ? raw.get("uniqueKey").toString() : "");
                    col.put("hideInPreview", raw.get("hideInPreview") != null ? raw.get("hideInPreview").toString() : "");
                    col.put("type", raw.get("type") != null ? raw.get("type").toString().trim() : "");
                    col.put("defaultValue", raw.get("defaultValue") != null ? raw.get("defaultValue").toString() : "");
                    col.put("dateFormat", raw.get("dateFormat") != null ? raw.get("dateFormat").toString().trim() : "");
                    col.put("pattern", raw.get("pattern") != null ? raw.get("pattern").toString() : "");
                    col.put("minValue", raw.get("minValue") != null ? raw.get("minValue").toString().trim() : "");
                    col.put("maxValue", raw.get("maxValue") != null ? raw.get("maxValue").toString().trim() : "");
                    col.put("allowedValues", raw.get("allowedValues") != null ? raw.get("allowedValues").toString() : "");
                    result.add(col);
                }
            }
        }
        return result;
    }

    /**
     * @return {@code true} when the host form's own store binder must be prevented from persisting
     *         its (parent) record, so only the imported Excel rows are stored.
     */
    protected boolean isParentStoreDisabled() {
        return "true".equalsIgnoreCase(getPropertyString("disableParentStore"));
    }

    /** @return the configured parent (foreign-key) column, or {@code null} when no link is configured. */
    protected String getParentColumn() {
        String col = getPropertyString("parentColumnName");
        return (col == null || col.trim().isEmpty()) ? null : col.trim();
    }

    /**
     * Resolves the value written into (and queried against) the parent column.
     *
     * <p>Resolution order:</p>
     * <ol>
     *   <li>A value captured at render time and carried through the submission as a hidden field.
     *       This is essential for request-param hash variables (e.g. {@code #requestParam.referent_id#}):
     *       on submit the request is the POST to the form action URL, which no longer carries the
     *       original load-URL query parameters, so re-resolving the hash variable here would yield
     *       an empty string.</li>
     *   <li>The configured {@code parentColumnValue} with hash variables resolved (this is what
     *       runs at render time, while the original request parameters are still present).</li>
     *   <li>The current submission's primary key.</li>
     * </ol>
     */
    protected String resolveParentValue(FormData formData) {
        if (formData != null) {
            String carried = formData.getRequestParameter(getParentCarrierName());
            if (carried != null && !carried.trim().isEmpty()) {
                return carried.trim();
            }
        }
        String v = getPropertyString("parentColumnValue");
        if (v != null && !v.trim().isEmpty()) {
            String processed = AppUtil.processHashVariable(v.trim(), null, null, null);
            return (processed != null && !processed.isEmpty()) ? processed : v.trim();
        }
        return formData != null ? formData.getPrimaryKeyValue() : null;
    }

    /**
     * Name of the hidden field that carries the render-time resolved parent value through the
     * submission, so {@link #store} does not depend on the (now gone) original request parameters.
     */
    protected String getParentCarrierName() {
        return FormUtil.getElementParameterName(this) + "_resolvedParent";
    }

    /**
     * Name of the hidden field that carries the selected file name through the submission. A file
     * input cannot be repopulated by the browser on a validation reload, but the parsed rows live
     * in the (round-tripped) hidden value; carrying the name too lets the widget show the file bar
     * ("name + Retirer") instead of the empty drop zone after the reload.
     */
    protected String getFileNameCarrierName() {
        return FormUtil.getElementParameterName(this) + "_fileName";
    }

    /** @return the submitted file name carried through the request, or {@code ""} on a fresh load. */
    protected String resolveFileName(FormData formData) {
        if (formData != null) {
            String name = formData.getRequestParameter(getFileNameCarrierName());
            if (name != null && !name.trim().isEmpty()) {
                return name.trim();
            }
        }
        return "";
    }

    /** @return the replace strategy: "parentId" (default) or "uniqueKeys". */
    protected String getReplaceStrategy() {
        String s = getPropertyString("replaceStrategy");
        return (s == null || s.trim().isEmpty()) ? "parentId" : s.trim();
    }

    /**
     * Resolves where the imported rows are stored. By default this is the current form/table
     * (the form hosting this element). When "custom storage target" is enabled, it is either
     * the selected form or a raw table name (with the {@code app_fd_} prefix stripped).
     *
     * @return the resolved target, or {@code null} when it cannot be determined.
     */
    protected StorageTarget resolveStorageTarget() {
        AppDefinition appDef = AppUtil.getCurrentAppDefinition();
        AppService appService = (AppService) AppUtil.getApplicationContext().getBean("appService");

        if ("true".equalsIgnoreCase(getPropertyString("customStorageTarget"))) {
            String cibleType = getPropertyString("cibleType");
            if ("table".equalsIgnoreCase(cibleType)) {
                String table = stripTablePrefix(getPropertyString("tableName"));
                if (table == null || table.isEmpty()) {
                    return null;
                }
                // No form definition for a raw table: reuse the table name as the DAO entity key.
                return new StorageTarget(table, table);
            }
            // Default custom branch: a selected form.
            String formDefId = getPropertyString("formDefId");
            if (formDefId == null || formDefId.isEmpty()) {
                return null;
            }
            return new StorageTarget(formDefId, appService.getFormTableName(appDef, formDefId));
        }

        // Default: the current form/table hosting this element.
        String formDefId = getCurrentFormDefId();
        if (formDefId == null || formDefId.isEmpty()) {
            return null;
        }
        return new StorageTarget(formDefId, appService.getFormTableName(appDef, formDefId));
    }

    /** @return the form definition id of the form hosting this element, or {@code null}. */
    protected String getCurrentFormDefId() {
        Form rootForm = FormUtil.findRootForm(this);
        if (rootForm != null) {
            String id = rootForm.getPropertyString("id");
            if (id != null && !id.isEmpty()) {
                return id;
            }
        }
        return null;
    }

    /** Strips the Joget {@code app_fd_} table prefix so the name is usable with the DAO classes. */
    protected static String stripTablePrefix(String table) {
        if (table == null) {
            return null;
        }
        table = table.trim();
        if (table.toLowerCase().startsWith("app_fd_")) {
            table = table.substring("app_fd_".length());
        }
        return table;
    }

    /** Resolved storage target: the DAO entity key (form def id or table name) and the actual table. */
    protected static class StorageTarget {
        final String formDefId;
        final String tableName;

        StorageTarget(String formDefId, String tableName) {
            this.formDefId = formDefId;
            this.tableName = tableName;
        }
    }

    protected String getPropertyStringOrDefault(String name, String def) {
        String v = getPropertyString(name);
        return (v == null || v.isEmpty()) ? def : v;
    }

    /** Reads a cell from a JSON row by Excel header, honouring case sensitivity. */
    protected static String getCell(JSONObject obj, String header, boolean caseSensitive) {
        if (obj == null || header == null) {
            return "";
        }
        if (obj.has(header)) {
            return obj.optString(header, "");
        }
        if (!caseSensitive) {
            for (String key : obj.keySet()) {
                if (key.equalsIgnoreCase(header)) {
                    return obj.optString(key, "");
                }
            }
        }
        return "";
    }

    protected static boolean hasKey(JSONObject obj, String header, boolean caseSensitive) {
        if (obj == null || header == null) {
            return false;
        }
        if (obj.has(header)) {
            return true;
        }
        if (!caseSensitive) {
            for (String key : obj.keySet()) {
                if (key.equalsIgnoreCase(header)) {
                    return true;
                }
            }
        }
        return false;
    }

    protected static String compositeKey(JSONObject obj, List<Map<String, String>> cols, boolean caseSensitive) {
        StringBuilder sb = new StringBuilder();
        for (Map<String, String> col : cols) {
            String v = coerce(getCell(obj, col.get("excelHeader"), caseSensitive), col);
            if (!caseSensitive && v != null) {
                v = v.toLowerCase();
            }
            sb.append(v == null ? "" : v.trim()).append("");
        }
        return sb.toString();
    }

    //
    // ---- Column type coercion & cleansing ----
    //
    // Applied identically on the client (excel-import-lib.js) so the preview, the committed JSON
    // and what is stored/validated all agree. Coercion is best-effort: a value that cannot be
    // parsed for the requested type is left untouched rather than blanked, so a downstream
    // validation (or the user) can still see the offending input.

    /**
     * Builds the per-column coercion rule handed to the client, or {@code null} when the column
     * requests no cleansing (so the client config stays small for the common no-op case).
     */
    protected JSONObject coercionRule(Map<String, String> col) {
        String type = col.get("type");
        String def = col.get("defaultValue");
        String dateFormat = col.get("dateFormat");

        boolean hasType = type != null && !type.isEmpty() && !"text".equalsIgnoreCase(type);
        boolean hasDefault = def != null && !def.isEmpty();
        if (!hasType && !hasDefault) {
            return null;
        }
        JSONObject rule = new JSONObject();
        rule.put("type", type == null ? "" : type);
        rule.put("defaultValue", def == null ? "" : def);
        rule.put("dateFormat", dateFormat == null ? "" : dateFormat);
        return rule;
    }

    /** Applies the column's cleansing rules (default value, then type coercion) to a single cell value. */
    protected static String coerce(String raw, Map<String, String> col) {
        if (col == null) {
            return raw == null ? "" : raw;
        }
        String type = col.get("type");
        String def = col.get("defaultValue");
        String dateFormat = col.get("dateFormat");

        String v = raw == null ? "" : raw;
        if (v.isEmpty() && def != null && !def.isEmpty()) {
            v = def;
        }
        if (v.isEmpty()) {
            return v;
        }
        if ("number".equalsIgnoreCase(type)) {
            return coerceNumber(v);
        }
        if ("date".equalsIgnoreCase(type)) {
            return coerceDate(v, dateFormat);
        }
        if ("boolean".equalsIgnoreCase(type)) {
            return coerceBoolean(v);
        }
        return v;
    }

    /**
     * Normalises a numeric string: strips currency/spaces, resolves thousands vs decimal
     * separators, and emits a canonical plain number ({@code "1,234.50"} -> {@code "1234.5"}).
     * When both {@code ,} and {@code .} occur, the last one is taken as the decimal separator.
     * A lone {@code ,} is read as grouped thousands when it forms {@code 1,000} / {@code 1,234,567},
     * otherwise as a decimal separator (European style, e.g. {@code 12,5}).
     */
    protected static String coerceNumber(String v) {
        String s = v.trim().replaceAll("[^0-9,.\\-]", "");
        if (s.isEmpty() || "-".equals(s)) {
            return v;
        }
        boolean hasComma = s.indexOf(',') >= 0;
        boolean hasDot = s.indexOf('.') >= 0;
        if (hasComma && hasDot) {
            if (s.lastIndexOf(',') > s.lastIndexOf('.')) {
                s = s.replace(".", "").replace(',', '.');
            } else {
                s = s.replace(",", "");
            }
        } else if (hasComma) {
            // Lone comma: grouped thousands (1,000 / 1,234,567) vs. decimal separator (12,5).
            if (s.matches("-?\\d{1,3}(,\\d{3})+")) {
                s = s.replace(",", "");
            } else {
                s = s.replace(',', '.');
            }
        }
        try {
            java.math.BigDecimal bd = new java.math.BigDecimal(s);
            if (bd.compareTo(java.math.BigDecimal.ZERO) == 0) {
                return "0";
            }
            return bd.stripTrailingZeros().toPlainString();
        } catch (NumberFormatException e) {
            return v;
        }
    }

    /** Maps common truthy/falsy tokens to {@code "true"}/{@code "false"}; leaves anything else as-is. */
    protected static String coerceBoolean(String v) {
        String s = v.trim().toLowerCase();
        if (TRUE_TOKENS.contains(s)) {
            return "true";
        }
        if (FALSE_TOKENS.contains(s)) {
            return "false";
        }
        return v;
    }

    private static final Set<String> TRUE_TOKENS = new HashSet<String>(java.util.Arrays.asList(
            "true", "1", "yes", "y", "oui", "o", "vrai", "x", "on"));
    private static final Set<String> FALSE_TOKENS = new HashSet<String>(java.util.Arrays.asList(
            "false", "0", "no", "n", "non", "faux", "off"));

    // yyyy-MM-dd [THH:mm[:ss]]
    private static final java.util.regex.Pattern ISO_DATE = java.util.regex.Pattern.compile(
            "^(\\d{4})-(\\d{1,2})-(\\d{1,2})(?:[ T](\\d{1,2}):(\\d{1,2})(?::(\\d{1,2}))?)?");
    // day-first dd/MM/yyyy (also - or . separators), matching the plugin's French locale
    private static final java.util.regex.Pattern DMY_DATE = java.util.regex.Pattern.compile(
            "^(\\d{1,2})[/\\-.](\\d{1,2})[/\\-.](\\d{2,4})(?:[ T](\\d{1,2}):(\\d{1,2})(?::(\\d{1,2}))?)?");

    /**
     * Parses an Excel serial number or a supported date string into
     * {@code [year, month, day, hour, minute, second]}, or {@code null} when unrecognised.
     * Supported: Excel serials (1900 date system), ISO {@code yyyy-MM-dd}, and day-first
     * {@code dd/MM/yyyy}.
     */
    protected static int[] parseDateParts(String value) {
        if (value == null) {
            return null;
        }
        String s = value.trim();
        if (s.isEmpty()) {
            return null;
        }
        if (s.matches("\\d+(\\.\\d+)?")) {
            double serial = Double.parseDouble(s);
            // Plausible Excel serial range: 1900-01-01 .. 9999-12-31.
            if (serial > 0 && serial < 2958466) {
                long ms = Math.round((serial - 25569) * 86400000.0);
                java.util.Calendar c = java.util.Calendar.getInstance(java.util.TimeZone.getTimeZone("UTC"));
                c.setTimeInMillis(ms);
                return new int[]{c.get(java.util.Calendar.YEAR), c.get(java.util.Calendar.MONTH) + 1,
                        c.get(java.util.Calendar.DAY_OF_MONTH), c.get(java.util.Calendar.HOUR_OF_DAY),
                        c.get(java.util.Calendar.MINUTE), c.get(java.util.Calendar.SECOND)};
            }
        }
        java.util.regex.Matcher m = ISO_DATE.matcher(s);
        if (m.find()) {
            return new int[]{parseInt(m.group(1)), parseInt(m.group(2)), parseInt(m.group(3)),
                    parseInt(m.group(4)), parseInt(m.group(5)), parseInt(m.group(6))};
        }
        m = DMY_DATE.matcher(s);
        if (m.find()) {
            int year = parseInt(m.group(3));
            if (year < 100) {
                year += 2000;
            }
            return new int[]{year, parseInt(m.group(2)), parseInt(m.group(1)),
                    parseInt(m.group(4)), parseInt(m.group(5)), parseInt(m.group(6))};
        }
        return null;
    }

    /** Formats a parsed date to the given pattern (default {@code yyyy-MM-dd}); tokens yyyy MM dd HH mm ss. */
    protected static String coerceDate(String v, String fmt) {
        int[] p = parseDateParts(v);
        if (p == null) {
            return v;
        }
        String f = (fmt == null || fmt.trim().isEmpty()) ? "yyyy-MM-dd" : fmt.trim();
        return f.replace("yyyy", pad(p[0], 4))
                .replace("MM", pad(p[1], 2))
                .replace("dd", pad(p[2], 2))
                .replace("HH", pad(p[3], 2))
                .replace("mm", pad(p[4], 2))
                .replace("ss", pad(p[5], 2));
    }

    private static int parseInt(String s) {
        return (s == null || s.isEmpty()) ? 0 : Integer.parseInt(s);
    }

    private static String pad(int n, int width) {
        String s = Integer.toString(n);
        while (s.length() < width) {
            s = "0" + s;
        }
        return s;
    }

    //
    // ---- Per-column value validation ----
    //
    // Applied identically on the client (excel-import-lib.js), after coercion, on the already
    // cleansed value. Only non-empty cells are checked; emptiness is governed by the required-cell
    // check. Best-effort like coercion: a value that cannot be parsed for a min/max comparison is
    // not flagged (a regex pattern is the tool for enforcing a strict format).

    /**
     * Builds the per-column validation rule handed to the client, or {@code null} when the column
     * declares no value rules (so the client config stays small for the common no-op case).
     */
    protected JSONObject validationRule(Map<String, String> col) {
        if (!hasCellRules(col)) {
            return null;
        }
        String pattern = col.get("pattern");
        String min = col.get("minValue");
        String max = col.get("maxValue");
        JSONObject r = new JSONObject();
        r.put("type", col.get("type") == null ? "" : col.get("type"));
        r.put("pattern", pattern == null ? "" : pattern);
        r.put("min", min == null ? "" : min);
        r.put("max", max == null ? "" : max);
        JSONArray allowed = new JSONArray();
        for (String s : parseAllowed(col.get("allowedValues"))) {
            allowed.put(s);
        }
        r.put("allowed", allowed);
        return r;
    }

    /** @return {@code true} when the column declares a regex, a min/max bound, or allowed values. */
    protected static boolean hasCellRules(Map<String, String> col) {
        if (col == null) {
            return false;
        }
        String pattern = col.get("pattern");
        String min = col.get("minValue");
        String max = col.get("maxValue");
        return (pattern != null && !pattern.isEmpty())
                || (min != null && !min.isEmpty())
                || (max != null && !max.isEmpty())
                || !parseAllowed(col.get("allowedValues")).isEmpty();
    }

    /** Splits the comma-separated allowed-values list into trimmed, non-empty entries. */
    protected static List<String> parseAllowed(String raw) {
        List<String> out = new ArrayList<String>();
        if (raw == null) {
            return out;
        }
        for (String part : raw.split(",")) {
            String t = part.trim();
            if (!t.isEmpty()) {
                out.add(t);
            }
        }
        return out;
    }

    /**
     * Validates a single (coerced, non-empty) cell against the column's value rules, in order:
     * allowed values, min/max range, regex pattern.
     *
     * @return the failing rule ({@code "allowed"}/{@code "range"}/{@code "pattern"}), or
     *         {@code null} when the value passes.
     */
    protected static String validateCell(String value, Map<String, String> col, boolean caseSensitive) {
        String s = value == null ? "" : value.trim();

        List<String> allowed = parseAllowed(col.get("allowedValues"));
        if (!allowed.isEmpty()) {
            boolean ok = false;
            for (String a : allowed) {
                if (caseSensitive ? a.equals(s) : a.equalsIgnoreCase(s)) {
                    ok = true;
                    break;
                }
            }
            if (!ok) {
                return "allowed";
            }
        }

        String min = col.get("minValue");
        String max = col.get("maxValue");
        if ((min != null && !min.isEmpty()) || (max != null && !max.isEmpty())) {
            if (cellOutOfRange(s, col.get("type"), min, max)) {
                return "range";
            }
        }

        String pattern = col.get("pattern");
        if (pattern != null && !pattern.isEmpty()) {
            try {
                // find() (not matches()) to mirror JavaScript RegExp.test — anchor with ^...$ for a
                // full match.
                if (!java.util.regex.Pattern.compile(pattern).matcher(s).find()) {
                    return "pattern";
                }
            } catch (Exception e) {
                /* invalid configured regex: skip */
            }
        }

        return null;
    }

    /** Range check for number/date typed columns; returns {@code false} for any other type. */
    protected static boolean cellOutOfRange(String s, String type, String min, String max) {
        String t = type == null ? "" : type.toLowerCase();
        if ("date".equals(t)) {
            Long val = dateToMillis(s);
            if (val == null) {
                return false;
            }
            Long lo = (min != null && !min.isEmpty()) ? dateToMillis(min) : null;
            Long hi = (max != null && !max.isEmpty()) ? dateToMillis(max) : null;
            return (lo != null && val < lo) || (hi != null && val > hi);
        }
        if ("number".equals(t)) {
            Double val = cellToNumber(s);
            if (val == null) {
                return false;
            }
            Double lo = (min != null && !min.isEmpty()) ? cellToNumber(min) : null;
            Double hi = (max != null && !max.isEmpty()) ? cellToNumber(max) : null;
            return (lo != null && val < lo) || (hi != null && val > hi);
        }
        return false;
    }

    /** Parses a (possibly grouped/European) number string to a Double, or {@code null}. */
    protected static Double cellToNumber(String v) {
        try {
            return Double.valueOf(coerceNumber(v));
        } catch (Exception e) {
            return null;
        }
    }

    /** Parses an Excel serial / ISO / day-first date to a comparable UTC millisecond value, or {@code null}. */
    protected static Long dateToMillis(String v) {
        int[] p = parseDateParts(v);
        if (p == null) {
            return null;
        }
        java.util.Calendar c = java.util.Calendar.getInstance(java.util.TimeZone.getTimeZone("UTC"));
        c.clear();
        c.set(p[0], p[1] - 1, p[2], p[3], p[4], p[5]);
        return c.getTimeInMillis();
    }

    protected String msg(String key) {
        return AppPluginUtil.getMessage(key, getClassName(), MESSAGES_PATH);
    }

    /**
     * Resolves a user-facing message: uses the value configured under {@code propName}
     * (the "Messages d'erreur" tab) when set, otherwise falls back to the localised
     * default keyed by {@code defaultKey} from messages.properties.
     */
    protected String customMsg(String propName, String defaultKey) {
        String custom = getPropertyString(propName);
        if (custom != null && !custom.trim().isEmpty()) {
            return custom;
        }
        return msg(defaultKey);
    }

    protected static String defaultStr(String v, String def) {
        return (v == null || v.trim().isEmpty()) ? def : v;
    }

    protected static double parseDouble(String v, double def) {
        try {
            return (v == null || v.trim().isEmpty()) ? def : Double.parseDouble(v.trim());
        } catch (NumberFormatException e) {
            return def;
        }
    }
}
