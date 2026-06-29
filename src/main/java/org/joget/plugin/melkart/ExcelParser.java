package org.joget.plugin.melkart;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import javax.servlet.http.HttpServletRequest;

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
import org.joget.workflow.util.WorkflowUtil;
import org.json.JSONArray;
import org.json.JSONObject;

/**
 * Excel Parser form element.
 *
 * <p>Renders a self-contained Excel import widget (own drop zone + client side preview,
 * parsed with a bundled copy of SheetJS/xlsx). The validated rows are written, keyed by
 * Excel header, into a hidden field as a JSON array.</p>
 *
 * <p>Persistence is delegated to a companion {@link ExcelParserBinder} (a real
 * {@link FormBinder}) which this element supplies through {@link #getStoreBinder()} /
 * {@link #getLoadBinder()}: on submit every Excel row is stored as one record in a configurable
 * target form/table (multi-row), each linked to the current submission via a configurable parent
 * foreign-key column. The binder is a separate class because Joget's {@code Element} and
 * {@code FormBinder} are distinct (sibling) types and a single class cannot be both. The
 * front-end validation (required headers, required cells across columns, composite duplicate
 * key across columns) is mirrored server-side in {@link #selfValidate(FormData)} and blocks
 * the submission on any error.</p>
 */
public class ExcelParser extends Element implements FormBuilderPaletteElement, FormContainer {

    public static final String MESSAGES_PATH = "messages/ExcelParser";

    @Override
    public String getClassName() {
        return getClass().getName();
    }

    @Override
    public String getName() {
        return "Excel Parser";
    }

    @Override
    public String getLabel() {
        return AppPluginUtil.getMessage("ExcelParser.pluginLabel", getClassName(), MESSAGES_PATH);
    }

    @Override
    public String getDescription() {
        return AppPluginUtil.getMessage("ExcelParser.pluginDesc", getClassName(), MESSAGES_PATH);
    }

    @Override
    public String getVersion() {
        try {
            Properties props = new Properties();
            try (InputStream is = getClass().getResourceAsStream(
                    "/META-INF/maven/org.joget.plugin.melkart/excel-parser/pom.properties")) {
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
        return AppUtil.readPluginResource(getClassName(), "/properties/ExcelParser.json", null, true, MESSAGES_PATH);
    }

    //
    // ---- Form builder palette ----
    //

    @Override
    public String getFormBuilderTemplate() {
        return "<div style='border: 2px dashed #667eea;padding: 20px;text-align: center;border-radius: 10px;color:#667eea;'>"
                + "<i class=\"fas fa-file-excel\"></i> Excel Parser</div>";
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
    // Instead we hand back a cached ExcelParserBinder instance. It is cached (via the inherited
    // store/load binder fields) because Joget keys the formatData() row set by the exact binder
    // instance returned here and later looks it up again to invoke store().

    @Override
    public FormStoreBinder getStoreBinder() {
        FormStoreBinder binder = super.getStoreBinder();
        if (binder == null) {
            ExcelParserBinder excelBinder = new ExcelParserBinder();
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
            ExcelParserBinder excelBinder = new ExcelParserBinder();
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
        String template = "ExcelParser.ftl";

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
        dataModel.put("jsConfig", buildClientConfig());
        dataModel.put("dropzoneText", defaultStr(getPropertyString("dropzoneText"),
                AppPluginUtil.getMessage("ExcelParser.dropzoneTextDefault", getClassName(), MESSAGES_PATH)));

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
            JSONObject mapping = new JSONObject();
            JSONArray requiredColumns = new JSONArray();
            JSONArray uniqueColumns = new JSONArray();
            for (Map<String, String> col : columns) {
                String header = col.get("excelHeader");
                headers.put(header);
                mapping.put(header, col.get("fieldId"));
                if ("true".equalsIgnoreCase(col.get("required"))) {
                    requiredColumns.put(header);
                }
                if ("true".equalsIgnoreCase(col.get("uniqueKey"))) {
                    uniqueColumns.put(header);
                }
            }

            cfg.put("elementId", getPropertyString("id"));
            cfg.put("headers", headers);
            cfg.put("mapping", mapping);
            cfg.put("requiredColumns", requiredColumns);
            cfg.put("uniqueColumns", uniqueColumns);
            cfg.put("caseSensitive", "true".equalsIgnoreCase(getPropertyString("caseSensitiveHeaders")));
            cfg.put("required", "true".equalsIgnoreCase(getPropertyString("required")));
            cfg.put("maxFileSizeMB", parseDouble(getPropertyString("maxFileSizeMB"), 5));
            cfg.put("previewHeight", defaultStr(getPropertyString("previewHeight"), "400"));
            cfg.put("showPreview", !"true".equalsIgnoreCase(getPropertyString("hidePreview")));

            // Localised messages shared with the client (overridable per element).
            JSONObject messages = new JSONObject();
            messages.put("missingHeaders", customMsg("msgMissingHeaders", "ExcelParser.err.missingHeaders"));
            messages.put("requiredCell", customMsg("msgRequiredCell", "ExcelParser.err.requiredCell"));
            messages.put("duplicate", customMsg("msgDuplicate", "ExcelParser.err.duplicate"));
            messages.put("emptyFile", customMsg("msgEmptyFile", "ExcelParser.err.emptyFile"));
            messages.put("readError", customMsg("msgReadError", "ExcelParser.err.readError"));
            messages.put("fileTooLarge", customMsg("msgFileTooLarge", "ExcelParser.err.fileTooLarge"));
            messages.put("rowsValid", customMsg("msgRowsValid", "ExcelParser.info.rowsValid"));
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
                    row.setProperty(fieldId, getCell(obj, header, caseSensitive));
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
                formData.addFormError(fieldId, customMsg("msgRequired", "ExcelParser.err.required"));
                return false;
            }
            return true;
        }

        List<Map<String, String>> columns = getColumns();
        boolean caseSensitive = "true".equalsIgnoreCase(getPropertyString("caseSensitiveHeaders"));

        JSONArray arr;
        try {
            arr = new JSONArray(value);
        } catch (Exception e) {
            formData.addFormError(fieldId, customMsg("msgInvalidData", "ExcelParser.err.invalidData"));
            return false;
        }

        if (arr.length() == 0) {
            if (required) {
                formData.addFormError(fieldId, customMsg("msgRequired", "ExcelParser.err.required"));
                return false;
            }
            return true;
        }

        // 1. Header structure.
        List<String> missingHeaders = new ArrayList<String>();
        JSONObject first = arr.optJSONObject(0);
        for (Map<String, String> col : columns) {
            String header = col.get("excelHeader");
            if (!hasKey(first, header, caseSensitive)) {
                missingHeaders.add(header);
            }
        }
        if (!missingHeaders.isEmpty()) {
            formData.addFormError(fieldId, format(customMsg("msgMissingHeaders", "ExcelParser.err.missingHeaders"), join(missingHeaders, ", "), ""));
            return false;
        }

        // 2. Required cells across the flagged columns.
        for (Map<String, String> col : columns) {
            if (!"true".equalsIgnoreCase(col.get("required"))) {
                continue;
            }
            String header = col.get("excelHeader");
            List<String> badRows = new ArrayList<String>();
            for (int i = 0; i < arr.length(); i++) {
                String cell = getCell(arr.optJSONObject(i), header, caseSensitive);
                if (cell == null || cell.trim().isEmpty()) {
                    badRows.add(String.valueOf(i + 1));
                }
            }
            if (!badRows.isEmpty()) {
                formData.addFormError(fieldId, format(customMsg("msgRequiredCell", "ExcelParser.err.requiredCell"), header, join(badRows, ", ")));
                return false;
            }
        }

        // 3. Composite duplicate key across the flagged columns (within the file).
        List<String> uniqueHeaders = new ArrayList<String>();
        for (Map<String, String> col : columns) {
            if ("true".equalsIgnoreCase(col.get("uniqueKey"))) {
                uniqueHeaders.add(col.get("excelHeader"));
            }
        }
        if (!uniqueHeaders.isEmpty()) {
            Map<String, Integer> seen = new LinkedHashMap<String, Integer>();
            List<String> dupRows = new ArrayList<String>();
            for (int i = 0; i < arr.length(); i++) {
                JSONObject obj = arr.optJSONObject(i);
                String key = compositeKey(obj, uniqueHeaders, caseSensitive);
                if (seen.containsKey(key)) {
                    dupRows.add(String.valueOf(i + 1));
                } else {
                    seen.put(key, i);
                }
            }
            if (!dupRows.isEmpty()) {
                formData.addFormError(fieldId, format(customMsg("msgDuplicate", "ExcelParser.err.duplicate"), join(uniqueHeaders, " + "), join(dupRows, ", ")));
                return false;
            }

            // 4. Duplicate against existing records in the target table (optional).
            if ("true".equalsIgnoreCase(getPropertyString("checkExistingDuplicates"))) {
                List<String> existingDupRows = findExistingDuplicates(arr, uniqueHeaders, columns, caseSensitive, formData);
                if (existingDupRows == null) {
                    formData.addFormError(fieldId, customMsg("msgInvalidData", "ExcelParser.err.invalidData"));
                    return false;
                }
                if (!existingDupRows.isEmpty()) {
                    formData.addFormError(fieldId, format(customMsg("msgExistingDuplicate", "ExcelParser.err.existingDuplicate"), join(uniqueHeaders, " + "), join(existingDupRows, ", ")));
                    return false;
                }
            }
        }

        return true;
    }

    /**
     * @return list of 1-based row numbers that already exist in the target table, or null on error.
     */
    protected List<String> findExistingDuplicates(JSONArray arr, List<String> uniqueHeaders,
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

            // Map unique Excel headers -> target field ids.
            Map<String, String> headerToField = new LinkedHashMap<String, String>();
            for (Map<String, String> col : columns) {
                if ("true".equalsIgnoreCase(col.get("uniqueKey"))) {
                    headerToField.put(col.get("excelHeader"), col.get("fieldId"));
                }
            }

            List<String> dupRows = new ArrayList<String>();
            for (int i = 0; i < arr.length(); i++) {
                JSONObject obj = arr.optJSONObject(i);
                StringBuilder condition = new StringBuilder(" WHERE ");
                List<Object> params = new ArrayList<Object>();
                boolean firstCond = true;
                for (String header : uniqueHeaders) {
                    if (!firstCond) {
                        condition.append(" AND ");
                    }
                    condition.append("e.customProperties.").append(headerToField.get(header)).append(" = ?");
                    params.add(getCell(obj, header, caseSensitive));
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

    protected static String compositeKey(JSONObject obj, List<String> headers, boolean caseSensitive) {
        StringBuilder sb = new StringBuilder();
        for (String h : headers) {
            String v = getCell(obj, h, caseSensitive);
            if (!caseSensitive && v != null) {
                v = v.toLowerCase();
            }
            sb.append(v == null ? "" : v.trim()).append("");
        }
        return sb.toString();
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

    protected static String format(String template, String arg0, String arg1) {
        if (template == null) {
            return "";
        }
        return template.replace("{0}", arg0 == null ? "" : arg0).replace("{1}", arg1 == null ? "" : arg1);
    }

    protected static String join(List<String> items, String sep) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < items.size(); i++) {
            if (i > 0) {
                sb.append(sep);
            }
            sb.append(items.get(i));
        }
        return sb.toString();
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
