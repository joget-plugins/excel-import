package org.joget.plugin.melkart;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import org.joget.apps.app.service.AppUtil;
import org.joget.apps.form.dao.FormDataDao;
import org.joget.apps.form.model.Element;
import org.joget.apps.form.model.FormBinder;
import org.joget.apps.form.model.FormData;
import org.joget.apps.form.model.FormLoadBinder;
import org.joget.apps.form.model.FormRow;
import org.joget.apps.form.model.FormRowSet;
import org.joget.apps.form.model.FormStoreBinder;
import org.joget.commons.util.LogUtil;
import org.joget.commons.util.UuidGenerator;
import org.joget.workflow.util.WorkflowUtil;

/**
 * Store/load binder for the {@link ExcelParser} element.
 *
 * <p>{@code ExcelParser} is a form {@link Element}; it cannot also be a {@link FormBinder} because
 * the two are sibling types in Joget (a class can be one or the other, not both). This binder
 * therefore lives as its own {@code FormBinder} subclass and is supplied by the element through
 * {@link ExcelParser#getStoreBinder()} / {@link ExcelParser#getLoadBinder()}.</p>
 *
 * <p>All configuration (columns, storage target, parent link, replace strategy) lives on the
 * element, so the binder reads it from the {@code element} passed to {@link #store} / {@link #load}.
 * It carries no properties of its own.</p>
 */
public class ExcelParserBinder extends FormBinder implements FormStoreBinder, FormLoadBinder {

    @Override
    public String getClassName() {
        return getClass().getName();
    }

    @Override
    public String getName() {
        return "Excel Parser Binder";
    }

    @Override
    public String getLabel() {
        return "Excel Parser Binder";
    }

    @Override
    public String getDescription() {
        return "Store/load binder for the Excel Parser element.";
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
        // No configuration of its own: everything is read from the host element.
        return "[]";
    }

    //
    // ---- Store binder ----
    //

    @Override
    public FormRowSet store(Element element, FormRowSet rows, FormData formData) {
        if (!(element instanceof ExcelParser)) {
            LogUtil.warn(getClassName(), "ExcelParserBinder is attached to a non-ExcelParser element; skipping store.");
            return rows;
        }
        ExcelParser ep = (ExcelParser) element;

        ExcelParser.StorageTarget target = ep.resolveStorageTarget();
        if (target == null) {
            LogUtil.warn(getClassName(), "No storage target could be resolved; skipping Excel import store.");
            return rows;
        }

        FormDataDao formDataDao = (FormDataDao) AppUtil.getApplicationContext().getBean("formDataDao");
        String parentColumn = ep.getParentColumn();              // may be null when no link is configured
        String parentValue = ep.resolveParentValue(formData);    // may be null when no value resolves
        LogUtil.info(getClassName(), "Excel import store: target=" + target.tableName
                + ", parentColumn=" + parentColumn + ", parentValue=" + parentValue);
        boolean linkParent = parentColumn != null && parentValue != null && !parentValue.isEmpty();

        try {
            // Replace strategy: remove the previous rows before saving the new ones.
            if ("true".equalsIgnoreCase(ep.getPropertyStringOrDefault("deleteExisting", "true"))) {
                if ("uniqueKeys".equalsIgnoreCase(ep.getReplaceStrategy())) {
                    deleteByUniqueKeys(ep, formDataDao, target, rows);
                } else if (linkParent) {
                    FormRowSet existing = formDataDao.find(target.formDefId, target.tableName,
                            " WHERE e.customProperties." + parentColumn + " = ?",
                            new Object[]{parentValue}, null, null, null, null);
                    if (existing != null && !existing.isEmpty()) {
                        formDataDao.delete(target.formDefId, target.tableName, existing);
                    }
                } else {
                    LogUtil.info(getClassName(), "Replace-by-parent requested but no parent column/value is configured; skipping delete.");
                }
            }

            if (rows != null && !rows.isEmpty()) {
                Date now = new Date();
                String user = WorkflowUtil.getCurrentUsername();
                for (FormRow row : rows) {
                    if (row.getId() == null || row.getId().isEmpty()) {
                        row.setId(UuidGenerator.getInstance().getUuid());
                    }
                    if (linkParent) {
                        row.setProperty(parentColumn, parentValue);
                    }
                    if (row.getDateCreated() == null) {
                        row.setDateCreated(now);
                    }
                    row.setDateModified(now);
                    if (row.getCreatedBy() == null) {
                        row.setCreatedBy(user);
                    }
                    row.setModifiedBy(user);
                }
                formDataDao.saveOrUpdate(target.formDefId, target.tableName, rows);
            }
        } catch (Exception e) {
            LogUtil.error(getClassName(), e, "Error storing Excel import rows to " + target.tableName);
        }
        return rows;
    }

    /**
     * Deletes existing rows whose unique-key combination matches any of the incoming rows
     * (upsert-by-unique-key replace strategy). No-op when no unique-key columns are configured.
     */
    protected void deleteByUniqueKeys(ExcelParser ep, FormDataDao formDataDao, ExcelParser.StorageTarget target, FormRowSet rows) {
        if (rows == null || rows.isEmpty()) {
            return;
        }
        List<String> uniqueFieldIds = new ArrayList<String>();
        for (Map<String, String> col : ep.getColumns()) {
            if ("true".equalsIgnoreCase(col.get("uniqueKey"))) {
                String fieldId = col.get("fieldId");
                if (fieldId != null && !fieldId.isEmpty()) {
                    uniqueFieldIds.add(fieldId);
                }
            }
        }
        if (uniqueFieldIds.isEmpty()) {
            LogUtil.warn(getClassName(), "Replace-by-unique-keys requested but no unique-key columns are configured; skipping delete.");
            return;
        }

        FormRowSet toDelete = new FormRowSet();
        for (FormRow row : rows) {
            StringBuilder condition = new StringBuilder(" WHERE ");
            List<Object> params = new ArrayList<Object>();
            boolean first = true;
            for (String fieldId : uniqueFieldIds) {
                if (!first) {
                    condition.append(" AND ");
                }
                condition.append("e.customProperties.").append(fieldId).append(" = ?");
                params.add(row.getProperty(fieldId));
                first = false;
            }
            FormRowSet existing = formDataDao.find(target.formDefId, target.tableName,
                    condition.toString(), params.toArray(), null, null, null, null);
            if (existing != null) {
                toDelete.addAll(existing);
            }
        }
        if (!toDelete.isEmpty()) {
            formDataDao.delete(target.formDefId, target.tableName, toDelete);
        }
    }

    //
    // ---- Load binder (rebuild preview on edit) ----
    //

    @Override
    public FormRowSet load(Element element, String primaryKey, FormData formData) {
        FormRowSet rowSet = new FormRowSet();
        rowSet.setMultiRow(true);
        if (!(element instanceof ExcelParser)) {
            return rowSet;
        }
        if (primaryKey == null || primaryKey.isEmpty()) {
            return rowSet;
        }
        ExcelParser ep = (ExcelParser) element;
        FormRowSet existing = ep.findChildRows(primaryKey, formData);
        if (existing != null) {
            rowSet.addAll(existing);
        }
        return rowSet;
    }
}
