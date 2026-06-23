package org.joget.plugin.melkart;

import org.joget.apps.app.service.AppUtil;
import org.joget.apps.form.model.Element;
import org.joget.apps.form.model.FormBuilderPalette;
import org.joget.apps.form.model.FormBuilderPaletteElement;
import org.joget.apps.form.model.FormData;
import org.joget.apps.form.service.FormUtil;
import java.util.Map;
import java.util.Properties;
import java.io.InputStream;
import java.io.IOException;
import org.joget.commons.util.LogUtil;

public class ExcelParser extends Element implements FormBuilderPaletteElement {

    public static String MESSAGES_PATH = "messages/ExcelParser";

    @Override
	public String getClassName() {
		return getClass().getName();
	}

	@Override
	public String getLabel() {
		return "ExcelParser";
	}

	@Override
	public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(), "/properties/ExcelParser.json", null, true, MESSAGES_PATH);
	}

	@Override
	public String getDescription() {
		return "ExcelParser description";
	}

	@Override
	public String getName() {
		return "ExcelParser";
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
            LogUtil.warn(getClassName(), "Error getting version from pom.xml file : " + e.getMessage());
        }
        return "unknown";
	}

	@Override
	public Boolean selfValidate(FormData formData) {
		return true;
	}

	@Override
	public FormData formatDataForValidation(FormData formData) {
		return formData;
	}

    @Override
	public String getFormBuilderTemplate() {
        return "<div style='border: 2px dashed black;padding: 20px;text-align: center;border-radius: 10px;'><i class=\"fas fa-list-alt\"></i> ExcelParser</div>";
	}

	@Override
	public String getFormBuilderCategory() {
		return FormBuilderPalette.CATEGORY_CUSTOM;
	}

	@Override
	public String getFormBuilderIcon() {
		return "<i class=\"fas fa-list-alt\"></i>";
	}

	@Override
	public int getFormBuilderPosition() {
		return 100;
	}

    @Override
	public String renderTemplate(FormData formData, Map dataModel) {
		String template = "ExcelParser.ftl";

        String value = FormUtil.getElementPropertyValue(this, formData);
        dataModel.put("value", value);

        String html = FormUtil.generateElementHtml(this, formData, template, dataModel);
        return html;
	}
}
