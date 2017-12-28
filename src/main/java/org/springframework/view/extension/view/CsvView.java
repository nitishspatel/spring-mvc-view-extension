package org.springframework.view.extension.view;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Date;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.springframework.view.extension.bean.CsvBean;
import org.springframework.view.extension.config.File;
import org.springframework.view.extension.config.Header;
import org.supercsv.io.ICsvBeanWriter;

/**
 * Spring MVC view to process CSV rendering which can be configured as
 * <code>AbstractView</code>. It converts List of <code>CsvBean</code>  returned in the model from controller to a CSV file.
 * List<CsvBean>
 * 
 * @author nitish
 *
 */
public class CsvView extends AbstractCsvView {

	private static final Logger log = Logger.getLogger(CsvView.class);

	@Override
	protected void buildCsvDocument(ICsvBeanWriter csvWriter, Map<String, Object> model) throws IOException {
		log.debug("Start");

		List<CsvBean> csvBeanList = null;
		for (Object o : model.values()) {
			if (o instanceof List && ((List) o).isEmpty()
					|| !((List) o).get(0).getClass().isAnnotationPresent(File.class)) {
				csvBeanList = (List) o;
				break;
			}
		}
		if (csvBeanList == null) {
			throw new IllegalArgumentException();
		}

		if (model.get("csvData") == null || ((List<CsvBean>) model.get("csvData")).isEmpty()
				|| !((List<CsvBean>) model.get("csvData")).get(0).getClass().isAnnotationPresent(File.class)) {
			throw new IllegalArgumentException();
		}

		if (csvBeanList != null && !csvBeanList.isEmpty()) {
			CsvBean csvBean = csvBeanList.get(0);

			File file = csvBean.getClass().getAnnotation(File.class);
			setFileName(file.title() + (file.appendDate() ? new Date().getTime() : "")
					+ (file.appendExtension() ? ".csv" : ""));

			Field[] fields = csvBean.getClass().getDeclaredFields();

			int columnCount = getColumnCount(fields);

			String[] headerTitleNameList = new String[columnCount];
			String[] headerFieldNameList = new String[columnCount];
			Header fAnno = null;
			for (Field field : fields) {
				if (field.isAnnotationPresent(Header.class)) {
					fAnno = field.getAnnotation(Header.class);
					headerFieldNameList[fAnno.columnNumber() - 1] = field.getName();
					headerTitleNameList[fAnno.columnNumber() - 1] = fAnno.title();
				}
			}

			csvWriter.writeHeader(headerTitleNameList);

			for (CsvBean csvRowItem : csvBeanList) {
				csvWriter.write(csvRowItem, headerFieldNameList);
			}
		}

	}

	private int getColumnCount(Field[] fields) {
		int columnCount = 0;
		for (Field field : fields) {
			if (field.isAnnotationPresent(Header.class)) {
				columnCount++;
			}
		}
		return columnCount;

	}
}
