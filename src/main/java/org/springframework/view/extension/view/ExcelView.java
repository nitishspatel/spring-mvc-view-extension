package org.springframework.view.extension.view;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Date;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.view.extension.bean.ExcelBean;
import org.springframework.view.extension.config.DynamicExcelHeader;
import org.springframework.view.extension.config.File;
import org.springframework.view.extension.config.Header;

/**
 * Spring MVC view to process CSV rendering which can be configured as
 * <code>AbstractView</code>. It converts List of <code>ExcelBean</code>  returned in the model from controller to an excel file.
 * List<CsvBean>
 * 
 * @author nitish
 *
 */
public class ExcelView extends AbstractExcelView {

	private static final Logger log = Logger.getLogger(ExcelView.class);

	@Override
	protected void buildExcelDocument(Map<String, Object> model, HttpServletRequest request,
			HttpServletResponse response) throws IOException {
		log.debug("Start");

		List<ExcelBean> excelBeanList = null;
		for (Object o : model.values()) {
			if (o instanceof List && ((List) o).isEmpty()
					|| !((List) o).get(0).getClass().isAnnotationPresent(File.class)) {
				excelBeanList = (List) o;
				break;
			}
		}
		if (excelBeanList == null) {
			throw new IllegalArgumentException();
		}

		if (excelBeanList != null && !excelBeanList.isEmpty()) {
			ExcelBean excelBean = excelBeanList.get(0);

			File file = excelBean.getClass().getAnnotation(File.class);
			setFileName(file.title() + (file.appendDate() ? new Date().getTime() : "")
					+ (file.appendExtension() ? ".xlsx" : ""));
			int dataRowStart = file.dataRowStart();

			XSSFWorkbook workbook = getTemplateSource(file.template(), request);

			XSSFSheet sheet = workbook.getSheet(file.sheetName());

			Field[] fields = excelBean.getClass().getDeclaredFields();

			int columnCount = getColumnCount(fields);

			String[] headerFieldNameList = new String[columnCount];
			Header fAnno = null;
			for (Field field : fields) {
				if (field.isAnnotationPresent(Header.class)) {
					fAnno = field.getAnnotation(Header.class);
					headerFieldNameList[fAnno.columnNumber() - 1] = field.getName();
				}
			}

			// csvWriter.writeHeader(headerTitleNameList);
			ExcelBean excelRowItem = null;
			Method m = null;
			XSSFCell cell = null;
			String fieldName = null;
			for (int i = 0; i < excelBeanList.size(); i++) {
				excelRowItem = excelBeanList.get(i);
				for (int j = 0; j < headerFieldNameList.length; j++) {
					fieldName = headerFieldNameList[j];
					cell = getCell(sheet, i + dataRowStart, j);
					try {
						Header headerAnnotation = excelBean.getClass().getDeclaredField(fieldName)
								.getAnnotation(Header.class);
						m = excelBean.getClass().getMethod("get" + getFirstCharInUpperCase(fieldName));

						if (headerAnnotation.type().equals(Integer.class)) {
							setIntValue(cell, m.invoke(excelRowItem) == null ? "" : m.invoke(excelRowItem).toString());
						} else if (headerAnnotation.type().equals(Double.class)) {
							setDoubleVal(cell, m.invoke(excelRowItem) == null ? "" : m.invoke(excelRowItem).toString());
						} else if (headerAnnotation.type().equals(Float.class)) {
							setFloatVal(cell, m.invoke(excelRowItem) == null ? "" : m.invoke(excelRowItem).toString());
						} else if (headerAnnotation.type().equals(Date.class)) {
							setDateValue(cell, m.invoke(excelRowItem) == null ? "" : m.invoke(excelRowItem).toString(),
									workbook);
						} else {
							setText(cell, m.invoke(excelRowItem) == null ? "" : m.invoke(excelRowItem).toString());
						}

						if (StringUtils.isNotBlank(headerAnnotation.colorFormatter())) {
							m = excelBean.getClass().getMethod(headerAnnotation.colorFormatter());
							Object obj = m.invoke(excelRowItem);
							if (obj != null) {
								if (headerAnnotation.type().equals(Date.class)) {
									setColorForDateField(cell, Short.valueOf(obj.toString()), workbook);
								} else {
									setColor(cell, Short.valueOf(obj.toString()), workbook);
								}
							}
						}

					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}

			Method[] methods = excelBean.getClass().getDeclaredMethods();

			DynamicExcelHeader dynamicExcelHeader = null;
			for (Method method : methods) {
				if (method.isAnnotationPresent(DynamicExcelHeader.class)) {
					dynamicExcelHeader = method.getAnnotation(DynamicExcelHeader.class);
					try {
						setText(getCell(sheet, dynamicExcelHeader.rowNumber(), dynamicExcelHeader.columnNumber()),
								method.invoke(null).toString());

						if (StringUtils.isNotBlank(dynamicExcelHeader.colorFormatter())) {
							m = excelBean.getClass().getMethod(dynamicExcelHeader.colorFormatter());
							Object obj = method.invoke(null);
							if (obj != null) {
								setColor(
										getCell(sheet, dynamicExcelHeader.rowNumber(),
												dynamicExcelHeader.columnNumber()),
										Short.valueOf(obj.toString()), workbook);
							}

						}
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}

			// Flush byte array to servlet output stream.
			ServletOutputStream out = response.getOutputStream();
			workbook.write(out);
			out.flush();
		}

	}

	private String getFirstCharInUpperCase(String fieldName) {
		return fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
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
