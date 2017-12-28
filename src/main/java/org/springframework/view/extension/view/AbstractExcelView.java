package org.springframework.view.extension.view;

import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.web.servlet.view.AbstractView;

/**
 * Spring MVC view to process excel rendering
 * 
 * @author nitish
 *
 */
public abstract class AbstractExcelView extends AbstractView {

	private static final Logger log = Logger.getLogger(AbstractExcelView.class);

	private String fileName;
	public static final String DATEFORMAT="dd-MMM-yyyy";
	public static final String BLANKVAL="";
	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	protected void prepareResponse(HttpServletRequest request,
			HttpServletResponse response) {


		log.debug("prepareResponse Start");
		String headerKey = "Content-Disposition";
		String headerValue = String.format("attachment; filename=\"%s\"",
				fileName==null?"Report_"+  new Date().getTime()+".xlsx":fileName);
		response.setContentType("application/xls");
//		response.setCharacterEncoding("ISO_8859_1");
		response.setCharacterEncoding("UTF-8");
		response.setHeader(headerKey, headerValue);
	}

	@Override
	protected void renderMergedOutputModel(Map<String, Object> model,
			HttpServletRequest request, HttpServletResponse response)
					throws Exception {
		log.debug("renderMergedOutputModel Start");

//		CsvPreference prefs = new CsvPreference.Builder('"', ',', "\r\n").useQuoteMode(new AlwaysQuoteMode()).build();
		
		
//		Writer writer = new OutputStreamWriter(response.getOutputStream(), StandardCharsets.ISO_8859_1);
//		Writer writer = new OutputStreamWriter(response.getOutputStream(), StandardCharsets.UTF_8);
//		ICsvBeanWriter csvWriter = new CsvBeanWriter(writer, prefs);

		buildExcelDocument( model,request,response);
//		writer.close();
	}
	
	/**
	 * Creates the workbook from an existing XLS document.
	 * @param url the URL of the Excel template without localization part nor extension
	 * @param request current HTTP request
	 * @return the XSSFWorkbook
	 * @throws Exception in case of failure
	 */
	protected XSSFWorkbook getTemplateSource(String file, HttpServletRequest request)  {
		
		
		try {
			ClassPathResource inputFile = new ClassPathResource(file);

			// Create the Excel document from the source.
			if (logger.isDebugEnabled()) {
				logger.debug("Loading Excel workbook from " + inputFile);
			}
			return new XSSFWorkbook(inputFile.getInputStream());
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return new XSSFWorkbook();
	}

	/**
	 * Convenient method to obtain the cell in the given sheet, row and column.
	 * <p>Creates the row and the cell if they still doesn't already exist.
	 * Thus, the column can be passed as an int, the method making the needed downcasts.
	 * @param sheet a sheet object. The first sheet is usually obtained by workbook.getSheetAt(0)
	 * @param row the row number
	 * @param col the column number
	 * @return the XSSFCell
	 */
	protected XSSFCell getCell(XSSFSheet sheet, int row, int col) {
		XSSFRow sheetRow = sheet.getRow(row);
		if (sheetRow == null) {
			sheetRow = sheet.createRow(row);
		}
		XSSFCell cell = sheetRow.getCell((short) col);
		if (cell == null) {
			cell = sheetRow.createCell((short) col);
		}
		return cell;
	}

	/**
	 * Convenient method to set a String as text content in a cell.
	 * @param cell the cell in which the text must be put
	 * @param text the text to put in the cell
	 */
	protected void setText(XSSFCell cell, String text) {
		if(StringUtils.isNumeric(text) && text.length()>0){
			cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
			cell.setCellValue(Integer.parseInt(text));
		}else{
			cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			cell.setCellValue(text);
		}
	}
			
	protected void setDoubleVal(XSSFCell cell, String doubleVal) {
		if(StringUtils.isNotBlank(doubleVal) && doubleVal.length()>0){
			cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
			cell.setCellValue(Double.parseDouble(doubleVal));
		}
		else{
			cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			cell.setCellValue(BLANKVAL);
		}
	}
	protected void setFloatVal(XSSFCell cell, String floatVal) {
		if(StringUtils.isNotBlank(floatVal) && floatVal.length()>0){
			cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
			cell.setCellValue(Float.parseFloat(floatVal));
		}else{
			cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			cell.setCellValue(BLANKVAL);
		}
	}
	protected void setIntValue(XSSFCell cell, String intVal) {
		if(StringUtils.isNotBlank(intVal) && intVal.length()>0){
			cell.setCellType(XSSFCell.CELL_TYPE_NUMERIC);
			cell.setCellValue(Integer.parseInt(intVal));
		}else{
			cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			cell.setCellValue(BLANKVAL);
		}
		
	}
	
	protected void setDateValue(XSSFCell cell, String dateVal,XSSFWorkbook wb) {
		if(StringUtils.isNotBlank(dateVal.trim())){
			try {
				Date date = new SimpleDateFormat(DATEFORMAT).parse(dateVal);
				CellStyle cellStyle = wb.createCellStyle();
				CreationHelper createHelper = wb.getCreationHelper();
				cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(DATEFORMAT));
				cell.setCellValue(date);
				cell.setCellStyle(cellStyle);
			} catch (ParseException e1) {
				e1.printStackTrace();
			}
		}else{
			cell.setCellValue(BLANKVAL);
		}
	}
	
	protected void setColor(XSSFCell cell, short color, XSSFWorkbook wb) {
		 // set the color of the cell
       XSSFCellStyle style = wb.createCellStyle();
       
       style.setFillPattern(CellStyle.SOLID_FOREGROUND);
       style.setFillForegroundColor(color);
       style.setFillBackgroundColor(color);
		
		cell.setCellStyle(style);
//		cell.setCellType(XSSFCell.CELL_TYPE_STRING);
//		cell.setCellValue(text);
	}
	
	protected void setColorForDateField(XSSFCell cell, short color, XSSFWorkbook wb) {
		 // set the color of the cell
        XSSFCellStyle style = cell.getCellStyle();
        
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFillForegroundColor(color);
        style.setFillBackgroundColor(color);
		
		cell.setCellStyle(style);
//		cell.setCellType(XSSFCell.CELL_TYPE_STRING);
//		cell.setCellValue(text);
	}
	
	/**
	 * The concrete view must implement this method.
	 */
	protected abstract void buildExcelDocument(
			Map<String, Object> model, HttpServletRequest request, HttpServletResponse response) throws IOException;
}