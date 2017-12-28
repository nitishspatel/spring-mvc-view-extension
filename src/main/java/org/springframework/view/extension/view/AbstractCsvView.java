package org.springframework.view.extension.view;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.Date;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.log4j.Logger;
import org.springframework.web.servlet.view.AbstractView;
import org.supercsv.io.CsvBeanWriter;
import org.supercsv.io.ICsvBeanWriter;
import org.supercsv.prefs.CsvPreference;
import org.supercsv.quote.AlwaysQuoteMode;

/**
 * Spring MVC view to process CSV rendering
 * 
 * @author nitish
 *
 */
public abstract class AbstractCsvView extends AbstractView {

	private static final Logger log = Logger.getLogger(AbstractCsvView.class);

	private String fileName;

	public void setFileName(String fileName) {
		this.fileName = fileName;
	}

	protected void prepareResponse(HttpServletRequest request,
			HttpServletResponse response) {


		log.debug("prepareResponse Start");
		String headerKey = "Content-Disposition";
		String headerValue = String.format("attachment; filename=\"%s\"",
				fileName==null?"Report_"+new Date().getTime()+".csv":fileName);
		response.setContentType("text/csv");
//		response.setCharacterEncoding("ISO_8859_1");
		response.setCharacterEncoding("UTF-8");
		response.setHeader(headerKey, headerValue);
	}

	@Override
	protected void renderMergedOutputModel(Map<String, Object> model,
			HttpServletRequest request, HttpServletResponse response)
					throws Exception {
		log.debug("renderMergedOutputModel Start");

		CsvPreference prefs = new CsvPreference.Builder('"', ',', "\r\n").useQuoteMode(new AlwaysQuoteMode()).build();
		
		
//		Writer writer = new OutputStreamWriter(response.getOutputStream(), StandardCharsets.ISO_8859_1);
		Writer writer = new OutputStreamWriter(response.getOutputStream(), StandardCharsets.UTF_8);
		ICsvBeanWriter csvWriter = new CsvBeanWriter(writer, prefs);

		buildCsvDocument(csvWriter, model);
		csvWriter.close();
	}
	
	/**
	 * The concrete view must implement this method.
	 */
	protected abstract void buildCsvDocument(ICsvBeanWriter csvWriter,
			Map<String, Object> model) throws IOException;
}