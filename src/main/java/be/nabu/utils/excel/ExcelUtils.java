package be.nabu.utils.excel;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.SAXException;

public class ExcelUtils {
	/**
	 * Excel does NOT use timezones, however poi by default uses the calendar in local time to parse a date. 
	 * For some reason around 1st of january 1900, the calendar does weird stuff, for example suppose you parse sequential hours with a calendar in timezone "Europe/Paris" and format the date (using simpledateformat) in "UTC", you get this:
	 * 	1899-12-31 22:00:00 UTC
	 * 	1899-12-31 23:00:00 UTC
	 * 	1900-01-01 00:50:39 UTC
	 * 	1900-01-01 01:50:39 UTC
	 * 
	 * It works if the calendar and formatter are in the same timezone, this method is the same as org.apache.poi.ss.usermodel.DateUtil.getJavaDate() except that it allows you to set the timezone for the calendar (in case you can not control the formatter)
	 */
	public static Date getJavaDate(double date, TimeZone timezone) {
		try {
			// excel (in an effort to stay compatible with lotus notes) introduced a bug where the year 1900 is a leap year, which is true for julian calendars
			// however in the gregorian calendar all centuries not divisible by 400 are NOT leap years, so 1900 is by current definition not a leap year
			// this means that any date beyond 1st of march should be decreased by 1 day to process
			// note that 1st of march is 60 days after the excel epoch defined below
			if (date >= 60)
				date--;
			// java/unix/utc epoch is at January 1, 1970, 00:00:00 UTC
			// excel offset is December 31, 1899, 00:00:00 UTC
			SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
			formatter.setTimeZone(timezone);
			Date epoch = formatter.parse("1970-01-01 00:00:00");
			Date excel = formatter.parse("1899-12-31 00:00:00");
			return new Date((long) (date * (1000L*60*60*24) - (epoch.getTime() - excel.getTime())));
		}
		catch (java.text.ParseException e) {
			throw new RuntimeException(e);
		}
	}
	
	/**
	 * Writes a matrix to a sheet in an excel file
	 */
	public static void write(OutputStream output, Object[][] matrix, String sheetName, FileType fileType, String dateFormat) throws IOException {
		if (dateFormat == null)
			dateFormat = "yyyy-MM-ddTHH:mm:ss";
		Workbook workbook = fileType == FileType.XLS ? new HSSFWorkbook() : new XSSFWorkbook();
		try {
			Sheet sheet = workbook.createSheet(sheetName);
			for (int i = 0; i < matrix.length; i++) {
				Row row = sheet.createRow(i);
				if (matrix[i] != null) {
					for (int j = 0; j < matrix[i].length; j++) {
						Cell cell = row.createCell(j);
						if (matrix[i][j] == null)
							cell.setCellType(Cell.CELL_TYPE_BLANK);
						else if (matrix[i][j] instanceof Number) {
							cell.setCellType(Cell.CELL_TYPE_NUMERIC);
							cell.setCellValue(((Number) matrix[i][j]).doubleValue());
						}
						else if (matrix[i][j] instanceof Boolean) {
							cell.setCellType(Cell.CELL_TYPE_BOOLEAN);
							cell.setCellValue(((Boolean) matrix[i][j]).booleanValue());
						}
						else if (matrix[i][j] instanceof Date) {
							cell.setCellType(Cell.CELL_TYPE_NUMERIC);
							cell.setCellValue((Date) matrix[i][j]);
							CellStyle style = workbook.createCellStyle();
							DataFormat formatter = workbook.createDataFormat();
							style.setDataFormat(formatter.getFormat(dateFormat));
							cell.setCellStyle(style);
						}
						else if (matrix[i][j] instanceof Calendar) {
							cell.setCellType(Cell.CELL_TYPE_NUMERIC);
							cell.setCellValue((Calendar) matrix[i][j]);						
						}
						else if (matrix[i][j] instanceof RichTextString) {
							cell.setCellType(Cell.CELL_TYPE_STRING);
							cell.setCellValue((RichTextString) matrix[i][j]);						
						}
						// formula
						else if (matrix[i][j] instanceof String && ((String) matrix[i][j]).startsWith("=")) {
							cell.setCellType(Cell.CELL_TYPE_FORMULA);
							// can not start with a "="
							cell.setCellFormula(((String) matrix[i][j]).substring(1));
						}
						else {
							cell.setCellType(Cell.CELL_TYPE_STRING);
							cell.setCellValue(matrix[i][j].toString());
						}
					}
				}
			}
			workbook.write(output);
		}
		finally {
			workbook.close();
		}
	}
	
	public static Object[][] parseSAX(InputStream input, String sheetName, Boolean useRegex, Boolean ignoreErrors, Integer offsetX, Integer offsetY, Integer maxX, Integer maxY) throws IOException, ParseException {
		if (ignoreErrors == null)
			ignoreErrors = true;

		if (useRegex == null)
			useRegex = false;
		try {
			OOXMLSaxParser parser = new OOXMLSaxParser(input);
			return parser.parse(sheetName, useRegex, ignoreErrors, offsetX, offsetY, maxX, maxY);
		}
		catch (OpenXML4JException e) {
			throw new RuntimeException(e);
		}
		catch (SAXException e) {
			throw new RuntimeException(e);
		}
		catch (ParserConfigurationException e) {
			throw new RuntimeException(e);
		}
	}
}
