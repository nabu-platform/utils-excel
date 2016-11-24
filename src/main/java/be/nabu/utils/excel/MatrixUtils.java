package be.nabu.utils.excel;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.TimeZone;

public class MatrixUtils {
	/**
	 * Rotates the resulting matrix
	 */
	public static Object[][] rotate(Object[][] matrix) {
		// get the most amount of columns (the matrix can have rows of varying length)
		int max = 0;
		for (int i = 0; i < matrix.length; i++) {
			if (matrix[i] != null && matrix[i].length > max)
				max = matrix[i].length;
		}
		Object[][] rotated = new Object[max][matrix.length];
		for (int i = 0; i < matrix.length; i++) {
			if (matrix[i] != null) {
				for (int j = 0; j < matrix[i].length; j++) {
					int iReverse = matrix[i].length - (matrix[i].length - j);
					int jReverse = matrix.length - (matrix.length - i);
					rotated[iReverse][jReverse] = matrix[i][j];
				}
			}
		}
		return rotated;
	}
	
	public static List<List<Object>> rotate(List<List<Object>> matrix) {
		int max = 0;
		for (int i = 0; i < matrix.size(); i++) {
			if (matrix.get(i) != null && matrix.get(i).size() > max) {
				max = matrix.get(i).size();
			}
		}
		List<List<Object>> rotated = new ArrayList<List<Object>>();
		for (int i = 0; i < max; i++) {
			rotated.add(new ArrayList<Object>(Arrays.asList(new Object[matrix.size()])));
		}
		for (int i = 0; i < matrix.size(); i++) {
			if (matrix.get(i) != null) {
				for (int j = 0; j < matrix.get(i).size(); j++) {
					int iReverse = matrix.get(i).size() - (matrix.get(i).size() - j);
					int jReverse = matrix.size() - (matrix.size() - i);
					rotated.get(iReverse).set(jReverse, matrix.get(i).get(j));
				}
			}
		}
		return rotated;
	}
	
	/**
	 * You should almost never set the timezone, look for code comments as to why.
	 * The rule is: the timezone used for formatting HAS to be the same as the timezone used for parsing.
	 * The timezone used for parsing is set to the "local" timezone of the jvm that did the parsing, so in almost all cases you will be formatting on the same jvm, so leave it open.
	 * If timezones are in sync for parser & formatter, the time returned should be the same as can be seen in the excel.
	 * 		-> note that excel does NOT use timezones!
	 */
	public static String[][] toString(Object [][] objects, String dateFormat, TimeZone timezone, Integer precision) {
		if (precision == null) {
			precision = 2;
		}
		if (dateFormat == null)
			dateFormat = "yyyy-MM-dd'T'HH:mm:ss.SSS";
		/**
		 * The timezone for this formatter HAS to be the local time, or put differently, it HAS to be in the same timezone as the code that parsed the excel because poi uses a calendar in the default timezone to parse the excel date.
		 * For some reason around 1st of january 1900, the calendar does weird stuff, for example suppose you parse sequential hours with a calendar in timezone "Europe/Paris" and format the date (using simpledateformat) in "UTC", you get this:
		 * 	1899-12-31 22:00:00 UTC
		 * 	1899-12-31 23:00:00 UTC
		 * 	1900-01-01 00:50:39 UTC
		 * 	1900-01-01 01:50:39 UTC
		 * 
		 * If you parse in "Europe/Brussels" and format in "UTC", you get this (note the missing hour):
		 * 	1899-12-31 22:00:00 UTC
		 * 	1900-01-01 00:00:00 UTC
		 * 	1900-01-01 01:00:00 UTC
		 * 	1900-01-01 02:00:00 UTC
		 * 
		 * Only if the calendar and the formatter use the exact same timezone, does the parsing work 
		 */
		SimpleDateFormat formatter = new SimpleDateFormat(dateFormat);
		if (timezone != null)
			formatter.setTimeZone(timezone);
		String[][] result = new String[objects.length][];
		for (int i = 0; i < objects.length; i++) {
			if (objects[i] != null) {
				result[i] = new String[objects[i].length];
				for (int j = 0; j < objects[i].length; j++) {
					if (objects[i][j] == null)
						result[i][j] = null;
					else if (objects[i][j] instanceof Date)
						result[i][j] = formatter.format(objects[i][j]);
					else if (objects[i][j] instanceof Double && precision != null)
						result[i][j] = String.format("%." + precision + "f", (Double) objects[i][j]);
					else if (objects[i][j] instanceof Float && precision != null)
						result[i][j] = String.format("%." + precision + "f", (Float) objects[i][j]);
					else if (objects[i][j] instanceof BigDecimal && precision != null)
						result[i][j] = String.format("%." + precision + "f", ((BigDecimal) objects[i][j]).doubleValue());
					else
						result[i][j] = objects[i][j].toString();
				}
			}
		}
		return result;
	}
}
