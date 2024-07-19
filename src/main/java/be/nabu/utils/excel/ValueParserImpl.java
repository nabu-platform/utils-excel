package be.nabu.utils.excel;

import java.math.BigDecimal;
import java.text.ParseException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public class ValueParserImpl implements ValueParser {
	private boolean ignoreErrors, useBigDecimals;

	@Override
	public Object parse(int rowIndex, int cellIndex, Cell cell, FormulaEvaluator evaluator) throws ParseException {
		// this evaluates any formula's in the cell and returns the type of the current cell value (http://poi.apache.org/spreadsheet/eval.html for example)
		// The patch in bug https://issues.apache.org/bugzilla/show_bug.cgi?id=49783 should be applied to the 3.6 release if necessary (IllegalArgumentException for unknown error type 23)
		// also note that there might be a small bug in the "evaluate()" method which does not allow CELL_TYPE_BLANK (will throw an error java.lang.IllegalStateException: Bad cell type (3)), this is a temporary workaround
		// file bug report + patch: https://issues.apache.org/bugzilla/show_bug.cgi?id=49873
		if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
			return null;
		}
		else {
			// the "evaluateInCell" does NOT work on "referenced" formula's, which are created by entering a formula and dragging that cell, at which point the copies refer to the first formula
			// when the evaluateInCell is called once, the formula disappears, so all referenced cells can not compute anymore
			// note sure if this is a feature or a bug
			// filed bug report: https://issues.apache.org/bugzilla/show_bug.cgi?id=49872
			try {
				// this can throw a runtimeexception
				CellValue cellValue = evaluator.evaluate(cell);
				switch(cellValue.getCellType()) {
					case Cell.CELL_TYPE_BLANK: 
						return null;
					case Cell.CELL_TYPE_BOOLEAN: 
						return (Boolean) cellValue.getBooleanValue();
					case Cell.CELL_TYPE_ERROR: 
						if (!ignoreErrors) {
							throw new ParseException("Exception detected at (" + rowIndex + ", " + cellIndex + "): " + FormulaError.forInt(cellValue.getErrorValue()).getString(), rowIndex);
						}
						return FormulaError.forInt(cellValue.getErrorValue()).getString();
					case Cell.CELL_TYPE_FORMULA:
						assert false : "According to the site (http://poi.apache.org/spreadsheet/eval.html) this shouldn't happen due to the evaluate";
					break;
					case Cell.CELL_TYPE_NUMERIC:
						if (DateUtil.isCellInternalDateFormatted(cell) || DateUtil.isCellDateFormatted(cell)) {
							// TODO: cell.getDateCellValue() (if DateUtil is fixed) 
							return DateUtil.getJavaDate(cellValue.getNumberValue());
						}
						else {
							return useBigDecimals ? new BigDecimal(cellValue.getNumberValue()) : cellValue.getNumberValue();
						}
					// add as string
					default: 
						return cellValue.getStringValue();
				}
			}
			// this is thrown when for example referencing an external excel file in a formula
			catch (RuntimeException e) {
				// try to get the current (cached) value if you have ignoreerrors set to true
				if (ignoreErrors) {
					switch (cell.getCachedFormulaResultType()) {
						case Cell.CELL_TYPE_BLANK: 
							throw new RuntimeException("The formula could not be executed but there is no cached value", e);
						case Cell.CELL_TYPE_ERROR: 
							throw new RuntimeException("The formula could not be executed but the cached value is also in error", e);
						case Cell.CELL_TYPE_BOOLEAN: 
							return (Boolean) cell.getBooleanCellValue();
						case Cell.CELL_TYPE_NUMERIC: 
							if (DateUtil.isCellInternalDateFormatted(cell) || DateUtil.isCellDateFormatted(cell)) {
								// TODO: cell.getDateCellValue() (if DateUtil is fixed) 
								return DateUtil.getJavaDate(cell.getNumericCellValue());
							}
							else {
								return useBigDecimals ? new BigDecimal(cell.getNumericCellValue()) : cell.getNumericCellValue();
							}
						default: 
							return cell.getStringCellValue();
					}
				}
				else {
					throw e;
				}
			}
		}
		return null;
	}
}
