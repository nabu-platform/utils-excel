package be.nabu.utils.excel;

import java.math.BigDecimal;
import java.text.ParseException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public class ValueParserImpl implements ValueParser {
	
	private boolean ignoreErrors = true, useBigDecimals;

	public CellType getCellType(int cellIndex, Cell cell, CellValue value) {
		return value.getCellType();
	}
	
	@Override
	public Object parse(int rowIndex, int cellIndex, Cell cell, FormulaEvaluator evaluator) throws ParseException {
		// this evaluates any formula's in the cell and returns the type of the current cell value (http://poi.apache.org/spreadsheet/eval.html for example)
		// The patch in bug https://issues.apache.org/bugzilla/show_bug.cgi?id=49783 should be applied to the 3.6 release if necessary (IllegalArgumentException for unknown error type 23)
		// also note that there might be a small bug in the "evaluate()" method which does not allow CELL_TYPE_BLANK (will throw an error java.lang.IllegalStateException: Bad cell type (3)), this is a temporary workaround
		// file bug report + patch: https://issues.apache.org/bugzilla/show_bug.cgi?id=49873
		if (cell.getCellType() == CellType.BLANK) {
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
				switch(getCellType(cellIndex, cell, cellValue)) {
					case BLANK: 
						return null;
					case BOOLEAN: 
						return (Boolean) cellValue.getBooleanValue();
					case ERROR: 
						if (!ignoreErrors) {
							throw new ParseException("Exception detected at (" + rowIndex + ", " + cellIndex + "): " + FormulaError.forInt(cellValue.getErrorValue()).getString(), rowIndex);
						}
						return FormulaError.forInt(cellValue.getErrorValue()).getString();
					case FORMULA:
						// this shouldn't happen due to evaluate: http://poi.apache.org/spreadsheet/eval.html
						return null;
					case NUMERIC:
						if (DateUtil.isCellInternalDateFormatted(cell) || DateUtil.isCellDateFormatted(cell)) {
							// TODO: cell.getDateCellValue() (if DateUtil is fixed) 
							return DateUtil.getJavaDate(cellValue.getNumberValue());
						}
						else {
							return useBigDecimals ? new BigDecimal(cellValue.getNumberValue()) : cellValue.getNumberValue();
						}
					// add as string
					default: 
						return cellValue.getCellType() == CellType.STRING ? cellValue.getStringValue() : cellValue.formatAsString();
				}
			}
			// this is thrown when for example referencing an external excel file in a formula
			catch (RuntimeException e) {
				// try to get the current (cached) value if you have ignoreerrors set to true
				if (ignoreErrors) {
					switch (cell.getCachedFormulaResultType()) {
						case BLANK: 
							throw new RuntimeException("The formula could not be executed but there is no cached value", e);
						case ERROR: 
							throw new RuntimeException("The formula could not be executed but the cached value is also in error", e);
						case BOOLEAN: 
							return (Boolean) cell.getBooleanCellValue();
						case NUMERIC: 
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
	}

	public boolean isIgnoreErrors() {
		return ignoreErrors;
	}

	public void setIgnoreErrors(boolean ignoreErrors) {
		this.ignoreErrors = ignoreErrors;
	}

	public boolean isUseBigDecimals() {
		return useBigDecimals;
	}

	public void setUseBigDecimals(boolean useBigDecimals) {
		this.useBigDecimals = useBigDecimals;
	}
	
}
