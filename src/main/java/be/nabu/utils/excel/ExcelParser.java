/*
* Copyright (C) 2015 Alexander Verbruggen
*
* This program is free software: you can redistribute it and/or modify
* it under the terms of the GNU Lesser General Public License as published by
* the Free Software Foundation, either version 3 of the License, or
* (at your option) any later version.
*
* This program is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
* GNU Lesser General Public License for more details.
*
* You should have received a copy of the GNU Lesser General Public License
* along with this program. If not, see <https://www.gnu.org/licenses/>.
*/

package be.nabu.utils.excel;

import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaError;

public class ExcelParser implements Closeable {
	
	// whether you want to use doubles or bigdecimals
	private boolean useBigDecimals;
	private boolean ignoreErrors = true, ignoreHiddenRows = false, ignoreHiddenSheets = false;
	private Integer offsetX, offsetY, maxX, maxY;
	
	private Workbook workbook;
	
	/**
	 * 
	 * @param input
	 * @param fileType
	 * @param password Pass along "null" if no password is used
	 * @return
	 * @throws IOException
	 */
	public ExcelParser(InputStream input, FileType fileType, String password) throws IOException {
		if (password != null) {
			synchronized(Biff8EncryptionKey.class) {
				// set password
				Biff8EncryptionKey.setCurrentUserPassword(password);
				// initialize
				workbook = fileType == FileType.XLSX ? new XSSFWorkbook(input) : new HSSFWorkbook(input);
				// reset password
				Biff8EncryptionKey.setCurrentUserPassword(null);
			}
		}
		else {
			workbook = fileType == FileType.XLS ? new HSSFWorkbook(input) : new XSSFWorkbook(input);
		}
	}
	
	public ExcelParser(Workbook workbook) {
		this.workbook = workbook;
	}
	
	public List<String> getSheetNames(boolean includeHidden) {
		List<String> names = new ArrayList<String>();
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			if (includeHidden || !workbook.isSheetHidden(i)) {
				names.add(workbook.getSheetName(i));
			}
		}
		return names;
	}
	
	public Sheet getSheet(String sheetName, boolean useRegex) throws IOException {
		if (useRegex) {
			boolean found = false;
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				if (workbook.getSheetName(i).matches(sheetName)) {
					if (ignoreHiddenSheets && (workbook.isSheetHidden(i) || workbook.isSheetVeryHidden(i)))
						continue;
					sheetName = workbook.getSheetName(i);
					found = true;
					break;
				}
			}
			if (!found)
				throw new IOException("Could not find a sheet that matches '" + sheetName + "'");
		}
		return workbook.getSheet(sheetName);
	}
	
	public Workbook getWorkbook() {
		return workbook;
	}
	
	public void write(OutputStream output) throws IOException {
		workbook.write(output);
	}
	
	public void replaceAll(Sheet sheet, String regex, String replacement) {
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() == CellType.STRING) {
					String content = cell.getStringCellValue();
					content = content.replaceAll(regex, replacement);
					cell.setCellValue(content);
				}
			}
		}
	}
	
	public void replaceAll(String regex, String replacement) {
		for (int i = 0; i < workbook.getNumberOfSheets(); i++)
			replaceAll(workbook.getSheetAt(i), regex, replacement);
	}
	
	public boolean isHidden(Sheet sheet) {
		int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
		return sheet.getWorkbook().isSheetHidden(sheetIndex) || sheet.getWorkbook().isSheetVeryHidden(sheetIndex);
	}
	
	public List<List<Object>> matrix(Sheet sheet) throws ParseException {
		return matrix(sheet, new ValueParserImpl());
	}
	
	public List<List<Object>> matrix(Sheet sheet, ValueParser parser) throws ParseException {
		List<List<Object>> matrix = new ArrayList<List<Object>>();

		// formula evaluator
		FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		
		// loop over the rows, if you use for (Row row : sheet), you will skip the empty rows, but we need the row location so can't do that
		// sheet.getPhysicalNumberOfRows returns the amount of non-null rows
		// sheet.getLastRowNum() however returns the actual number of the last non-null row which is what we want here because we want to keep the numbering intact
		// note that getLastRowNum() is 0-based
		for (int r = 0; r <= sheet.getLastRowNum(); r++) {
			Row row = sheet.getRow(r);
			
			// if you don't want to show hidden rows and it is hidden, skip
			if (ignoreHiddenRows) {
				// check if a row style is set, however it seems that excel never sets this
//				if (row.isFormatted() && row.getRowStyle().getHidden())
//					continue;
				// if the height of the row is zero, it is considered hidden, this does seem to work with excel > hide functionality
				if (row.getZeroHeight())
					continue;
				// worst case scenario: check all the cells for their hidden toggle
				// however this may impact performance and also does not work with excel it seems so currently the code is disabled
//				boolean isHidden = true;
//				for (int i = 0; i < row.getLastCellNum(); i++) {
//					Cell cell = row.getCell(i);
//					if (cell != null && cell.getCellStyle() != null && !cell.getCellStyle().getHidden()) {
//						isHidden = false;
//						break;
//					}
//				}
//				if (isHidden)
//					continue;
			}
				
			
			// skip this row if it is not within the window
			if ((offsetY != null && row.getRowNum() < offsetY) || (maxY != null && row.getRowNum() > maxY))
				continue;
			
			// initialize a new row
			List<Object> rowList = new ArrayList<Object>();
			
			// only loop if the row is not null
			if (row != null) {
				// loop over the cells
				// WARNING: do not use the for(:) loop as it SKIPS null-valued cells!
				for (int i = 0; i < row.getLastCellNum(); i++) {
					Cell cell = row.getCell(i);
					if (cell == null) {
						rowList.add(null);
						continue;
					}
					// skip this cell if it's not within the window
					if ((offsetX != null && cell.getColumnIndex() < offsetX) || (maxX != null && cell.getColumnIndex() > maxX))
						continue;
					rowList.add(parser.parse(row.getRowNum(), cell.getColumnIndex(), cell, evaluator));
					
//					
//					// this evaluates any formula's in the cell and returns the type of the current cell value (http://poi.apache.org/spreadsheet/eval.html for example)
//					// The patch in bug https://issues.apache.org/bugzilla/show_bug.cgi?id=49783 should be applied to the 3.6 release if necessary (IllegalArgumentException for unknown error type 23)
//					// also note that there might be a small bug in the "evaluate()" method which does not allow CELL_TYPE_BLANK (will throw an error java.lang.IllegalStateException: Bad cell type (3)), this is a temporary workaround
//					// file bug report + patch: https://issues.apache.org/bugzilla/show_bug.cgi?id=49873
//					if (cell.getCellType() == CellType.BLANK)
//						rowList.add(null);
//					else {
//						// the "evaluateInCell" does NOT work on "referenced" formula's, which are created by entering a formula and dragging that cell, at which point the copies refer to the first formula
//						// when the evaluateInCell is called once, the formula disappears, so all referenced cells can not compute anymore
//						// note sure if this is a feature or a bug
//						// filed bug report: https://issues.apache.org/bugzilla/show_bug.cgi?id=49872
//						try {
//							// this can throw a runtimeexception
//							CellValue cellValue = evaluator.evaluate(cell);
//							switch(cellValue.getCellType()) {
//								case BLANK: rowList.add(null); break;
//								case BOOLEAN: rowList.add((Boolean) cellValue.getBooleanValue()); break;
//								case ERROR: 
//									if (!ignoreErrors)
//										throw new ParseException("Exception detected at (" + row.getRowNum() + ", " + i + "): " + FormulaError.forInt(cellValue.getErrorValue()).getString(), row.getRowNum());
//									rowList.add(FormulaError.forInt(cellValue.getErrorValue()).getString()); break;
//								case FORMULA:
//									assert false : "According to the site (http://poi.apache.org/spreadsheet/eval.html) this shouldn't happen due to the evaluate";
//								break;
//								case NUMERIC:
//									if (DateUtil.isCellInternalDateFormatted(cell) || DateUtil.isCellDateFormatted(cell))
//										// TODO: cell.getDateCellValue() (if DateUtil is fixed) 
//										rowList.add(DateUtil.getJavaDate(cellValue.getNumberValue()));
//									else
//										rowList.add(useBigDecimals ? new BigDecimal(cellValue.getNumberValue()) : cellValue.getNumberValue());
//								break;
//								// add as string
//								default: rowList.add(cellValue.getStringValue());
//							}
//						}
//						// this is thrown when for example referencing an external excel file in a formula
//						catch (RuntimeException e) {
//							// try to get the current (cached) value if you have ignoreerrors set to true
//							if (ignoreErrors) {
//								switch (cell.getCachedFormulaResultType()) {
//									case BLANK: throw new RuntimeException("The formula could not be executed but there is no cached value", e);
//									case ERROR: throw new RuntimeException("The formula could not be executed but the cached value is also in error", e);
//									case BOOLEAN: rowList.add((Boolean) cell.getBooleanCellValue()); break;
//									case NUMERIC: 
//										if (DateUtil.isCellInternalDateFormatted(cell) || DateUtil.isCellDateFormatted(cell))
//											// TODO: cell.getDateCellValue() (if DateUtil is fixed) 
//											rowList.add(DateUtil.getJavaDate(cell.getNumericCellValue()));
//										else
//											rowList.add(new BigDecimal(cell.getNumericCellValue()));
//									break;
//									default: rowList.add(cell.getStringCellValue());
//								}
//							}
//							else
//								throw e;
//						}
//
//					}
				}
			}
			// add the row
			matrix.add(rowList);
		}
		return matrix;
	}

	@Deprecated
	public Object[][] parse(Sheet sheet) throws IOException {
		try {
			List<List<Object>> parseAsList = matrix(sheet);
			List<Object[]> matrix = new ArrayList<Object[]>();
			
			for (List<?> parsed : parseAsList) {
				matrix.add(parsed.toArray());
			}
			return (Object[][]) matrix.toArray(new Object[matrix.size()][]);
		}
		catch (ParseException e) {
			throw new IOException(e);
		}
	}

	public boolean isIgnoreErrors() {
		return ignoreErrors;
	}

	public void setIgnoreErrors(boolean ignoreErrors) {
		this.ignoreErrors = ignoreErrors;
	}

	public boolean isIgnoreHiddenRows() {
		return ignoreHiddenRows;
	}

	public void setIgnoreHiddenRows(boolean ignoreHiddenRows) {
		this.ignoreHiddenRows = ignoreHiddenRows;
	}

	public boolean isIgnoreHiddenSheets() {
		return ignoreHiddenSheets;
	}

	public void setIgnoreHiddenSheets(boolean ignoreHiddenSheets) {
		this.ignoreHiddenSheets = ignoreHiddenSheets;
	}

	public Integer getOffsetX() {
		return offsetX;
	}

	public void setOffsetX(Integer offsetX) {
		this.offsetX = offsetX;
	}

	public Integer getOffsetY() {
		return offsetY;
	}

	public void setOffsetY(Integer offsetY) {
		this.offsetY = offsetY;
	}

	public Integer getMaxX() {
		return maxX;
	}

	public void setMaxX(Integer maxX) {
		this.maxX = maxX;
	}

	public Integer getMaxY() {
		return maxY;
	}

	public void setMaxY(Integer maxY) {
		this.maxY = maxY;
	}

	@Override
	public void close() throws IOException {
		workbook.close();
	}

	public boolean isUseBigDecimals() {
		return useBigDecimals;
	}

	public void setUseBigDecimals(boolean useBigDecimals) {
		this.useBigDecimals = useBigDecimals;
	}

}