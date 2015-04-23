package be.nabu.utils.excel;

import java.io.BufferedInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * This class allows you to build excel files starting from a template.
 * For example, you could fill in "%date%" in a cell and pass along a java.util.Date object mapped to the string "date" in the variables of the substitute method.
 * The template would then substitute the variable for its actual value.
 * If a variable that is mapped as input is of the type java.util.List, new rows will be inserted for each element in that list. If you have multiple lists on the same sheet next to each other, it will try to intelligently expand them all.
 * If a variable that is mapped is an array of java.util.Map (not a single map!) the map will be expanded into the variables, this will result in the column being duplicated for each Map in that array. Note however that this can _only_ be done with the last column and only once per sheet! (though that one time, that single column can be duplicated as many times as needed)
 * So for example, if (in the last column) you fill in "%records.date%" in a cell and pass along a map array to a variable called "records" and each map contains a "date", the column will be duplicated for each map and the correct date will be filled in.
 * 
 * Note that there is some basic styling as well. Currently you can only define the "border", for example if you have a list of dates, you know the rows will be repeated, you can then manipulate the border of the generated rows as such:
 * 		%date/border:0101%
 * Another style option is "fit", it allows you to refit generated columns/rows on the content that is added, the only value it has currently is "auto", any other value (or no "fit" option at all) will simply copy/paste the cells so they keep their original size
 * 		%date/border:0101;fit:auto%
 * This means that the date list will be expanded and each row will have its top and bottom border stripped while left and right are left alone. (order: top-right-bottom-left)
 * This does not guarantee a border on the left or right! It only guarantees the top and bottom border will be stripped.
 * Other styling options may be added later (";" separated) 
 * 
 * This should work with both the binary xls and the XML based xlsx
 * 
 * Rather hard to unit test as the resulting files are not binary compatible
 * 
 * The "duplicateAll" variable can be set that when a row/column is duplicated, ALL the cells should be duplicated or only those belonging to the "record" that is in the list that is being looped through.
 * If you set duplicateAll to false, you can still embed constant values that are replicated across columns/rows by writing them like this: %"value"% (WITH the quotes)
 * 
 * Note: this does not work well with formula's (for example with added rows/columns, the formula's are not "updated" to reflect their new position).
 */
public class Template {
	
	public enum Direction {
		VERTICAL, HORIZONTAL
	}
	
	private final static Logger logger = LoggerFactory.getLogger(Template.class);
	
	private File file = null;
	
	public Template(File file) {
		logger.debug("Loading template file '" + file + "'");
		this.file = file;
	}
	
	/**
	 * Transforms the template with the given variables and stores it in the target file
	 * The following variable types are recognized:
	 * 
	 * - java.util.Date
	 * - java.util.Calendar
	 * - java.lang.Integer
	 * - java.lang.Boolean
	 * - java.lang.Double
	 * - a java.util.List of the above
	 * - a java.util.Map array of the above (must be an array)
	 * 
	 * The reason for using map arrays is because excel itself is 2D. At one point we considered introducing more complex objects so as to allow better structuring in java/xml/... (more hierarchical that is)
	 * but then the problem becomes what happens if you start nesting maps? So the input is in a way 2D itself so as to force the user to input it like that.
	 * @param target
	 * @param variables
	 * @param duplicateAll When a map array is exploded, the entire column is duplicated by default. If this is set to false, only those fields that contain complex variable names are duplicated.
	 * 						If this is set to true, the entire column will be duplicated, if set to false, only the fields that contain a complex variable name will be duplicated 
	 * @throws IOException 
	 */
	public void substitute(OutputStream target, Map<String, Object> variables, boolean duplicateAll, Direction direction) throws IOException {
		substitute(target, variables, duplicateAll, direction, false);
	}
	
	@SuppressWarnings("resource")
	public void substitute(OutputStream target, Map<String, Object> variables, boolean duplicateAll, Direction direction, boolean removeNonExistent) throws IOException {
		if (direction == null)
			direction = Direction.VERTICAL;
		logger.debug("Starting substitution in template '" + file + "'");
		InputStream input = new BufferedInputStream(new FileInputStream(file));
		try {
			Workbook wb = null;
			// if it's an xlsx file, use the new workbook
			if (file.getName().matches("(?i)^.*\\.xlsx$"))
				wb = new XSSFWorkbook(input);
			// otherwise, use the old one
			else if (file.getName().matches("(?i)^.*\\.xls$"))
				wb = new HSSFWorkbook(input);
			else
				throw new IOException("Unknown file type, expecting xls or xlsx");
			
			// the pattern to detect variable names
			Pattern pattern = Pattern.compile("%[^%]+%");
			// loop over the sheets
			for (int i = 0; i < wb.getNumberOfSheets(); i++) {
				String sheetName = wb.getSheetName(i);
				logger.debug("Checking sheet '" + sheetName + "'");
				// check name of sheet for variables
				Matcher sheetMatcher = pattern.matcher(sheetName);
				while (sheetMatcher.find()) {
					String variable = sheetMatcher.group().replaceAll("%", "");
					logger.debug("Found variable '" + variable + "' = '{}' in sheetname", variables.get(variable));
					if (variables.containsKey(variable))
						sheetName = sheetName.replaceAll(sheetMatcher.group(), variables.get(variable).toString());
				}
				wb.setSheetName(i, sheetName);
				
				// the "shifts" are a runtime count of which rows have been shifted by how much
				// for example suppose you have two lists next to one another, the first list has 3 elements, so i add 2 rows
				// the second list has 4 elements, i don't want to add 3 new rows, but actually reuse the ones added before, in this case i want to add 1 additional row
				// to be able to do this, i need to keep track of the shifts that have already occurred in this sheet
				Map<Integer, Integer> shifts = new HashMap<Integer, Integer>();

				// get the sheet
				Sheet sheet = wb.getSheetAt(i);
				
				// indicates that processing is still going on. Due to creation of new rows and columns, the counters may get confused, so i basically start over again
				boolean processing = true;
			
				// the actual processing loop, this is jumped back on whenever new rows/columns are created on the fly
				loop: while(processing) {
					logger.trace("Processing loop started");
					// loop over the rows
					for (Row row : sheet) {
						// loop over the cells
						for (Cell cell : row) {
							Matcher matcher = pattern.matcher(cell.toString());
							// loop over any variable names that are present
							while(matcher.find()) {
								// the full variable
								String variable = matcher.group().replaceAll("%", "");
								logger.debug("Found variable reference in excel: " + variable);
								// you can add additional information after a "/"
								String [] parts = variable.split("/");
								// a name can consist of multiple parts
								String [] variableParts = getMappedVariable(parts[0], variables);
								// check if the variable is currently in the variables store as a whole
								if (variables.containsKey(parts[0])) {
									// check the type of the variable
									// a list means i need to possibly add some rows/columns
									if (variables.get(parts[0]) instanceof List) {
										logger.debug("Variable '" + parts[0] + "' is a list");
										List<?> list = (List<?>) variables.get(parts[0]);
										switch(direction) {
											case VERTICAL:
												// the row where the variable is at, is already one row, hence the "-1"
												insertRows(sheet, shifts, row.getRowNum(), list.size() - 1);
												// loop over the values
												for (int j = 0; j < list.size(); j++) {
													// get the row we are filling up
													Row currentRow = sheet.getRow(row.getRowNum() + j);
													// get the cell
													Cell currentCell = currentRow.getCell(cell.getColumnIndex());
													writeValue(currentCell, list.get(j), parts.length > 1 ? parts[1] : null);
												}
												// replace constants in the main row
												if (cell.toString().matches(".*%\"[^\"]+\"%.*"))
													cell.setCellValue(cell.toString().replaceAll("%\"([^\"]+)\"%", "$1"));
											break;
											case HORIZONTAL:
												// write [0] to the current cell
												writeValue(cell, list.get(0), parts.length > 1 ? parts[1] : null);
												// move the cells that are beyond this column
												for (int j = row.getLastCellNum() - 1; j > cell.getColumnIndex(); j--) {
													Cell tmp = row.getCell(j);
													if (tmp != null) {
														Cell newCell = row.createCell(j + list.size() - 1);
														switch (tmp.getCellType()) {
															case Cell.CELL_TYPE_BOOLEAN:
																newCell.setCellValue(tmp.getBooleanCellValue());
															break;
															case Cell.CELL_TYPE_NUMERIC:
																newCell.setCellValue(tmp.getNumericCellValue());
															break;
															case Cell.CELL_TYPE_FORMULA:
																newCell.setCellFormula(tmp.getCellFormula());
															break;
															default:
																newCell.setCellValue(tmp.toString());
														}
														newCell.setCellStyle(tmp.getCellStyle());
														newCell.setCellType(tmp.getCellType());
													}
												}
												// create new cells for the other values
												for (int j = list.size() - 1; j >= 1; j--) {
													Cell newCell = row.createCell(cell.getColumnIndex() + j, cell.getCellType());
													newCell.setCellStyle(cell.getCellStyle());
													writeValue(newCell, list.get(j), parts.length > 1 ? parts[1] : null);
												}
											break;
										}
										// need to start again cause rows may have changed
										// since the variables up till now have been resolved, this should not lead to an infinite loop
										continue loop;
									}
									// if it's a string, it can be a partial match
									else if (variables.get(parts[0]) instanceof String) {
										logger.debug("Variable '" + parts[0] + "' is a string");
										writeValue(cell, cell.toString().replaceAll(matcher.group(), variables.get(parts[0]).toString()), parts.length > 1 ? parts[1] : null);
									}
									// otherwise, just write it as is
									else {
										logger.debug("Variable '" + parts[0] + "' is written as is (" + variables.get(parts[0]) + ")");
										writeValue(cell, variables.get(parts[0]), parts.length > 1 ? parts[1] : null);
									}
								}
								// if the variable itself is not in the store, it may be a "." separated name, in which case i may need to explode an array of maps
								// e.g. "records.name" could indicate multiple records
								else if (variableParts != null && variableParts.length > 1 && variables.containsKey(variableParts[0])) {
									// an array of maps can be exploded onto the variable space
									if (variables.get(variableParts[0]) instanceof Map[]) {
										logger.debug("Variable '" + variableParts[0] + "' is a map array");
										Map<?,?>[] maps = (Map[]) variables.get(variableParts[0]);
										switch (direction) {
											case VERTICAL:
												// duplicate the last column a number of times
												if (maps.length > 1)
													duplicateColumn(sheet, maps.length, parts.length > 1 ? parts[1] : null, cell.getColumnIndex(), variableParts[0], duplicateAll);
												for (int j = 0; j < maps.length; j++) {
													// we need to create a new column for each record (except of course the first column, which is the template)
//													if (j > 0) {
//														duplicateColumn(sheet, j, parts.length > 1 ? parts[1] : null, cell.getColumnIndex());
//														out.println("Duplicated column: " + j + "." + parts[0]);
//													}
													logger.debug("Exploding map " + j + " for " + variableParts[0]);
													// explode all the variables in the map in a way: "<counter>.<name>"
													for (Object key : maps[j].keySet()) {
														String name = j + "." + variableParts[0] + "." + key.toString();
														logger.debug("\t" + name + ": " + maps[j].get(key));
														variables.put(name, maps[j].get(key));
													}
												}
												// all the newly duplicated columns will have automatically updated the names of the variables to incorporate the exploded name
												// however the original column is not yet updated, do that now
												// update the original column to use prefixes (now that all the copies have been made)
												for (Row tmp : sheet) {
													if (tmp.getCell(cell.getColumnIndex()) != null) {
														Cell tmpCell = tmp.getCell(cell.getColumnIndex());
														// check if it contains a variable
//														if (tmpCell.toString().matches(".*%[^%]+%.*"))
//															tmpCell.setCellValue(tmpCell.toString().replaceAll("%([^%]+%)", "%0.$1"));
														// check if it contains a complex variable of the same type as the map
														// constants
														if (tmpCell.toString().matches(".*%\"[^\"]+\"%.*"))
															tmpCell.setCellValue(tmpCell.toString().replaceAll("%\"([^\"]+)\"%", "$1"));
														else if (tmpCell.toString().matches(".*%" + variableParts[0] + "\\.[^%]+%.*"))
															tmpCell.setCellValue(tmpCell.toString().replaceAll("%([^%]+%)", "%0.$1"));
													}
												}
											break;
											case HORIZONTAL:
												int blocksize = getBlocksize(sheet, row.getRowNum(), variableParts[0]);
												// duplicate the row a few times
												if (maps.length > 1)
													duplicateRowBlock(sheet, maps.length, parts.length > 1 ? parts[1] : null, row.getRowNum(), blocksize, variableParts[0], duplicateAll);
												// the "first" row block, so the original, is not yet updated, update it
												for (int j = 0; j < blocksize; j++) {
													Row blockRow = sheet.getRow(row.getRowNum() + j);
													for (Cell tmp : blockRow) {
														// constants
														if (tmp.toString().matches(".*%\"[^\"]+\"%.*"))
															tmp.setCellValue(tmp.toString().replaceAll("%\"([^\"]+)\"%", "$1"));
														else if (tmp.toString().matches("^.*%" + variableParts[0] + "\\.[^%]+%.*$"))
															tmp.setCellValue(tmp.toString().replaceAll("^(.*%)(" + variableParts[0] + "\\.[^%]+%.*)$", "$1" + "0" + ".$2"));
													}
												}
												// explode the maps
												for (int j = 0; j < maps.length; j++)
													explode(variables, maps[j], j + "." + variableParts[0]);
											break;
										}
										// remove the map
										// 2013-12-19: don't remove the variable, this way it can be used multiple times
										// this could be optimized so we don't explode it every time.
//										variables.remove(variableParts[0]);
										// the sheet has changed, restart processing
										continue loop;
									}
									else
										logger.warn("The complex variable '" + variableParts[0] + "' is not a map");
								}
								else if (removeNonExistent) {
									if (cell.toString().equals(matcher.group())) {
										logger.debug("The non-existent variable is replaced with a blank cell");
										cell.setCellType(Cell.CELL_TYPE_BLANK);
									}
									else {
										logger.debug("The non-existent variable is a partial string, replace with an empty string");
										writeValue(cell, cell.toString().replaceAll(matcher.group(), ""), null);
									}
								}
								else
									logger.warn("Could not find variable '" + parts[0] + "'");
							}
						}
					}
					// stop processing
					break loop;
				}
			}
			wb.write(target);
		}
		finally {
			input.close();
		}
	}
	
	private void explode(Map<String, Object> variables, Map<?, ?> source, String name) {
		for (Object key : source.keySet()) {
			String childName = name + "." + key.toString();
			if (source.get(key) instanceof Map)
				explode(variables, (Map<?, ?>) source.get(key), childName);
			else {
				logger.debug("\t" + childName + ": " + source.get(key));
				variables.put(childName, source.get(key));
			}
		}
	}
	
	public void substitute(OutputStream target, Map<String, Object> variables, boolean duplicateAll) throws IOException {
		substitute(target, variables, duplicateAll, null);
	}
	
	public void substitute(OutputStream target, Map<String, Object> variables) throws IOException {
		substitute(target, variables, true);
	}
	
	public void substitute(File target, Map<String, Object> variables, boolean duplicateAll) throws IOException {
		OutputStream output = new FileOutputStream(target);
		try {
			substitute(output, variables, duplicateAll);
		}
		finally {
			output.close();
		}
	}
	
	public void substitute(File target, Map<String, Object> variables) throws IOException {
		substitute(target, variables, true);
	}
	
	public byte[] substitute(Map<String, Object> variables, boolean duplicateAll) throws IOException {
		return substitute(variables, duplicateAll, null);
	}
	public byte[] substitute(Map<String, Object> variables, boolean duplicateAll, Direction direction, boolean removeNonExistent) throws IOException {
		ByteArrayOutputStream output = new ByteArrayOutputStream();
		substitute(output, variables, duplicateAll, direction, removeNonExistent);
		return output.toByteArray();
	}
	public byte[] substitute(Map<String, Object> variables, boolean duplicateAll, Direction direction) throws IOException {
		return substitute(variables, duplicateAll, direction, false);
	}
	
	public byte[] substitute(Map<String, Object> variables) throws IOException {
		return substitute(variables, true);
	}
	
	/**
	 * If the target name is not in the variables array as is, it can be part of a map that needs to be exploded. 
	 * This service returns an array of size 2: [name_of_map_variable, name_of_variable_in_map]
	 * @return
	 */
	private String [] getMappedVariable(String name, Map<String, Object> variables) {
		String [] parts = name.split("\\.");
		int i = 0;
		String key = "";
		do {
			key += (key.equals("") ? "" : ".") + parts[i];
			if (variables.containsKey(key))
				return new String[] { key, name.replaceFirst(Pattern.quote(key), "") };
			i++;
		} while(i < parts.length - 1);
		return null;
	}
	
	private void writeValue(Cell cell, Object value, String style) {
		// the type of the value determines how to write the value
		if (value instanceof Integer)
			cell.setCellValue((Integer) value);
		else if (value instanceof Double)
			cell.setCellValue((Double) value);
		else if (value instanceof Boolean)
			cell.setCellValue((Boolean) value);
		else if (value instanceof Date) {
			// if it is within the first day (1st of january 1970 by default), let's map it to a fractional time
			if (((Date) value).getTime() < 86400000)
				cell.setCellValue(((Date) value).getTime() / 86400000.0);
			else
				cell.setCellValue((Date) value);
		}
		else if (value instanceof Calendar)
			cell.setCellValue((Calendar) value);
		else if (value == null)
			cell.setCellType(Cell.CELL_TYPE_BLANK);
		else
			cell.setCellValue(value.toString());
		applyStyle(cell, style);
		// i can not reliably get the cell type from a template, so the following code is deprecated until i can
		// it would be cleaner to format the output as dictated in the template as opposed to what the actual input type is
//		switch(cell.getCellType()) {
//			case Cell.CELL_TYPE_NUMERIC:
//				out.println("Writing a number...");
//				if (value instanceof Integer)
//					cell.setCellValue((Integer) value);
//				else if (value instanceof Double)
//					cell.setCellValue((Double) value);
//				else {
//					// try to parse it as an integer
//					try {
//						cell.setCellValue(new Integer(value.toString()));
//					}
//					catch (NumberFormatException e) {
//						// try to parse it as a double, any exceptions are thrown
//						cell.setCellValue(new Double(value.toString()));
//					}
//				}				
//			break;
//			case Cell.CELL_TYPE_BOOLEAN:
//				if (value instanceof Boolean)
//					cell.setCellValue((Boolean) value);
//				else
//					cell.setCellValue(new Boolean(value.toString()));
//			break;
//			default:
//				cell.setCellValue(value.toString());
//		}
	}
	
	private void insertRows(Sheet sheet, Map<Integer, Integer> shifts, int row, int amount) {
		// the offset of already shifted rows, if we shift at the beginning, earlier data may shift downwards
		int offset = 0;

		// previous shifts of this area (by other columns) should be taken into consideration
		if (shifts.containsKey(row)) {
			// the offset at which the new rows will be inserted
			offset = shifts.get(row);
			// less rows have to be inserted
			amount -= offset;
		}
		
		// if there is still a shift to be done, do so
		if (amount > 0) {
			logger.debug("Inserting " + amount + " new rows after row " + row);
			// shift it
			// the physical number of rows is...not always accurate though i have no clue why
//			sheet.shiftRows(row + 1 + offset, sheet.getPhysicalNumberOfRows(), amount);
			sheet.shiftRows(row + 1 + offset, sheet.getLastRowNum(), amount);
			// now copy the previous row styles
			for (int i = 0; i < amount; i++) {
				Row newRow = sheet.createRow(row + i + 1 + offset);
				for (Cell cell : sheet.getRow(row)) {
					Cell newCell = newRow.createCell(cell.getColumnIndex());
					newCell.setCellStyle(cell.getCellStyle());
					// duplicate constants
					if (cell.toString().matches(".*%\"[^\"]+\"%.*"))
						newCell.setCellValue(cell.toString().replaceAll("%\"([^\"]+)\"%", "$1"));
					// otherwise leave empty
					else
						newCell.setCellType(Cell.CELL_TYPE_BLANK);
				}
			}
		
			// register the shift in the provided map so following shifts are also done properly
			if (shifts.containsKey(row))
				shifts.put(row, amount + shifts.get(row));
			else
				shifts.put(row, amount);
			
			// shift any other rows that are impacted (so that lay beneath this row)
			boolean processing = true;
			
			// keeps track of which rows have already been shifted to prevent double shifts
			List<Integer> shifted = new ArrayList<Integer>();
			// this is a default (though not pretty) construct to loop through a list and change it without using an iterator
			loop : while(processing) {
				for (int key : shifts.keySet()) {
					if (!shifted.contains(key) && key > row) {
						if (shifts.containsKey(key + amount))
							throw new RuntimeException("Unexpected condition: row appears twice");
						shifts.put(key + amount, shifts.get(key));
						shifts.remove(key);
						// add it to the shifted
						shifted.add(key + amount);
						// the map has changed, start looping again
						continue loop;
					}
				}
				// if no hits were found, stop the loop
				break loop;
			}
		}
	}
	
	private Map<String, String> parseStyle(String style) {
		Map<String, String> map = new HashMap<String, String>();
		String [] parts = style.split(";");
		for (String part : parts) {
			String [] subParts = part.split(":");
			map.put(subParts[0], subParts.length > 1 ? subParts[1] : null);
		}
		return map;
	}
	
	/**
	 * Duplicates the last column a number of times
	 * Due to excel/library constrictions it is hard to duplicate any column other then the last one (no column shift!)
	 * Also note that functions are NOT updated, so if you have a SUM(C) in excel and this is copied to D, the function in the D column will also say SUM(C)!
	 * @param sheet
	 * @param count The total amount of columns you need (this INCLUDES the one in the template, so the code will do -1 to compensate)
	 * @param style additional style information
	 * @param columnIndex The index of the column that was requested for duplication (each row will only be copied if the columnIndex matches the last one
	 * @param complexName If this is set to "null", all the columns in the row will be moved. If it is set to a value, only the rows who's columns (@ columIndex) have the complex name will be moved.
	 */
	private void duplicateColumn(Sheet sheet, int recordCount, String style, int columnIndex, String complexName, boolean duplicateAll) {
		logger.debug("Moving columns after '" + columnIndex + "' " + recordCount + " to the right" + (complexName != null ? " (where complex variable is '" + complexName + "')" : " (all)"));
		// for the count of the records to be inserted
		// for each row, copy the cell
		for (Row row : sheet) {
			// if you give a complex name, only cells that contain a variable with such a complex name are copied
			// otherwise if you have two complex variables in the same column, they will both get copied
			if (!duplicateAll && complexName != null && (row.getCell(columnIndex) != null && !row.getCell(columnIndex).toString().matches("^.*%" + complexName + "\\.[^%]+%.*$") && !row.getCell(columnIndex).toString().matches("^.*%\"[^\"]+\"%.*$")))
				continue;
			// move the cells that are beyond this column
			for (int i = row.getPhysicalNumberOfCells() - 1; i > columnIndex; i--) {
				Cell cell = row.getCell(i);
				if (cell != null) {
					Cell newCell = row.createCell(i + recordCount - 1);
					switch (cell.getCellType()) {
						case Cell.CELL_TYPE_BOOLEAN:
							newCell.setCellValue(cell.getBooleanCellValue());
						break;
						case Cell.CELL_TYPE_NUMERIC:
							newCell.setCellValue(cell.getNumericCellValue());
						break;
						case Cell.CELL_TYPE_FORMULA:
							newCell.setCellFormula(cell.getCellFormula());
						break;
						default:
							newCell.setCellValue(cell.toString());
					}
					newCell.setCellStyle(cell.getCellStyle());
					newCell.setCellType(cell.getCellType());
				}
			}
			// copy the indicated cell
			Cell cell = row.getCell(columnIndex);
			// the last cell can be null if the line is empty for instance, this may however lead to unknown fringe cases
			// only copy it if it's not null
			if (cell != null) {
				logger.trace("Copying cell with value " + cell.toString() + "...");
				// copy it for "count" times, the output will start with 1
				for (int i = 1; i < recordCount; i++) {
					// create a new cell at the end
					Cell newCell = row.createCell(columnIndex + i);
					// copy the style information
					newCell.setCellStyle(cell.getCellStyle());
					// copy the cell type
					newCell.setCellType(cell.getCellType());
					// set the value
					switch (cell.getCellType()) {
						case Cell.CELL_TYPE_BOOLEAN:
							newCell.setCellValue(cell.getBooleanCellValue());
						break;
						case Cell.CELL_TYPE_NUMERIC:
							newCell.setCellValue(cell.getNumericCellValue());
						break;
						case Cell.CELL_TYPE_FORMULA:
							newCell.setCellFormula(cell.getCellFormula());
						break;
						default:
							newCell.setCellValue(cell.toString());
					}
					// replace constants
					if (newCell.toString().matches(".*%\"[^\"]+\"%.*"))
						newCell.setCellValue(newCell.toString().replaceAll("%\"([^\"]+)\"%", "$1"));
					// replace variables (named or if no name is given all)
					else if ((complexName == null && newCell.toString().matches(".*%[^%]+%.*")) || (newCell.toString().matches(".*%" + complexName + "\\.[^%]+%.*"))) {
						// this quick fix does not apply to columns because each column must be auto-sized
//						if (i < recordCount - 1)
//							newCell.setCellValue(newCell.toString().replaceAll("%([^%]+%)", "%" + i + ".$1").replaceAll("(/|;)fit:auto", ""));
//						else
							newCell.setCellValue(newCell.toString().replaceAll("%([^%]+%)", "%" + i + ".$1"));
					}
				}
			}
		}
	}
	
	@Deprecated	// use the duplicateBlockRow, it allows for block based duplication
	@SuppressWarnings("unused")
	private void duplicateRow(Sheet sheet, int recordCount, String style, int rowIndex, String complexName, boolean duplicateAll) {
		// only do something if the record count is bigger then 1, there is already one row available
		if (recordCount > 1) {
			// shift the rows starting with the first row after the one you indicated
			// this allows me to insert new rows behind that one
			// only shift the rows behind the current row if there are actually any rows
			if (rowIndex < sheet.getLastRowNum())
				sheet.shiftRows(rowIndex + 1, sheet.getLastRowNum(), recordCount - 1);
			// the recordCount includes the existing row
			for (int i = 1; i < recordCount; i++) {
				Row row = sheet.createRow(rowIndex + i);
				for (Cell cell : sheet.getRow(rowIndex)) {
					if (cell == null)
						continue;
					if (!duplicateAll && complexName != null && !(cell.toString().matches("^.*%" + complexName + "\\.[^%]+%.*$")) && !cell.toString().matches("^.*%\"[^\"]+\"%.*$"))
						continue;
					Cell newCell = row.createCell(cell.getColumnIndex());
					newCell.setCellStyle(cell.getCellStyle());
					// replace constants
					if (cell.toString().matches(".*%\"[^\"]+\"%.*"))
						newCell.setCellValue(cell.toString().replaceAll("%\"([^\"]+)\"%", "$1"));
					// if there is a variable in the cell, prefix it with the index
					else if ((complexName == null && cell.toString().matches(".*%[^%]+%.*")) || cell.toString().matches(".*%" + complexName + "\\.[^%]+%.*")) {
						// quick fix for a problem we encountered: we had like 4 sheets with 15 columns each and hundreds of rows. The "fit:auto" property was copied into every cell which was slow as hell
						// removing the fix auto reduced the convert time from 8 minutes to 5 seconds.
						// this bit of code will check that unless you are doing the last row, the fit:auto should NOT be copied
						// doing a fit does the entire column anyway
						if (i < recordCount - 1)
							newCell.setCellValue(cell.toString().replaceAll("%([^%]+%)", "%" + i + ".$1").replaceAll("(/|;)fit:auto", ""));
						else
							newCell.setCellValue(cell.toString().replaceAll("%([^%]+%)", "%" + i + ".$1"));
					}
					else {
						switch (cell.getCellType()) {
							case Cell.CELL_TYPE_BOOLEAN: newCell.setCellValue(cell.getBooleanCellValue()); break;
							case Cell.CELL_TYPE_ERROR: newCell.setCellValue(cell.getErrorCellValue()); break;
							case Cell.CELL_TYPE_FORMULA: newCell.setCellFormula(cell.getCellFormula()); break;
							case Cell.CELL_TYPE_NUMERIC: newCell.setCellValue(cell.getNumericCellValue()); break;
							case Cell.CELL_TYPE_BLANK: newCell.setCellType(Cell.CELL_TYPE_BLANK); break;
							default: newCell.setCellValue(cell.toString());
						}
					}
				}
			}
		}
	}
	
	private void duplicateRowBlock(Sheet sheet, int recordCount, String style, int rowIndex, int blocksize, String complexName, boolean duplicateAll) {
		System.out.println("Duplicating row block: " + rowIndex + " > " + blocksize);
		// only duplicate if there is more than 1 record
		if (recordCount > 1) {
			// shift rows after the block, if any
			if (rowIndex + (blocksize - 1) < sheet.getLastRowNum())
				sheet.shiftRows(rowIndex + blocksize, sheet.getLastRowNum(), (recordCount * blocksize) - 1);
			// not the first record, it uses the rows already available in the template
			for (int i = 1; i < recordCount; i++) {
				for (int j = 0; j < blocksize; j++) {
					Row row = sheet.createRow(rowIndex + (i * blocksize) + j);
					for (Cell cell : sheet.getRow(rowIndex + j)) {
						if (cell == null)
							continue;
						if (!duplicateAll && complexName != null && !(cell.toString().matches("^.*%" + complexName + "\\.[^%]+%.*$")) && !cell.toString().matches("^.*%\"[^\"]+\"%.*$"))
							continue;
						Cell newCell = row.createCell(cell.getColumnIndex());
						newCell.setCellStyle(cell.getCellStyle());
						// replace constants
						if (cell.toString().matches(".*%\"[^\"]+\"%.*"))
							newCell.setCellValue(cell.toString().replaceAll("%\"([^\"]+)\"%", "$1"));
						// if there is a variable in the cell, prefix it with the index
						else if ((complexName == null && cell.toString().matches(".*%[^%]+%.*")) || cell.toString().matches(".*%" + complexName + "\\.[^%]+%.*")) {
							// quick fix for a problem we encountered: we had like 4 sheets with 15 columns each and hundreds of rows. The "fit:auto" property was copied into every cell which was slow as hell
							// removing the fix auto reduced the convert time from 8 minutes to 5 seconds.
							// this bit of code will check that unless you are doing the last row, the fit:auto should NOT be copied
							// doing a fit does the entire column anyway
							if (i < recordCount - 1)
								newCell.setCellValue(cell.toString().replaceAll("%([^%]+%)", "%" + i + ".$1").replaceAll("(/|;)fit:auto", ""));
							else
								newCell.setCellValue(cell.toString().replaceAll("%([^%]+%)", "%" + i + ".$1"));
						}
						else {
							switch (cell.getCellType()) {
								case Cell.CELL_TYPE_BOOLEAN: newCell.setCellValue(cell.getBooleanCellValue()); break;
								case Cell.CELL_TYPE_ERROR: newCell.setCellValue(cell.getErrorCellValue()); break;
								case Cell.CELL_TYPE_FORMULA: newCell.setCellFormula(cell.getCellFormula()); break;
								case Cell.CELL_TYPE_NUMERIC: newCell.setCellValue(cell.getNumericCellValue()); break;
								case Cell.CELL_TYPE_BLANK: newCell.setCellType(Cell.CELL_TYPE_BLANK); break;
								default: newCell.setCellValue(cell.toString());
							}
						}
					}
				}
			}
		}
	}
	
	private int getBlocksize(Sheet sheet, int rowIndex, String complexName) {
		int blocksize = 0;
		// any lines that are empty at the end are counted with the block size but should only really count if followed by a variable line
		int trailingEmptyLines = 0;
		for (int i = rowIndex; i <= sheet.getLastRowNum(); i++) {
			boolean hasComplex = false;
			boolean hasOtherVariables = false;
			for (Cell cell : sheet.getRow(i)) {
				if (cell.toString().matches("%" + complexName + "\\.[^%]+%")) {
					hasComplex = true;
					break;
				}
				else if (cell.toString().matches("%[^%]+%"))
					hasOtherVariables = true;
			}
			if (hasComplex) {
				trailingEmptyLines = 0;
				blocksize++;
			}
			else if (!hasOtherVariables) {
				trailingEmptyLines++;
				blocksize++;
			}
			else
				break;
		}
		return blocksize - trailingEmptyLines;
	}
	
	private void applyStyle(Cell cell, String style) {
		if (style != null) {
			logger.debug("Applying style '" + style + "' on " + cell.getRowIndex() + ", " + cell.getColumnIndex());
			Map<String, String> map = parseStyle(style);
			if (map.containsKey("border")) {
				if (map.get("border").substring(0, 1).equals("0"))
					cell.getCellStyle().setBorderTop(CellStyle.BORDER_NONE);
				if (map.get("border").substring(1, 2).equals("0"))
					cell.getCellStyle().setBorderRight(CellStyle.BORDER_NONE);							
				if (map.get("border").substring(2, 3).equals("0"))
					cell.getCellStyle().setBorderBottom(CellStyle.BORDER_NONE);
				if (map.get("border").substring(3, 4).equals("0"))
					cell.getCellStyle().setBorderLeft(CellStyle.BORDER_NONE);
			}
			if (map.containsKey("fit")) {
				if (map.get("fit").equals("auto"))
					cell.getSheet().autoSizeColumn(cell.getColumnIndex());
			}
		}
	}
}