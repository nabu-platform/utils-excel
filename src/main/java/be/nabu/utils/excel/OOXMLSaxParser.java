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

import java.awt.Point;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.ParseException;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

public class OOXMLSaxParser {
	
	public enum DataType {
		BOOL, ERROR, FORMULA, INLINE_STRING, SHARED_STRING, NUMBER, DATE
	}
	
	public OOXMLSaxParser(InputStream stream) throws InvalidFormatException, IOException {
		this(OPCPackage.open(stream));
	}
	
	public OOXMLSaxParser(File file) throws InvalidFormatException, IOException {
		this(OPCPackage.open(file.getPath(), PackageAccess.READ));
	}
	
	public OOXMLSaxParser(OPCPackage pack) throws InvalidFormatException, IOException {
		this.pack = pack;
	}
	
	public Object[][] parse(String sheet, boolean useRegex, boolean ignoreErrors, Integer offsetX, Integer offsetY, Integer maxX, Integer maxY) throws IOException, OpenXML4JException, SAXException, ParserConfigurationException, ParseException {
		XSSFReader reader = new XSSFReader(pack);
		XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.getSheetsData();
		while (iter.hasNext()) {
			InputStream input = iter.next();
			try {
				if ((useRegex && iter.getSheetName().matches(sheet)) || iter.getSheetName().equals(sheet))
					return parse(input, reader.getStylesTable(), reader.getSharedStringsTable(), ignoreErrors, offsetX, offsetY, maxX, maxY);
			}
			finally {
				input.close();
			}
		}
		throw new SAXException("No sheet found by that name");
	}
	
	private Object[][] parse(InputStream input, StylesTable styles, SharedStrings shared, boolean ignoreErrors, Integer offsetX, Integer offsetY, Integer maxX, Integer maxY) throws SAXException, ParserConfigurationException, IOException, ParseException {
	    XMLReader reader = SAXParserFactory.newInstance().newSAXParser().getXMLReader();
	    SheetHandler handler = new SheetHandler(styles, shared, ignoreErrors, offsetX, offsetY, maxX, maxY);
	    reader.setContentHandler(handler);
	    reader.parse(new InputSource(input));
	    
	    if (handler.isInError())
	    	throw new ParseException(handler.getError(), 0);
	    
	    // transform the object map into an array
	    Map<Integer, Map<Integer, Object>> objects = handler.getObjects();
	    // get the amount of rows
	    int rowCount = Collections.max(objects.keySet());
	    // get the max amount of columns
	    int columnCount = -1;
	    for (Integer row : objects.keySet()) {
	    	int tmp = Collections.max(objects.get(row).keySet());
	    	if (tmp > columnCount)
	    		columnCount = tmp;
	    }
	    Object[][] result = new Object[rowCount + 1][];
	    for (Integer row : objects.keySet()) {
	    	result[row] = new Object[columnCount + 1];
	    	for (Integer column : objects.get(row).keySet())
	    		result[row][column] = objects.get(row).get(column);
	    }
	    
	    return result;
	}
	
	private OPCPackage pack = null;
	
	private class Cell {
		private StylesTable styles = null;
		
		private DataType type = null;
		private Point point = null;
		
		public Cell(Attributes attributes, StylesTable styles) throws SAXException {
			this.styles = styles;
			
			type = getCellType(attributes.getValue("t"), attributes.getValue("s"));
			point = getCellPoint(attributes.getValue("r"));
		}
		
		public DataType getType() {
			return type;
		}
		
		public Point getPoint() {
			return point;
		}
		
		private Point getCellPoint(String cellReference) throws SAXException {
			if (cellReference == null)
				return null;
			int rowStarts = -1;
			int column = -1;
			for (int i = 0; i < cellReference.length(); i++) {
				if (Character.isDigit(cellReference.charAt(i))) {
					rowStarts = i;
					break;
				}
				int character = cellReference.charAt(i);
				column = (column + 1) * 26 + character - 'A';
			}
			if (rowStarts == -1)
				throw new SAXException("Could not find the row indicator of '" + cellReference + "'");
			if (column == -1)
				throw new SAXException("Could not find column indicator of '" + cellReference + "'");
			return new Point(column, new Integer(cellReference.substring(rowStarts)) - 1);
		}
		
		private DataType getCellType(String cellType, String cellStyle) {
			// if none is present, it's a number
			if (cellType == null) {
				// if no cell style defined, it's just a number
				if (cellStyle == null)
					return DataType.NUMBER;
				// otherwise, it's either a formatted number or a formatted date
				else {
					CTXf style = styles.getCellXfAt(new Integer(cellStyle));
					if (DateUtil.isInternalDateFormat((int)style.getNumFmtId()) || DateUtil.isADateFormat((int)style.getNumFmtId(), styles.getNumberFormatAt((short) style.getNumFmtId())))
						return DataType.DATE;
					else
						return DataType.NUMBER;
				}
			}
			else if (cellType.equals("b"))
				return DataType.BOOL;
			else if (cellType.equals("e"))
				return DataType.ERROR;
			else if (cellType.equals("inlineStr"))
				return DataType.INLINE_STRING;
			else if (cellType.equals("s"))
				return DataType.SHARED_STRING;
			else if (cellType.equals("str"))
				return DataType.FORMULA;
			else
				return DataType.NUMBER;
		}
	}
	
	/**
	 * This class can parse a sheet
	 * @author MSV044
	 *
	 */
	private class SheetHandler extends DefaultHandler {
		
		private Integer maxX = null, maxY = null, offsetX = null, offsetY = null;
		
		public SheetHandler(StylesTable styles, SharedStrings shared, boolean ignoreErrors, Integer offsetX, Integer offsetY, Integer maxX, Integer maxY) {
			this.shared = shared;
			this.styles = styles;
			this.maxX = maxX;
			this.maxY = maxY;
			this.offsetX = offsetX;
			this.offsetY = offsetY;
			this.ignoreErrors = ignoreErrors;
		}
		
		/**
		 * This is a table of shared strings that is kept separate from the sheets. The goal here is to save space and processing time for strings that are repeated often.
		 * The cell in question will only hold a reference to this shared table
		 */
		private SharedStrings shared = null;
		private StylesTable styles = null;
		
		/**
		 * Holds the objects
		 */
		private Map<Integer, Map<Integer, Object>> objects = new HashMap<Integer, Map<Integer, Object>>();
		
		/**
		 * Set to true once "v" is found (this is the value of a cell)
		 */
		private boolean isValue = false;
		/**
		 * this contains the actual characters inside the "v" element of a cell
		 */
		private StringBuffer buffer = null;
		
		/**
		 * The last attributes that were encountered
		 */
		private Cell cell = null;
		
		private boolean ignoreErrors = true;
		
		private String exception = null;
		
		public Map<Integer, Map<Integer, Object>> getObjects() {
			return objects;
		}
		
		public boolean isInError() {
			return !ignoreErrors && exception != null && exception.length() != 0;
		}
		
		public String getError() {
			return exception;
		}
		
		public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
			// reset the value
			if (name.equals("v")) {
				buffer = new StringBuffer();
				isValue = true;
			}
			// copy the cell attributes
			else if (name.equals("c"))
				cell = new Cell(attributes, styles);
		}
		
		public void endElement(String uri, String localName, String name) throws SAXException {
			// only act on a closing value tag
			if (name.equals("v")) {
				// stop the value gathering
				isValue = false;
				
				// put the object in the correct position
				if (cell.getPoint() == null)
					throw new SAXException("Could not find the location of a cell");
				
				// we work 0-based, excel works 1-based, hence "-1"
				int row = new Double(cell.getPoint().getY()).intValue();
				int column = new Double(cell.getPoint().getX()).intValue();
				
				// check if it is within the defined range (if any)
				if ((offsetY != null && row < offsetY) || (maxY != null && row > maxY))
					return;
				if ((offsetX != null && column < offsetX) || (maxX != null && column > maxX))
					return;
				
				// translate the object to the proper type
				Object object = null;
				switch (cell.getType()) {
					// TODO: DateUtil.getJavaDate
					case DATE: object = DateUtil.getJavaDate(new Double(buffer.toString())); break;
					case NUMBER: object = new BigDecimal(buffer.toString()); break;
					case BOOL: object = buffer.toString().equals("1"); break;
					case ERROR: 
						exception = "Exception detected, cell unknown: " + buffer.toString();
						object = buffer.toString(); 
					break;
					// this should take the last calculated value
					case FORMULA: object = buffer.toString(); break;
					case SHARED_STRING: object = shared.getItemAt(new Integer(buffer.toString())).getString(); break;
					case INLINE_STRING: object = buffer.toString(); break;
					default: throw new SAXException("Could not process type " + cell.getType());
				}

				// create a list for this row if it does not yet exist
				if (!objects.containsKey(row))
					objects.put(row, new HashMap<Integer, Object>());
				// add the element
				objects.get(row).put(column, object);
			}
		}
		public void characters(char[] ch, int start, int length) throws SAXException {
			if (isValue)
				buffer.append(ch, start, length);
		}
	}
}