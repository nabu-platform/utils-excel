package be.nabu.utils.excel;

import java.text.ParseException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public interface ValueParser {
	public Object parse(int rowIndex, int cellIndex, Cell cell, FormulaEvaluator evaluator) throws ParseException;
}
