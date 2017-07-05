package uk.co.terminological.excel;

import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

import uk.co.terminological.datatypes.DateFormatString;

public class ExcelCell {
	
	Cell cell;
	
	protected ExcelCell(Cell cell) {
		this.cell = cell;
	}

	/**
	 * @return the string value of a HSSP cell, as displayed by Excel, or null if blank
	 */
	public String toString() {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return DateFormatString.ISO_DATE_TIME_ZONE_FORMAT.format(cell.getDateCellValue());
			} else {
				return Double.toString(cell.getNumericCellValue());
			}
		case Cell.CELL_TYPE_BOOLEAN:
			return cell.getBooleanCellValue() ? "true" : "false";
		case Cell.CELL_TYPE_FORMULA:
			Workbook wb = cell.getSheet().getWorkbook();
			FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
			CellValue cellValue = evaluator.evaluate(cell);
			switch (cellValue.getCellType()) {
			case Cell.CELL_TYPE_BOOLEAN:
				return cellValue.getBooleanValue() ? "true" : "false";
			case Cell.CELL_TYPE_NUMERIC:
				return Double.toString(cellValue.getNumberValue());
			case Cell.CELL_TYPE_STRING:
				return cellValue.getStringValue();
			case Cell.CELL_TYPE_BLANK:
				break;
			case Cell.CELL_TYPE_ERROR:
				break;
			case Cell.CELL_TYPE_FORMULA: 
				break;
			}	
			return null;
		default:
			return null;
		}
	}
	
	public void setValue(Object o) {
		if (o == null) cell.setCellType(Cell.CELL_TYPE_BLANK);
		else if (o instanceof Boolean || o.getClass().equals(Boolean.TYPE)) {
			cell.setCellValue((Boolean) o);
		} else if (o instanceof Double || o.getClass().equals(Double.TYPE)
				|| o instanceof Integer || o.getClass().equals(Integer.TYPE)
				|| o instanceof Float || o.getClass().equals(Float.TYPE)
				) {
			cell.setCellValue(Double.parseDouble(o.toString()));
		} else if (o instanceof Date) {
			cell.setCellValue((Date) o);
		} else if (o instanceof Boolean || o.getClass().equals(Boolean.TYPE)) {
			cell.setCellValue((Boolean) o);
		} else {
			cell.setCellValue(o.toString());	
		}
	}
}
