package uk.co.terminological.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import uk.co.terminological.datatypes.CrossMap;
import uk.co.terminological.datatypes.EavMap;


public class ExcelIO {

	public static Workbook readExcel(String filename) throws IOException {
		return readExcel(new File(filename));
	}


	public static Workbook readExcel(File file) throws IOException {
		try {
			return new HSSFWorkbook(new FileInputStream(file));
		} catch (OfficeXmlFileException e) {
			try {
				return new XSSFWorkbook(new FileInputStream(file));
			} catch (Exception e2) {
				throw new IOException(e);
			}
		} 
	}

	public static Sheet readExcel(File file, String sheetname) throws IOException {
		Workbook wb = readExcel(file);
		return wb.getSheet(sheetname);
	}

	public static void writeExcel(Workbook workbook, File file) throws IOException {
		workbook.write(new FileOutputStream(file));
	}

	public static void writeExcel(Sheet sheet, File file) throws IOException {
		sheet.getWorkbook().write(new FileOutputStream(file));
	}

	public static Sheet newExcel(String sheetname) {
		Workbook book = new XSSFWorkbook();
		Sheet sheet = book.createSheet(sheetname);
		return sheet;
	}

	public static CrossMap<String, String> loadXmap(File file, String sheetname, ColHeader hr) throws IOException, ExcelIOException {
		return loadXmap(file, sheetname, hr, 0, 1);
	}

	public static ArrayList<String> getSheetnames(File f) throws IOException {
		ArrayList<String> sheetnames = new ArrayList<String>();  
		Workbook book = ExcelIO.readExcel(f);
		for( int i = 0; i<book.getNumberOfSheets(); i++) {
			sheetnames.add(book.getSheetName(i));
		}
		return sheetnames;
	}

	public static EavMap<String,String,String> loadEavMap(File file, String sheetname) throws IOException {
		return loadEavMap(file, sheetname, 0, 0);
	}

	public static EavMap<String,String,String> loadEavMap(Sheet sheet) throws IOException {
		return loadEavMap(sheet, 0, 0);
	}

	public static EavMap<String,String,String> loadEavMap(File file, String sheetname, String originName) throws IOException {
		Sheet sheet = readExcel(file, sheetname);
		if (null == sheet) throw new ExcelIOException("no sheet found: "+sheetname);
		int numberedEntCol=0;
		int numberedAttrRow=0;
		for (Row row: sheet) {
			for (Cell cell: row) {
				if (ExcelIO.cellToString(cell).equalsIgnoreCase(originName)) {
					numberedEntCol=cell.getColumnIndex();
					numberedAttrRow=cell.getRowIndex();
					return loadEavMap(file, sheetname, numberedEntCol, numberedAttrRow);
				}
			}
		}
		throw new ExcelIOException("origin not found");
	}

	public static EavMap<String,String,String> loadEavMap(File file, String sheetname, int numberedEntCol, int numberedAttrRow) throws IOException {
		Sheet sheet = readExcel(file, sheetname);
		if (null == sheet) throw new ExcelIOException("no sheet found: "+sheetname);
		return loadEavMap(sheet, numberedEntCol, numberedAttrRow);
	}

	public static EavMap<String,String,String> loadEavMap(Sheet sheet, int numberedEntCol, int numberedAttrRow) {
		EavMap<String, String, String> out = new EavMap<String, String, String>();
		HashMap<Integer, String> colheads = new HashMap<Integer, String>();	
		for (Row row: sheet) {
			if (row.getRowNum() < numberedAttrRow) {
				//ignore
			} else if (row.getRowNum()==numberedAttrRow) {
				for (Cell cell: row) {
					colheads.put(cell.getColumnIndex(), ExcelIO.cellToString(cell));
				}
			} else {
				String rowhead = null;
				for (Cell cell: row) {
					if (cell.getColumnIndex() < numberedEntCol) {
						//ignore
					} else if (cell.getColumnIndex()==numberedEntCol) {
						rowhead = ExcelIO.cellToString(cell);
					} else {
						if (rowhead != null) {
							String colhead = colheads.get(cell.getColumnIndex());
							if (colhead != null) {
								String value = ExcelIO.cellToString(cell);
								out.put(rowhead, colhead, value);
							}
						}
					}
				}
			}
		}
		return out;

	}
	
	public static EavMap<Integer,String,String> loadColumnarEavMap(Sheet sheet, int numberedEntCol, int numberedAttrRow) {
		EavMap<Integer, String, String> out = new EavMap<Integer, String, String>();
		HashMap<Integer, String> colheads = new HashMap<Integer, String>();	
		for (Row row: sheet) {
			if (row.getRowNum() < numberedAttrRow) {
				//ignore
			} else if (row.getRowNum()==numberedAttrRow) {
				for (Cell cell: row) {
					colheads.put(cell.getColumnIndex(), ExcelIO.cellToString(cell));
				}
			} else {
				Integer rowhead = null;
				for (Cell cell: row) {
					if (cell.getColumnIndex() < numberedEntCol) {
						//ignore
					} 
					else {
						rowhead = row.getRowNum();
							String colhead = colheads.get(cell.getColumnIndex());
							if (colhead != null) {
								String value = ExcelIO.cellToString(cell);
								out.put(rowhead, colhead, value);
							}
						}
					}
				}
			
		}
		return out;

	}

	public static CrossMap<String, String> loadXmap(File file, String sheetname, String srcRow, String trgRow) throws IOException, ExcelIOException {
		Sheet sheet = readExcel(file, sheetname);
		CrossMap<String, String> out = new CrossMap<String, String>();
		if (null == sheet) throw new ExcelIOException("no sheet found: "+sheetname);
		Integer srcIndex = null;
		Integer trgIndex = null;
		for (Row row: sheet) {
			if (row.getRowNum()==0) {
				for (Cell cell: row) {
					if (cellToString(cell).equals(srcRow)) srcIndex = cell.getColumnIndex();
					if (cellToString(cell).equals(trgRow)) trgIndex = cell.getColumnIndex();
				}
				out.setHeader(srcRow,trgRow);
			} else {
				if (srcIndex == null) throw new ExcelIOException("no header found: "+srcRow);
				if (trgIndex == null) throw new ExcelIOException("no header found: "+trgRow);
				Cell keyCell = row.getCell(srcIndex, Row.RETURN_NULL_AND_BLANK);
				Cell valueCell = row.getCell(trgIndex, Row.RETURN_NULL_AND_BLANK);
				out.add(cellToString(keyCell), cellToString(valueCell));
			}
		}
		return out;
	}

	public static CrossMap<String, String> loadXmap(File file, String sheetname, ColHeader hr, int srcRow, int trgRow) throws IOException, ExcelIOException {
		Sheet sheet = readExcel(file, sheetname);
		CrossMap<String, String> out = new CrossMap<String, String>();
		if (null == sheet) throw new ExcelIOException("no sheet found: "+sheetname);
		for (Row row: sheet) {
			if ((hr.equals(ColHeader.YES) && row.getRowNum()==0)) {
				out.setHeader(
						cellToString(row.getCell(srcRow, Row.RETURN_NULL_AND_BLANK)),
						cellToString(row.getCell(trgRow, Row.RETURN_NULL_AND_BLANK)));
			} else {
				Cell keyCell = row.getCell(srcRow, Row.RETURN_NULL_AND_BLANK);
				Cell valueCell = row.getCell(trgRow, Row.RETURN_NULL_AND_BLANK);
				out.add(cellToString(keyCell), cellToString(valueCell));
			}
		}
		return out;
	}

	public enum ColHeader {YES, NO;}
	public enum RowHeader {YES, NO;}

	/**
	 * @param cell
	 * @return the string value of a HSSP cell, as displayed by Excel, or null if blank
	 */
	public static String cellToString(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case Cell.CELL_TYPE_NUMERIC:
			//if (DateUtil.isCellDateFormatted(cell)) {
			//	return DateFormatString.ISO_DATE_TIME_ZONE_FORMAT.format(cell.getDateCellValue());
			//} else {
				return Double.toString(cell.getNumericCellValue());
			//}
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

	public static <E extends Object, A extends Object, V extends Object>
	void writeEavMap(String sheetname, EavMap<E,A,V> map, List<A> columns, File f) throws IOException {
		Sheet out = newExcel(sheetname);
		mapToSheet(out,map,columns);
		writeExcel(out,f);
	}

	public static <E extends Object, A extends Object, V extends Object>
	void writeEavMap(String sheetname, EavMap<E,A,V> map, File f) throws IOException {
		Sheet out = newExcel(sheetname);
		mapToSheet(out,map);
		writeExcel(out,f);
	}

	public static <E extends Object, A extends Object, V extends Object>
	Workbook writeEavMap(String sheetname, EavMap<E,A,V> map) throws IOException {
		Sheet out = newExcel(sheetname);
		mapToSheet(out,map);
		return out.getWorkbook();
	}

	public static <E extends Object, A extends Object, V extends Object>
	Workbook writeEavMap(String sheetname, EavMap<E,A,V> map, List<A> columns) throws IOException {
		Sheet out = newExcel(sheetname);
		mapToSheet(out,map,columns);
		return out.getWorkbook();
	}


	private static <E extends Object, A extends Object, V extends Object>
	void mapToSheet(Sheet out, EavMap<E,A,V> map) {
		mapToSheet(out, map, new ArrayList<A>(map.getAttributeSet()));
	}

	private static <E extends Object, A extends Object, V extends Object>
	void mapToSheet(Sheet out, EavMap<E,A,V> map, List<A> sortOrder) {
		Row header = out.createRow(0);
		ArrayList<A> headings = new ArrayList<A>();
		Cell zero = header.createCell(0);
		zero.setCellValue("entity name");
		int i=1;
		for (A o: sortOrder) {
			headings.add(o);
			Cell cell = header.createCell(i);
			i++;
			writeToCell(cell,o);
		}
		int j=1;
		for (E entity: map.getEntitySet()) {
			Row entityRow = out.createRow(j);
			j++;
			Cell entId = entityRow.createCell(0);
			writeToCell(entId,entity);

			for (A att: map.getAttributeSet(entity)) {
				int index = headings.indexOf(att)+1;
				Cell cell = entityRow.createCell(index);
				if (map.containsKey(entity, att)) 
					writeToCell(cell,map.get(entity, att));
			}
		}
		for (int k=0;k<headings.size();k++) {
			out.autoSizeColumn(k);
		}
	}


	public static <E extends Object, A extends Object, V extends Object>
	Workbook appendEavMap(String sheetname, EavMap<E,A,V> map, Workbook wb) throws IOException {
		Sheet sht = wb.getSheet(sheetname);
		if (sht == null) sht = wb.createSheet(sheetname);
		mapToSheet(sht, map);
		return wb;
	}

	public static HashMap<String, EavMap<String,String,String>> loadEavMap(File file) throws IOException {
		HashMap<String, EavMap<String,String,String>> out = new HashMap<String, EavMap<String,String,String>>();
		for (String sheetname: ExcelIO.getSheetnames(file)) {
			out.put(sheetname, ExcelIO.loadEavMap(file, sheetname));
		}
		return out;
	}
	
	public static HashMap<String, EavMap<Integer,String,String>> loadColumnarEavMap(File file) throws IOException {
		HashMap<String, EavMap<Integer,String,String>> out = new HashMap<String, EavMap<Integer,String,String>>();
		for (String sheetname: ExcelIO.getSheetnames(file)) {
			out.put(sheetname, ExcelIO.loadColumnarEavMap(file, sheetname));
		}
		return out;
	}

	private static EavMap<Integer, String, String> loadColumnarEavMap(File file, String sheetname) throws IOException {
		Sheet sheet = readExcel(file, sheetname);
		return loadColumnarEavMap(sheet, 0, 0);
	}

	private static void writeToCell(Cell cell, Object o) {
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
