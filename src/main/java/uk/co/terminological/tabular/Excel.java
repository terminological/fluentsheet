package uk.co.terminological.tabular;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import uk.co.terminological.datatypes.EavMap;

public class Excel {

	private Workbook excel;
	
	private Excel() {}
	
	public static Excel fromFile(File file) throws IOException {
		return Excel.fromStream(new FileInputStream(file));
	}
	
	public static Excel fromStream(InputStream stream) throws IOException {
		Excel out = new Excel();
		try {
			out.excel = WorkbookFactory.create(stream);
		} catch (InvalidFormatException e) {
			throw new IOException(e);
		} 
		return out;
	}

	public static Excel create() {
		Excel out = new Excel();
		out.excel = new XSSFWorkbook();
		return out;
	}
	
	public boolean hasSheet(String name) {
		return this.getSheetnames().contains(name);
	}
	
	public ExcelSheet addSheet(String name) {
		Sheet sheet = excel.createSheet(name);
		return new ExcelSheet(sheet);
	}
	
	public ExcelSheet getSheet(String name) {
		Sheet sheet = excel.getSheet(name);
		return new ExcelSheet(sheet);
	}
	
	public ArrayList<String> getSheetnames() {
		ArrayList<String> sheetnames = new ArrayList<String>();  
		for( int i = 0; i<excel.getNumberOfSheets(); i++) {
			sheetnames.add(excel.getSheetName(i));
		}
		return sheetnames;
	}
	
	public ArrayList<ExcelSheet> getSheets() {
		ArrayList<ExcelSheet> sheetnames = new ArrayList<ExcelSheet>();  
		for( int i = 0; i<excel.getNumberOfSheets(); i++) {
			sheetnames.add(new ExcelSheet(excel.getSheetAt(i)));
		}
		return sheetnames;
	}
	
	public Stream<ExcelSheet> stream() {
		return this.getSheets().stream();
	}
	
	public ExcelSheet getFirst() {
		return new ExcelSheet(excel.getSheetAt(0));
	}
	
	public <E extends Object, A extends Object, V extends Object>
	Excel addSheet(String name, EavMap<E,A,V> map, List<A> sortOrder) {
		Sheet out = excel.createSheet(name);
		Row header = out.createRow(0);
		ArrayList<A> headings = new ArrayList<A>();
		Cell zero = header.createCell(0);
		zero.setCellValue("entity name");
		int i=1;
		for (A o: sortOrder) {
			headings.add(o);
			ExcelCell cell = new ExcelCell(header.createCell(i));
			i++;
			cell.setValue(o);
		}
		int j=1;
		for (E entity: map.getEntitySet()) {
			Row entityRow = out.createRow(j);
			j++;
			ExcelCell cell = new ExcelCell(entityRow.createCell(0));
			cell.setValue(entity);

			for (A att: map.getAttributeSet(entity)) {
				int index = headings.indexOf(att)+1;
				ExcelCell cell2 = new ExcelCell(entityRow.createCell(index));
				if (map.containsKey(entity, att)) 
					cell2.setValue(map.get(entity, att));
			}
		}
		for (int k=0;k<headings.size();k++) {
			out.autoSizeColumn(k);
		}
		return this;
	}
	
	public <E extends Object, A extends Object, V extends Object>
	Excel addSheet(String name, EavMap<E,A,V> map) {
		return addSheet(name, map, (List<A>) new ArrayList<A>(map.getAttributeSet()));
	}
	
	public void write(File file) throws IOException {
		excel.write(new FileOutputStream(file));
	}
}
