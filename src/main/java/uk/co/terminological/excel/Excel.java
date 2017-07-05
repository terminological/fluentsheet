package uk.co.terminological.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.stream.Stream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	private Workbook excel;
	
	private Excel() {}
	
	public static Excel fromFile(File file) throws IOException {
		return Excel.fromStream(new FileInputStream(file));
	}
	
	public static Excel fromStream(InputStream stream) throws IOException {
		Excel out = new Excel();
		try {
			out.excel = new HSSFWorkbook(stream);
		} catch (OfficeXmlFileException e) {
			try {
				out.excel = new XSSFWorkbook(stream);
			} catch (Exception e2) {
				throw new IOException(e);
			}
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
		return new ExcelSheet(sheet, this);
	}
	
	public ExcelSheet getSheet(String name) {
		Sheet sheet = excel.getSheet(name);
		return new ExcelSheet(sheet, this);
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
			sheetnames.add(new ExcelSheet(excel.getSheetAt(i),this));
		}
		return sheetnames;
	}
	
	public Stream<ExcelSheet> stream() {
		return this.getSheets().stream();
	}
	
	public ExcelSheet getFirst() {
		return new ExcelSheet(excel.getSheetAt(0),this);
	}
	
	
}
