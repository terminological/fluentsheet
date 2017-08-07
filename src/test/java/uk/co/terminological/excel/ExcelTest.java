package uk.co.terminological.excel;

import java.io.File;
import java.io.IOException;

import uk.co.terminological.datatypes.NoMatchException;
import uk.co.terminological.tabular.Excel;
import uk.co.terminological.tabular.ExcelSheet;

public class ExcelTest {

	public static void main(String[] args) throws IOException, NoMatchException {
		
		String file = ExcelTest.class.getResource("/test.xlsx").getFile();
		
		Excel xl = Excel.fromFile(new File(file));
		
		System.out.println("============");
		
		ExcelSheet tmp = xl.getSheet("vertical");
		
		tmp.with()
			.vertical()
			.headerLabels()
			.origin("A1");
		
		tmp.getContents().stream().forEach(System.out::println);

		System.out.println("============");
		
		ExcelSheet tmp2 = xl.getSheet("horizontal");
		
		tmp2.with()
			.horizontal()
			.headerLabels()
			.origin("A1");
		
		tmp2.getContents().stream().forEach(System.out::println);
		
		System.out.println("============");
		
		ExcelSheet tmp3 = xl.getSheet("displaced");
		
		tmp3.with()
			.vertical()
			.headerLabels()
			.origin("C4");
		
		tmp3.getContents().stream().forEach(System.out::println);

		
		
	}

}
