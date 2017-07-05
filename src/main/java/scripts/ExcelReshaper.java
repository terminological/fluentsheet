package scripts;

import java.io.File;
import java.io.IOException;
import java.util.Map;

import uk.co.terminological.datatypes.EavMap;
import uk.co.terminological.excel.ExcelIO;

public class ExcelReshaper {

	public ExcelReshaper() {
		// TODO Auto-generated constructor stub
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		Map<String, EavMap<Integer,String,String>> map = ExcelIO.loadColumnarEavMap(new File("/home/terminological/Share/Windows10/annuity/Technical/MagnumMedications.xlsx"));
		
		map.entrySet().stream().forEach(kv -> {
			String sheet = kv.getKey();
			EavMap<Integer, String, String> eav = kv.getValue();
			eav.stream().forEach(trip -> {System.out.println(sheet+"\t"+trip.getSecond()+"\t"+trip.getThird());});
				
		});
		
		
		//
		
	}

}
