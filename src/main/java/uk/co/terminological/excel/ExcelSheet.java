package uk.co.terminological.excel;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;

import uk.co.terminological.datatypes.EavMap;
import uk.co.terminological.datatypes.NoMatchException;

public class ExcelSheet {

	protected Excel xl;
	protected Sheet sheet;
	protected Integer instanceIdSource = null;
	protected boolean transpose = false;
	protected EavMap<Integer,Integer,ExcelCell> values;


	static Pattern pat = Pattern.compile("^([A-Z][A-Z]*)([1-9][0-9]*)$");

	protected ExcelSheet(Sheet sheet, Excel xl) {
		this.values = new EavMap<Integer,Integer,ExcelCell>();
		this.sheet = sheet;
		this.xl = xl;
		for (Row row: sheet) {
			for (Cell cell: row) {
				int colNum = cell.getColumnIndex();
				int rowNum = cell.getRowIndex();
				values.add(rowNum, colNum, new ExcelCell(cell));
			}
		}
	}

	protected ExcelSheet(ExcelSheet xlsheet) {
		this.sheet = xlsheet.sheet;
		this.values = xlsheet.values;
		this.xl = xlsheet.xl;
		this.transpose = xlsheet.transpose;
		this.instanceIdSource = xlsheet.instanceIdSource;
	}

	

	public ExcelCell getCellAtIndex(int columnOne, int rowOne) {
		return values.get(rowOne, columnOne);
	}

	public ExcelCell getCellAtReference(String A1) throws NoMatchException {
		A1 = A1.replace("$","").toUpperCase();
		Matcher m = pat.matcher(A1);
		if (!m.matches()) throw new NoMatchException(A1 + " is not a valid cell reference");
		String colRef = m.group(1);
		int row = Integer.parseInt(m.group(2));
		int column = colRef.charAt( 0 )-'A'+1;
		if (colRef.length() == 2) {
			column = column*26 + colRef.charAt(1)-'A'+1;
		}
		return getCellAtIndex(column,row);
	}

	public Horizontal horizontal() {
		return new Horizontal(this);
	}

	public Vertical vertical() {
		return new Vertical(this);
	}
	
	

	public ExcelSheet identifier(int rowOrColumnIndex) {
		this.instanceIdSource = rowOrColumnIndex;
		return this;
	}

	public static interface Labelable {
		default public Labelled labelled() {
			return new Labelled(this);
		}
	}
	
	public static class Horizontal extends ExcelSheet {
		protected Horizontal(ExcelSheet xlsheet) {
			super(xlsheet);
			values = values.transpose();
		}
	}

	public static class Vertical extends ExcelSheet {
		protected Vertical(ExcelSheet xlsheet) {
			super(xlsheet);
		}
	}
	
	public static class Labelled extends ExcelSheet {

		private HashMap<String,Integer> labels = new HashMap<>();

		protected Labelled(ExcelSheet xlsheet) {
			super(xlsheet);
			values.get(1)
				.entrySet()
				.stream()
				.forEach(kv -> {
					labels.put(kv.getValue().toString(), kv.getKey());
				});
		}

		private Integer getIndexForLabel(String label) {
			return labels.get(label);
		}

		public ExcelSheet identifier(String label) {
			this.instanceIdSource = getIndexForLabel(label);
			return this;
		}

	}





}
