package uk.co.terminological.tabular;

import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import uk.co.terminological.datatypes.EavMap;
import uk.co.terminological.datatypes.NoMatchException;
import uk.co.terminological.datatypes.Tuple;

/**
 * Access to the sheet with default configuration for a vertically oriented table of values where the instance
 * is given by the row number.
 * @author rc538
 *
 */
public class ExcelSheet {

	private Sheet sheet;
	private EavMap<Integer,Integer,String> rowCache = null;
	private EavMap<Integer,Integer,String> colCache = null;
	private Metadata meta;

	private static Pattern pat = Pattern.compile("^([A-Z][A-Z]*)([1-9][0-9]*)$");

	/**
	 * creates and reads an excel sheet
	 * @param sheet
	 * @param xl
	 */
	protected ExcelSheet(Sheet sheet) {
		this.sheet = sheet;
		this.meta = new Metadata(this);
	}

	/**
	 * Overriding of default configuration
	 * @return
	 */
	public Metadata with() {
		return meta;
	}

	public ExcelSheet recalc() {
		fillCaches();
		return this;
	}

	/**
	 * returns a row based EAV map of the contents of the sheet (row E, column A, string of contents V)
	 * Calculation errors, nulls, blank cells are not returned.
	 * Numbering is zero based.
	 * This method caches the result.
	 */
	public EavMap<Integer,Integer,String> getContentsByRow() {
		if (rowCache == null) {
			fillCaches();
		}
		return rowCache;
	}



	/**
	 * returns a column based EAV map of the contents of the sheet (column E, row A, string of contents V)
	 * Calculation errors, nulls, blank cells are not returned.
	 * Numbering is zero based.
	 * This method caches the result.
	 */
	public EavMap<Integer,Integer,String> getContentsByColumn() {
		if (colCache == null) {
			fillCaches();
		}
		return colCache;
	}

	private void fillCaches() {
		colCache = new EavMap<>();
		rowCache = new EavMap<>();
		for (Row row: sheet) {
			for (Cell cell: row) {
				int colNum = cell.getColumnIndex();
				int rowNum = cell.getRowIndex();
				String tmp = new ExcelCell(cell).toString();
				if (tmp != null) {
					colCache.add(colNum, rowNum, tmp);
					rowCache.add(rowNum, colNum, tmp);
				}
			}
		}
	}

	/**
	 * converts a sheet to a EAV map based on header labels  
	 * @return
	 */
	public EavMap<String,String,String> getContents() {

		EavMap<String,String,String> out = new EavMap<>();
		EavMap<Integer,Integer,String> tmp = meta.getRawContents();
		Optional<Map<Integer,String>> labels = meta.getLabels();

		//iterate by raw entity (e.g. by row if vertical)
		for (Integer rawEnt: tmp.getEntitySet() ) {

			//Integer rawEnt = eav.getKey();
			//don't map if the cell is outside of the start range
			if (!(rawEnt < meta.getRawEntityStart())) {

				Map<Integer,String> rawAttrValue = tmp.get(rawEnt);

				String entity;

				if (meta.getIdRawAttribute().isPresent()) {
					entity = rawAttrValue.get(meta.getIdRawAttribute().get());
				} else {
					entity = ""+(rawEnt - meta.getRawEntityStart());
				}

				//iterate by raw attribute (e.g. by column if vertical)
				for (Integer rawAttr: tmp.getAttributeSet(rawEnt)) {

					//Don't map if the raw attribute is the entity identifier 
					if (!meta.getIdRawAttribute().filter(i -> i.equals(rawAttr)).isPresent()) {
							
						//Don't map if the cell is outside the start range
						if (!(rawAttr < meta.getRawAttributeStart())) {

							String attribute;

							if (labels.isPresent()) {
								if (labels.get().containsKey(rawAttr)) {
									attribute = labels.get().get(rawAttr);
								} else {
									//there was no label defined for this item
									//we could fall back to the row index as the attribute.
									attribute = ""+(rawAttr - meta.getRawAttributeStart());
								}
							} else {
								attribute = ""+(rawAttr - meta.getRawAttributeStart());
							}

							out.add(entity, attribute, tmp.get(rawEnt, rawAttr));
						}
					}
				}
			}
		}
		return out;

	}

	/**
	 * Returns the cell at a given sheet index.
	 * The index is one based in both row and column to correlate with excel.
	 * @param columnOne
	 * @param rowOne
	 * @return
	 */
	public ExcelCell getCellAtIndex(int columnOne, int rowOne) {
		return new ExcelCell(sheet.getRow(rowOne).getCell(columnOne, Row.RETURN_NULL_AND_BLANK));
	}

	/**
	 * Returns the cell at a specific cell reference location
	 * @param A1
	 * @return
	 * @throws NoMatchException
	 */
	public ExcelCell getCellAtReference(String A1) throws NoMatchException {
		Tuple<Integer,Integer> xy = cellReftoCoordinates(A1);
		return getCellAtIndex(xy.getFirst(),xy.getSecond());
	}
	
	
	

	//Configuration metadata object

	public static class Metadata {

		protected Metadata(ExcelSheet sheet) {
			this.sheet = sheet;
		}

		private ExcelSheet sheet;
		private Labelling attributes = Labelling.LABELLED;
		private Orientation orientation = Orientation.VERTICAL;
		private int xOrigin = 0;
		private int yOrigin = 0;
		private Optional<Integer> idLabel = Optional.of(0); //entity id is in first column (vertical) / row (horizontal)
		private Map<Integer,String> labelMap;

		protected Optional<Map<Integer,String>> getLabels() {
			if (!attributes.equals(Labelling.LABELLED)) return Optional.empty();
			if (labelMap == null) {
				if (orientation.equals(Orientation.HORIZONTAL)) {
					EavMap<Integer,Integer,String> raw = sheet.getContentsByColumn();
					labelMap = raw.get(xOrigin);
				} else {
					EavMap<Integer,Integer,String> raw = sheet.getContentsByRow();
					labelMap = raw.get(yOrigin);
				}
			}
			return Optional.of(labelMap);
		}

		protected EavMap<Integer,Integer,String> getRawContents() {
			if (orientation.equals(Orientation.HORIZONTAL)) {
				return sheet.getContentsByColumn();
			} else {
				return sheet.getContentsByRow();
			}
		}

		protected int getRawEntityStart() {
			if (orientation.equals(Orientation.HORIZONTAL)) {
				return xOrigin + (attributes.equals(Labelling.LABELLED) ? 1: 0);
			} else {
				return yOrigin + (attributes.equals(Labelling.LABELLED) ? 1: 0);
			}
		}

		protected int getRawAttributeStart() {
			if (orientation.equals(Orientation.HORIZONTAL)) {
				return yOrigin;
			} else {
				return xOrigin;
			}
		}

		protected Optional<Integer> getIdRawAttribute() {
			return idLabel.map(i -> getRawAttributeStart()+i);
		}

		public Metadata origin(String A1) throws NoMatchException {
			cellReftoCoordinates(A1).consume(xy -> {
				xOrigin = xy.getFirst();
				yOrigin = xy.getSecond();
			});
			return this;
		}

		public Metadata withLabels(String... strings) {
			int i = getRawAttributeStart();
			labelMap = new HashMap<>();
			for (String string: strings) {
				labelMap.put(i, string);
				i++;
			}
			return this;
		}

		public Metadata identifiers(String label) {
			idLabel = labelMap.entrySet().stream().filter(kv -> kv.getValue().equals(label)).findFirst().map(kv -> kv.getKey());
			return this;
		}

		public Metadata identifiers(int rowOrColumnOffset) {
			idLabel = Optional.of(rowOrColumnOffset+
					(orientation.equals(Orientation.HORIZONTAL) ? yOrigin : xOrigin)
					);
			return this;
		}

		public Metadata identifiers() {
			return identifiers(0);
		}

		public Metadata noIdentifiers() {
			idLabel = Optional.empty();
			return this;
		}

		public Metadata headerLabels() {
			attributes = Labelling.LABELLED;
			return this;
		}

		public Metadata noHeaderLabels() {
			attributes = Labelling.LABELLED;
			return this;
		}

		public Metadata horizontal() {
			orientation = Orientation.HORIZONTAL;
			return this;
		}

		public Metadata vertical() {
			orientation = Orientation.VERTICAL;
			return this;
		}

	}

	public static enum Labelling {
		LABELLED,
		UNLABELLED
	}

	public static enum Orientation {
		HORIZONTAL,
		VERTICAL
	}

	//Internal utility methods

	private static Tuple<Integer,Integer> cellReftoCoordinates(String A1) throws NoMatchException {
		A1 = A1.replace("$","").toUpperCase();
		Matcher m = pat.matcher(A1);
		if (!m.matches()) throw new NoMatchException(A1 + " is not a valid cell reference");
		String colRef = m.group(1);
		int row = Integer.parseInt(m.group(2))-1;
		int column = colRef.charAt( 0 )-'A';
		if (colRef.length() == 2) {
			column = column*26 + colRef.charAt(1)-'A';
		}
		return Tuple.create(column, row);
	}
}
