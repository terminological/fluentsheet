package uk.co.terminological.tabular;

import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import uk.co.terminological.datatypes.EavMap;
import uk.co.terminological.datatypes.NoMatchException;
import uk.co.terminological.datatypes.Triple;
import uk.co.terminological.datatypes.Tuple;

/**
 * Access to the sheet with default configuration for a vertically oriented table of values where the instance
 * is given by the row number.
 * @author rc538
 *
 */
public class ExcelSheet {

	private Sheet sheet;
	private EavMap<Long,Integer,String> rowCache = null;
	private EavMap<Long,Integer,String> colCache = null;
	private Content meta;

	private static Pattern pat = Pattern.compile("^([A-Z][A-Z]*)([1-9][0-9]*)$");

	/**
	 * creates and reads an excel sheet
	 * @param sheet
	 * @param xl
	 */
	protected ExcelSheet(Sheet sheet) {
		this.sheet = sheet;
		this.meta = new Content(this);
	}

	/**
	 * Overriding of default configuration
	 * @return
	 */
	public Content with() {
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
	public EavMap<Long,Integer,String> getContentsByRow() {
		if (rowCache == null) {
			fillCaches();
		}
		return rowCache;
	}

	public Stream<Triple<Long,Integer,String>> streamContentsByRow() {
		return getContentsByRow().stream();
	}


	/**
	 * returns a column based EAV map of the contents of the sheet (column E, row A, string of contents V)
	 * Calculation errors, nulls, blank cells are not returned.
	 * Numbering is zero based.
	 * This method caches the result.
	 */
	public EavMap<Long,Integer,String> getContentsByColumn() {
		if (colCache == null) {
			fillCaches();
		}
		return colCache;
	}
	
	/**
	 * returns a column based stream of column index, row index, value for non empty cells
	 * in the spreadsheet
	 * @return
	 */
	public Stream<Triple<Long,Integer,String>> streamContentsByColumn() {
		return getContentsByColumn().stream();
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
					colCache.add(Integer.toUnsignedLong(colNum), rowNum, tmp);
					rowCache.add(Integer.toUnsignedLong(rowNum), colNum, tmp);
				}
			}
		}
	}

	/**
	 * A stream of labelled values
	 * @return
	 */
	public Stream<Triple<String,String,String>> streamContents() {
		return getContents().stream();
	}
	
	/**
	 * converts a sheet to a EAV map based on header labels.
	 * Only non empty / non blank cells will be present here. 
	 * i.e. missing values in the sheet will not be in this map.   
	 * @return
	 */
	public EavMap<String,String,String> getContents() {

		EavMap<String,String,String> out = new EavMap<>();
		EavMap<Long,Integer,String> tmp = meta.getRawContents();
		Optional<Map<Integer,String>> labels = meta.getLabels();

		//iterate by raw entity (e.g. by row if vertical)
		for (Long rawEnt: tmp.getEntitySet() ) {

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

	public static class Content {

		protected Content(ExcelSheet sheet) {
			this.sheet = sheet;
		}

		private ExcelSheet sheet;
		private Labelling attributes = Labelling.LABELLED;
		private Orientation orientation = Orientation.VERTICAL;
		private int xOrigin = 0;
		private int yOrigin = 0;
		private Optional<Integer> idLabel = Optional.of(0); //entity id is in first column (vertical) / row (horizontal)
		private Map<Integer,String> labelMap; //maybe discontinuous and not zero based

		protected Optional<Map<Integer,String>> getLabels() {
			if (!attributes.equals(Labelling.LABELLED)) return Optional.empty();
			if (labelMap == null) {
				if (orientation.equals(Orientation.HORIZONTAL)) {
					EavMap<Long,Integer,String> raw = sheet.getContentsByColumn();
					labelMap = raw.get(Integer.toUnsignedLong(xOrigin));
				} else {
					EavMap<Long,Integer,String> raw = sheet.getContentsByRow();
					labelMap = raw.get(Integer.toUnsignedLong(yOrigin));
				}
			}
			return Optional.of(labelMap);
		}

		protected EavMap<Long,Integer,String> getRawContents() {
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

		public Content origin(String A1) throws NoMatchException {
			cellReftoCoordinates(A1).consume(xy -> {
				xOrigin = xy.getFirst();
				yOrigin = xy.getSecond();
			});
			return this;
		}

		public Content withLabels(String... strings) {
			int i = getRawAttributeStart();
			labelMap = new HashMap<>();
			for (String string: strings) {
				labelMap.put(i, string);
				i++;
			}
			return this;
		}

		public Content identifiers(String label) {
			idLabel = labelMap.entrySet().stream().filter(kv -> kv.getValue().equals(label)).findFirst().map(kv -> kv.getKey());
			return this;
		}

		public Content identifiers(int rowOrColumnOffset) {
			idLabel = Optional.of(rowOrColumnOffset+
					(orientation.equals(Orientation.HORIZONTAL) ? yOrigin : xOrigin)
					);
			return this;
		}

		public Content identifiers() {
			return identifiers(0);
		}

		public Content noIdentifiers() {
			idLabel = Optional.empty();
			return this;
		}

		public Content headerLabels() {
			attributes = Labelling.LABELLED;
			return this;
		}

		public Content noHeaderLabels() {
			attributes = Labelling.LABELLED;
			return this;
		}

		public Content horizontal() {
			orientation = Orientation.HORIZONTAL;
			return this;
		}

		public Content vertical() {
			orientation = Orientation.VERTICAL;
			return this;
		}

		public ExcelSheet begin() {
			return sheet;
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
