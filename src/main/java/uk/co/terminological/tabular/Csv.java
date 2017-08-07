package uk.co.terminological.tabular;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.util.Optional;

import uk.co.terminological.datatypes.DelimitedParser;
import uk.co.terminological.tabular.ExcelSheet.Metadata;
import uk.co.terminological.tabular.ExcelSheet.Orientation;

public class Csv {

	private Reader reader;
	private Metadata metadata = new Metadata();
	private DelimitedParser parser;
	
	public static Csv fromFile(File file) throws FileNotFoundException {
		return fromReader(new FileReader(file));
	}
	
	public static Csv fromStream(InputStream is) {
		return fromReader(new InputStreamReader(is));
	}
	
	public static Csv fromReader(Reader reader) {
		Csv out = new Csv();
		out.reader = reader;		
		return out;
	}
	
	public Metadata with() {
		return metadata;
	}
	
	public static class Metadata {
		
		boolean labelled = true;
		private Optional<Integer> idLabel = Optional.of(0); //entity id is in first column (vertical) / row (horizontal)
		
		public Metadata identifiers(String label) {
			idLabel = labelMap.entrySet().stream().filter(kv -> kv.getValue().equals(label)).findFirst().map(kv -> kv.getKey());
			return this;
		}

		public Metadata identifiers(int rowOrColumnOffset) {
			idLabel = Optional.of(rowOrColumnOffset);
			return this;
		}

		public Metadata identifiers() {
			return identifiers(0);
		}
		public Metadata labelled() {
			labelled = true;
			return this;
		}
		
		public Metadata unlabelled() {
			labelled = false;
			return this;
		}
		
		
	}
	
}
