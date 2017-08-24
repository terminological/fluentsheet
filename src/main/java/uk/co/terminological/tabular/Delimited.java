package uk.co.terminological.tabular;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Stream;

import uk.co.terminological.datatypes.DelimitedParser;
import uk.co.terminological.datatypes.EavMap;
import uk.co.terminological.datatypes.Triple;
import uk.co.terminological.datatypes.DelimitedParser.EOFException;
import uk.co.terminological.datatypes.DelimitedParser.MalformedCSVException;
import uk.co.terminological.datatypes.Tuple;

/**
 * 
 * @author terminological
 *
 */
public class Delimited {

	private Reader reader;
	private Content metadata;

	/**
	 * Open a delimited data file and provide access to a set of configuration options
	 * 
	 * @param file
	 * @return
	 * @throws FileNotFoundException
	 */
	public static Format fromFile(File file) throws FileNotFoundException {
		return fromReader(new FileReader(file));
	}

	/**
	 * Open a delimited data stream and provide access to a set of configuration options
	 * 
	 * @param file
	 * @return
	 * @throws FileNotFoundException
	 */
	public static Format fromStream(InputStream is) {
		return fromReader(new InputStreamReader(is));
	}

	/**
	 * Open a delimited data reader (e.g. System.in and provide access to a set of configuration options
	 * 
	 * @param file
	 * @return
	 * @throws FileNotFoundException
	 */
	public static Format fromReader(Reader reader) {
		return new Format(reader);
	}
	
	/**
	 * converts a labelled csv file to a stream of identity, attribute, value triples. 
	 * This loads all the data into memory first. 
	 * @return
	 * @throws MalformedCSVException
	 */
	public Stream<Triple<String,String,String>> streamContents() throws MalformedCSVException {
		return getContents().stream();
	}
	
	public void close() throws IOException {
		reader.close();
	}
	
	/**
	 * converts a labelled csv to a EAV map based on header labels defined in the csv or provided during configuration  
	 * @return
	 */
	public EavMap<String,String,String> getContents() throws MalformedCSVException {
		EavMap<String,String,String> out = new EavMap<>();
		boolean eof = false;
		while (!eof) {
			try {
				Tuple<String,Map<String,String>> tmp = metadata.convertLine();
				out.add(tmp.getKey(), tmp.getValue());
			} catch (EOFException e) {
				eof = true;
			}
		}
		return out;
	}
	
	/**
	 * returns a row based EAV map of the contents of the sheet (row E, column A, string of contents V)
	 * Numbering is zero based.
	 * @throws MalformedCSVException 
	 */
	public EavMap<Long,Integer,String> getContentsByRow() throws MalformedCSVException {
		EavMap<Long,Integer,String> out = new EavMap<>();
		boolean eof = false;
		while (!eof) {
			try {
				Tuple<Long,List<String>> tmp = metadata.readLine();
				int i=0;
				for (String value: tmp.getValue()) {
					out.add(tmp.getKey(),  i, value);
				}
			} catch (EOFException e) {
				eof = true;
			}
		}
		return out;
	}
	
	/**
	 * A stream of individual indexed entries in the csv based on a row then column coordinate system.
	 * The current implementation loads the csv into memory before streaming.
	 * @return
	 * @throws MalformedCSVException
	 */
	public Stream<Triple<Long,Integer,String>> streamContentsByRow() throws MalformedCSVException {
		return getContentsByRow().stream();
	}

	/**
	 * Defines the detailed format of the delimited file we are parsing.
	 * Defaults to a csv file with unix line endings, for which every data item is enclosed with 
	 * double quotes.
	 * @author terminological
	 *
	 */
	public static class Format {

		Delimited out;
		String sep = ",";
		String term = "\r\n";
		String enc = "\"";
		String escape = "\""; 
		boolean enclosedMandatory = false;
		

		public Format(Reader reader) {
			out = new Delimited();
			out.reader = reader;
			out.metadata = new Content(out);
		}

		public Format windows() {
			term = "\r\n";
			return this;
		}
		
		public Format unix() {
			term = "\n";
			return this;
		}

		/**
		 * Construct the parser based on the inputs given and returns a default configuration
		 * option for the content itself. 
		 * @return
		 */
		public Content parse() {
			out.metadata.parser = new DelimitedParser(out.reader,sep,term,enc,escape,enclosedMandatory);
			return out.metadata;
		}

		/**
		 * Convenience method for a csv file  
		 * @return
		 */
		public Content csv() {
			sep = ",";
			return parse();
		}

		/**
		 * Convenience method for a tab separated file  
		 * @return
		 */
		public Content tsv() {
			sep = "\t";
			return parse();
		}

		/**
		 * Convenience method for a pipe separated file  
		 * @return
		 */
		public Content pipe() {
			sep = "|";
			return parse();
		}

		/**
		 * Convenience method for a space separated file  
		 * @return
		 */
		public Content space() {
			sep = " ";
			return parse();
		}

		public Format separator(String sep) {
			this.sep = sep;
			return this;
		}

		public Format terminator(String term) {
			this.term = term;
			return this;
		}
		
		

		/**
		 * The data may be enclosed with the given string parameter. The escape
		 * character is the same (e.g. "john said ""blah blah"""). This is typical of excel output
		 * @param enc
		 * @return
		 */
		public Format optionalEnclosed(String enc) {
			this.enc = enc;
			this.escape = enc;
			this.enclosedMandatory = false;
			return this;
		}

		/**
		 * The data may be enclosed with the first string parameter and within a 
		 * quoted string the escape character is the second parameter
		 * (e.g. 'john said \'blah blah\'')
		 * @param enc - the quote characters
		 * @param esc - the escape characters
		 * @return
		 */
		public Format optionalEnclosed(String enc, String esc) {
			this.enc = enc;
			this.escape = esc;
			this.enclosedMandatory = false;
			return this;
		}

		/**
		 * All data is enclosed with the given string parameter. The escape
		 * character is the same (e.g. "john said ""blah blah""").
		 * @param enc
		 * @return
		 */
		public Format mandatoryEnclosed(String enc) {
			this.enc = enc;
			this.escape = enc;
			this.enclosedMandatory = false;
			return this;
		}

		/**
		 * All data is enclosed with the given string parameter and within a 
		 * quoted string the escape character is the second parameter
		 * (e.g. 'john said \'blah blah\'')
		 * @param enc - the quote characters
		 * @param esc - the escape characters
		 * @return
		 */
		public Format mandatoryEnclosed(String enc, String esc) {
			this.enc = enc;
			this.escape = esc;
			this.enclosedMandatory = false;
			return this;
		}
	}

	/**
	 * Thrown if a label cannot be extracted from the csv file itself.
	 * @author terminological
	 *
	 */
	public static class LabelNotAvailableException extends RuntimeException {
		protected LabelNotAvailableException(Exception e) {super(e);}
	}

	/**
	 * The content configuration allows us to specify the different types of csv file.
	 * By default the parser will behave as if the columns in the content is all labelled with a header row
	 * and that the first column contains some form of row identifier.
	 * 
	 * @author terminological
	 *
	 */
	public static class Content {

		private DelimitedParser parser;
		private boolean labelled = true;
		private Delimited csv;
		private Optional<List<String>> labelMap = Optional.empty();
		private Optional<Integer> idLabel = Optional.of(0); //entity id is in first column if empty it is record number
		private Long recordNumber = 0L;
		Optional<String> nullValue = Optional.of("");

		protected Content(Delimited csv) {
			this.csv = csv;
		}

		public Content notNullable() {
			this.nullValue = Optional.empty();
			return this;
		}
		
		public Content nullable(String nullValue) {
			this.nullValue = Optional.ofNullable(nullValue);
			return this;
		}
		
		/**
		 * configure the column with this label as the identifier for each row. This enforces the labelled
		 * property on the configuration, and if nothing else is specified this will default to using the first 
		 * row of the input as a set of labels. 
		 * @param n
		 * @return
		 */
		public Content identifiers(String label) throws LabelNotAvailableException {
			if (!labelMap.isPresent()) headerLabels();
			idLabel = Optional.ofNullable(labelMap.get().indexOf(label) != -1 ? labelMap.get().indexOf(label) : null);
			return this;
		}

		/**
		 * configure the nth column (zero based) as the identifier for each row 
		 * @param n
		 * @return
		 */
		public Content identifiers(int n) {
			idLabel = Optional.of(n);
			return this;
		}

		/**
		 * configure the parser to use the contents of the first column 
		 * as the primary identifier for each row.
		 * @return
		 */
		
		public Content identifiers() {
			return identifiers(0);
		}

		/**
		 * configure the parser to use the row number as the primary 
		 * identifier for each row.
		 * @return
		 */
		public Content noIdentifiers() {
			idLabel = Optional.empty();
			return this;
		}

		/**
		 * Instruct parser to use the given labels for the columns of the csv file.
		 * If the number labels do not match the size of any given row then an exception will
		 * be raised.
		 *  
		 * @return
		 */
		public Content withLabels(String... strings)  {
			labelled = true;
			labelMap = Optional.of(Arrays.asList(strings));
			return this;
		}

		/**
		 * Instruct parser to use labels from the first line of the CSV file.
		 *  
		 * @return
		 * @throws LabelNotAvailableException -if first line does not exist or is not readable
		 */
		public Content headerLabels() throws LabelNotAvailableException {
			labelled = true;
			if (!labelMap.isPresent()) {
				try {
					labelMap = Optional.of(parser.readLine());
				} catch (Exception e) {
					throw new LabelNotAvailableException(e);
				}
			}
			return this;
		}

		/**
		 * Instruct parser to not try and label data which will be available through on a by row 
		 * basis. If labelled data is requested later the
		 * parser will return a map of the string values of the index of the different column
		 * as the label.
		 *   
		 * @param identifierColumn
		 * @return
		 */
		public Content noLabels(int identifierColumn) {
			labelled = false;
			labelMap = Optional.empty();
			idLabel = Optional.of(identifierColumn);
			return this;
		}

		/**
		 * return a fully configured parser which can be used to read content.
		 * @return
		 * @throws LabelNotAvailableException
		 */
		public Delimited begin() throws LabelNotAvailableException {
			if (labelled && !labelMap.isPresent()) headerLabels();
			return csv;
		}

		//get record as a record number, and list of string values
		//missing values are empty strings at this point.
		protected Tuple<Long,List<String>> readLine() throws EOFException, MalformedCSVException {
			Tuple<Long,List<String>> out = Tuple.create(
				this.recordNumber,
				this.parser.readLine());
			this.recordNumber += 1;
			return out;
		}

		//convert the list of values to a map
		//string is either labels (if they are defined) or list index number as String
		//empty strings are treated as missing values and omitted for consistency with the 
		//Excel parser.
		private Map<String,String> label(List<String> values) {
			Map<String,String> out = new HashMap<>();
			Iterator<String> i1 = 
					labelMap.map(l -> l.iterator())
					.orElse(
						// an iterator of "0","1","2",..., "n-1", "n" for labels if none defined 
						new Iterator<String>() {
							int count = -1;
							@Override
							public boolean hasNext() {
								return count+1 < values.size();
							}
							@Override
							public String next() {
								count++;
								return ""+count;
							}
						}
					);
			Iterator<String> i2 = values.iterator();
			while (i1.hasNext() || i2.hasNext()) {
				String value = i2.next();
				String key = i1.next();
				if (nullValue.isPresent() && value.equals(nullValue.get())) {
					// out.put(i1.next(), null);
				} else {
					out.put(key, value);
				}
			}
			return out;
		}

		//if the identifier is a column in the data then 
		//lookup the identifier in the data otherwise
		//use the record number from readLine()
		//convert the list of values to a map
		protected Tuple<String, Map<String,String>> convertLine() throws EOFException, MalformedCSVException {
			Tuple<Long,List<String>> raw = readLine();
			return Tuple.create(
					this.idLabel.map(l -> raw.getValue().get(l)).orElse(raw.getKey().toString()), 
					label(raw.getValue()));
		}
		
		
	}

}
