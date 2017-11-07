/**
 * 
 */
package uk.co.terminological.excel.tabular;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.junit.Before;
import org.junit.Test;

import uk.co.terminological.parser.ParserException;

import uk.co.terminological.tabular.Delimited;
import uk.co.terminological.tabular.Delimited.LabelNotAvailableException;

/**
 * @author terminological
 *
 */
public class DelimitedTest {

	static File xlsxCsv = new File(DelimitedTest.class.getResource("/test.xlsx.csv").getFile());
	static File xlsxTsv = new File(DelimitedTest.class.getResource("/test.xlsx.tsv").getFile());
	static File odsCsv = new File(DelimitedTest.class.getResource("/test.ods.csv").getFile());
	static File mysqlCsv = new File(DelimitedTest.class.getResource("/mysql.csv").getFile());
	
	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
	}

	/**
	 * Test method for {@link uk.co.terminological.tabular.Delimited#fromFile(java.io.File)}.
	 * @throws IOException 
	 * @throws FileNotFoundException 
	 * @throws LabelNotAvailableException 
	 * @throws StateMachineException 
	 * @throws MalformedCSVException 
	 */
	@Test
	public final void testFromFile() throws LabelNotAvailableException, FileNotFoundException, IOException, ParserException {
		{
			Delimited d = Delimited.fromFile(mysqlCsv)
					.csv()
					.nullable("NULL")
					.noIdentifiers()
					.begin();
			d.streamContents().forEach(System.out::println);
			d.close();
		}
		{
			Delimited d = Delimited.fromFile(mysqlCsv)
					.csv()
					.nullable("NULL")
					.noIdentifiers()
					.begin();
			d.streamContentsByRow().forEach(System.out::println);
			d.close();
		}
	}

	/**
	 * Test method for {@link uk.co.terminological.tabular.Delimited#streamContents()}.
	 *
	@Test
	public final void testStreamContents() {
		fail("Not yet implemented");
	}

	/**
	 * Test method for {@link uk.co.terminological.tabular.Delimited#getContents()}.
	 *
	@Test
	public final void testGetContents() {
		fail("Not yet implemented");
	}

	/**
	 * Test method for {@link uk.co.terminological.tabular.Delimited#getContentsByRow()}.
	 *
	@Test
	public final void testGetContentsByRow() {
		fail("Not yet implemented");
	}

	/**
	 * Test method for {@link uk.co.terminological.tabular.Delimited#streamContentsByRow()}.
	 *
	@Test
	public final void testStreamContentsByRow() {
		fail("Not yet implemented");
	}
	*/

}
