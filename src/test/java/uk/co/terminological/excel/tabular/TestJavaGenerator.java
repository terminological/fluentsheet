package uk.co.terminological.excel.tabular;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Collections;
import java.util.Map;
import java.util.Optional;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;

import uk.co.terminological.datatypes.EavMap;
import uk.co.terminological.mappers.StringCaster;
import uk.co.terminological.parser.StateMachineException;
import uk.co.terminological.parser.StateMachineExecutor.ErrorHandler;
import uk.co.terminological.tabular.Delimited;
import uk.co.terminological.tabular.Delimited.LabelNotAvailableException;
import uk.co.terminological.tabular.JavaGenerator;

public class TestJavaGenerator {
	
	static File mysqlCsv = new File(DelimitedTest.class.getResource("/mysql.csv").getFile());

	//@Test
	public static void main(String[] args) throws LabelNotAvailableException, FileNotFoundException, IOException, StateMachineException {
			BasicConfigurator.configure();
			Logger.getRootLogger().setLevel(Level.ALL);
			Delimited d = Delimited.fromFile(mysqlCsv)
					.separatedByEnclosedWith(",", "\"")
					.parse(ErrorHandler.STRICT)
					.nullable("NULL")
					.noIdentifiers()
					.begin();
			
			EavMap<String,String,String> map = d.getContents();
			Optional<Map<String,String>> first = map.streamEntities().findFirst().map(kv -> kv.getValue());
			
			String ifaceDef = 
					JavaGenerator.createInterface(
							"Test", "org.example", Collections.emptyList(), 
							StringCaster.guessTypes(first.get())
						);
			System.out.println(ifaceDef);
		
		
	}

}
