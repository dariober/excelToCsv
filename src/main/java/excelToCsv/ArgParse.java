package excelToCsv;

import net.sourceforge.argparse4j.ArgumentParsers;
import net.sourceforge.argparse4j.impl.Arguments;
import net.sourceforge.argparse4j.inf.ArgumentParser;
import net.sourceforge.argparse4j.inf.ArgumentParserException;
import net.sourceforge.argparse4j.inf.Namespace;

public class ArgParse {
	
	public static String PROG_NAME= "excelToCsv";
	public static String VERSION= "0.1.0";
	public static String WEB_ADDRESS= "https://github.com/dariober/...";
	
	/* Parse command line args */
	public static Namespace argParse(String[] args){
		ArgumentParser parser= ArgumentParsers
				.newFor(PROG_NAME)
				.build()
				.defaultHelp(true)
				.version("${prog} " + VERSION)
				.description("DESCRIPTION\n"
+ "Export Excel files to CSV");
		
		parser.addArgument("input")
			.type(String.class)
			.required(false)
			.nargs("+")
			.help("xlsx or xls files to convert");

		parser.addArgument("--delimiter", "-d")
			.type(String.class)
			.required(false)
			.setDefault("\\t")
			.help("Column delimiter");

		parser.addArgument("--na-string", "-na")
			.type(String.class)
			.required(false)
			.setDefault("")
			.help("String for missing values (empty cells)");

		parser.addArgument("--quote", "-q")
			.type(String.class)
			.required(false)
			.setDefault("\"")
			.help("Character for quoting");
		
		parser.addArgument("--sheet", "-s")
			.type(String.class)
			.required(false)
			.nargs("*")
			.help("Optional list of sheet names or sheet indexes to export (first sheet has index 1)");
		
		parser.addArgument("--drop-empty-rows", "-r")
			.action(Arguments.storeTrue())
			.help("Skip rows with only empty cells");
		
		parser.addArgument("--drop-empty-cols", "-c")
			.action(Arguments.storeTrue())
			.help("Skip columns with only empty cells");
		
		parser.addArgument("--date-as-iso", "-i")
			.action(Arguments.storeTrue())
			.help("Convert dates to ISO 8601 format and UTC standard. E.g 2020-03-28T11:40:10Z");

		parser.addArgument("--no-prefix", "-p")
			.action(Arguments.storeTrue())
			.help("Do not prefix rows with filename, sheet index, sheet name");
		
		parser.addArgument("--version", "-v").action(Arguments.version());
		
		Namespace opts= null;
		try{
			opts= parser.parseArgs(args);
		}
		catch(ArgumentParserException e) {
			parser.handleError(e);
			throw new RuntimeException();
		}		
		return(opts);
	}
}
