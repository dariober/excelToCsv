package excel2csv;

import java.io.BufferedWriter;
import java.io.File;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.QuoteMode;
import org.apache.commons.text.StringEscapeUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import net.sourceforge.argparse4j.inf.Namespace;

public class Main {
	
	private static List<String> makeEmptyRow(int lastColNo, String x) {
		List<String> row = new ArrayList<String>();
    	for (int c = 0, cn = lastColNo; c < cn ; c++) {
        	row.add(x);
    	}
    	return row;
	}

	private static void printSheet(Workbook wb, String excelFile, String sheetName, CSVPrinter csvPrinter) throws IOException {
		FormulaEvaluator fe = wb.getCreationHelper().createFormulaEvaluator();
		DataFormatter formatter = new DataFormatter();
		
		Sheet sheet = wb.getSheet(sheetName);
		
		int lastColNo = Utils.getLastColNum(sheet);
		int lastRowNo = sheet.getLastRowNum();
	    
		List <String> prefix = new ArrayList<String>();
		prefix.add(excelFile);
		prefix.add(Integer.toString(wb.getSheetIndex(sheetName) + 1));
		prefix.add(sheetName);
		
		for (int r = 0; r <= lastRowNo; r++) {
	        
	    	Row row = sheet.getRow(r);
	        List<String> line = makeEmptyRow(lastColNo, null);
	        if ( row != null ) { 
		        for (int c = 0, cn = lastColNo; c < cn ; c++) {
		        	
		            Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
		            if ( cell != null ) {
		                cell = fe.evaluateInCell(cell);
		                String value = formatter.formatCellValue(cell);
		                line.set(c, value);
		            }
		        }
	        }
	        List<String> pline = new ArrayList<String>();
	        pline.addAll(prefix);
	        pline.addAll(line);
	        line.addAll(prefix);
	        csvPrinter.printRecord(pline);
	        csvPrinter.flush();
	    }
	}

	private static CSVPrinter makeCSVPrinter(String na, String delimiter, String quote) throws IOException {

		if(delimiter.length() != 1) {
			System.err.println("Delimiter must be a single character got '" + delimiter + "'");
			throw new RuntimeException();
		}
		
		CSVFormat csvFormat = CSVFormat.EXCEL
				.withEscape('\\')
        		.withNullString(na)
        		.withDelimiter(delimiter.charAt(0))
        		.withRecordSeparator('\n');
		
		if(quote.length() == 1) {
			csvFormat = csvFormat.withQuote(quote.charAt(0));
		} else if(quote.length() == 0) {
			csvFormat = csvFormat.withQuoteMode(QuoteMode.NONE);
		} else {
			System.err.println("Quote must be a single character or an empty string for no quoting");
			throw new RuntimeException();			
		}
		
		BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(System.out));

        CSVPrinter csvPrinter = new CSVPrinter(writer, csvFormat);
        return csvPrinter;
	}
	
	protected static void run(String[] args) throws IOException, InvalidFormatException {
		Namespace opts= ArgParse.argParse(args);

		List<String> input = opts.getList("input");
		for(String x : input) {
			File tmp = new File(x);
			if( ! tmp.exists() || ! Files.isReadable(tmp.toPath())) {
				System.err.println("File '" + x + "' does not exist or is not readable");
				throw new RuntimeException();
			}
		}
		
		String delimiter = StringEscapeUtils.unescapeJava(opts.getString("delimiter")); // Utils.unescapeJavaString(opts.getString("delimiter"));
		String na = opts.getString("na_string");
		String quote = opts.getString("quote");
		
		CSVPrinter csvPrinter = makeCSVPrinter(na, delimiter, quote);
    
		String excelFile = input.get(0);
		Workbook wb;
		try {
			wb = WorkbookFactory.create(new File(excelFile));
		} catch(NotOLE2FileException e) {
			System.err.println("File '" + excelFile + "' is not a valid Excel document");
			throw new RuntimeException();
		}
		
		for (int i=0; i<wb.getNumberOfSheets(); i++) {
			String sheetName = wb.getSheetName(i);
			printSheet(wb, excelFile, sheetName, csvPrinter);
		}
		
		wb.close();
		csvPrinter.close();
	}
	
	public static void main(String[] args) throws IOException, InvalidFormatException {
		try {
			run(args);
		} catch(RuntimeException e) {
			System.exit(1);
		}
	}
}
