package excel2csv;

import java.io.File;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.nio.file.Files;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import org.apache.commons.text.StringEscapeUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.supercsv.io.CsvListWriter;
import org.supercsv.prefs.CsvPreference;

import net.sourceforge.argparse4j.inf.Namespace;

public class Main {
	
	private static List<String> makeEmptyRow(int lastColNo, String x) {
		List<String> row = new ArrayList<String>();
    	for (int c = 0, cn = lastColNo; c < cn ; c++) {
        	row.add(x);
    	}
    	return row;
	}

	private static Set<Integer> indexOfEmptyColumns(Sheet sheet){
		
		// Get indexes of non-empty columns
		Set<Integer> nonEmpty = new HashSet<Integer>();		
	    for (int r = 0; r <= sheet.getLastRowNum(); r++) {
	        Row row = sheet.getRow(r);
	        if(row == null) {
	        	continue;
	        }
	        for(int c = 0; c < row.getLastCellNum(); c++) {
	        	if(row.getCell(c) != null) {
	        		nonEmpty.add(c);
	        	}
	        }
	    }
	    
		int maxCol = 0;
		for(int i : nonEmpty) {
			if(i > maxCol) {
				maxCol = i;
			}
		}
		
		// Add to output array indexes non in nonEmpty:
		Set<Integer> emptyIdx = new HashSet<Integer>();
		for(int i = 0; i < maxCol; i++) {
			if(!nonEmpty.contains(i)) {
				emptyIdx.add(i);
			}
		}
		return emptyIdx;
	}
	
	private static void printSheet(Workbook wb, String excelFile, String sheetName, 
			CsvListWriter listWriter, String na, boolean dropEmptyRows, boolean dropEmptyColumns, boolean dateAsIso, boolean noPrefix) throws IOException {

		FormulaEvaluator fe = wb.getCreationHelper().createFormulaEvaluator();
		DataFormatter formatter = new DataFormatter();
		
		Sheet sheet = wb.getSheet(sheetName);
		
		int lastColNo = Utils.getLastColNum(sheet);
		int lastRowNo = sheet.getLastRowNum();
	    
		List <String> prefix = new ArrayList<String>();
		prefix.add(excelFile);
		prefix.add(Integer.toString(wb.getSheetIndex(sheetName) + 1));
		prefix.add(sheetName);

		Set<Integer> emptyColsIdx = new HashSet<Integer>();
		if(dropEmptyColumns) {
			emptyColsIdx = indexOfEmptyColumns(sheet);
		}
		
		for (int r = 0; r <= lastRowNo; r++) {
	        
	    	Row row = sheet.getRow(r);
	        if(row == null && dropEmptyRows) {
	        	continue;
	        }
	    	List<String> line = makeEmptyRow(lastColNo, na);
	        if ( row != null ) { 
		        for (int c = 0, cn = lastColNo; c < cn ; c++) {
		        	
		            Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
		            if ( cell != null ) {
		                cell = fe.evaluateInCell(cell);
		                String value;
		                if(dateAsIso && cell.getCellType().equals(CellType.NUMERIC) && DateUtil.isCellDateFormatted(cell)) {
		                	Date d = cell.getDateCellValue();
		                	value = d.toInstant().atZone(ZoneId.of("UTC")).format(DateTimeFormatter.ISO_INSTANT);
		                } else {
		                	value = formatter.formatCellValue(cell);
		                }
		                line.set(c, value);
		            } else {
		            	line.set(c, na);
		            }
		        }
	        }
	        List<String> pline = new ArrayList<String>();
	        if(!noPrefix) {
	        	pline.addAll(prefix);
	        }
	        for(int i = 0; i < line.size(); i++) {
	        	if(! emptyColsIdx.contains(i)) {
	        		pline.add(line.get(i));
	        	}
	        }
	        listWriter.write(pline);
	        listWriter.flush();
	    }
	}

	private static CsvListWriter makeCsvListWriter(String delimiter, String quote) {
		
		if(delimiter.length() != 1) {
			System.err.println("Delimiter must be a single character got '" + delimiter + "'");
			throw new RuntimeException();
		}
		
		if(quote.length() != 1) {
			System.err.println("Quote must be a single character or an empty string for no quoting");
			throw new RuntimeException();			
		}
		
		CsvPreference csvFormat = new CsvPreference.Builder(quote.charAt(0), delimiter.charAt(0), "\n")
				.surroundingSpacesNeedQuotes(false)
				.build();
	
		CsvListWriter listWriter = new CsvListWriter(new OutputStreamWriter(System.out),
	             csvFormat);
		return listWriter;
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
		boolean dropEmptyRows = opts.getBoolean("drop_empty_rows");
		boolean dropEmptyCols = opts.getBoolean("drop_empty_cols");
		boolean dateAsIso = opts.getBoolean("date_as_iso");
		boolean noPrefix = opts.getBoolean("no_prefix");
		List<String> requestSheets = opts.getList("sheet");
		
		CsvListWriter listWriter = makeCsvListWriter(delimiter, quote);
		
		for(String excelFile : input) {
			Workbook wb;
			try {
				wb = WorkbookFactory.create(new File(excelFile));
			} catch(NotOLE2FileException e) {
				System.err.println("File '" + excelFile + "' is not a valid Excel document");
				throw new RuntimeException();
			}
			
			for (int i=0; i<wb.getNumberOfSheets(); i++) {
				String sheetName = wb.getSheetName(i);
				boolean print = isRequestedSheet(requestSheets, sheetName, wb.getSheetIndex(sheetName));
				if(print) {
					printSheet(wb, excelFile, sheetName, listWriter, na, dropEmptyRows, dropEmptyCols, dateAsIso, noPrefix);
				}
			}
			
			wb.close();
		}
		listWriter.close();
	}
	
	private static boolean isRequestedSheet(List<String> reqsheets, String sheetName, int sheetIndex) {
		if(reqsheets == null || reqsheets.size() == 0) {
			return true;
		}
		if(reqsheets.contains(sheetName)) {
			return true;
		}
		for(String x : reqsheets) {
			int idx;
			try {
				idx = Integer.parseInt(x) - 1;
				if(idx == sheetIndex) {
					return true;
				}
			} catch( NumberFormatException e) {
				//
			}
		}
		return false;
	}

	public static void main(String[] args) throws IOException, InvalidFormatException {
		try {
			run(args);
		} catch(RuntimeException e) {
			System.exit(1);
		}
	}
}
