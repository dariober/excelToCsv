package excelToCsv;

import static org.junit.Assert.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

import org.supercsv.io.CsvListWriter;
import org.supercsv.prefs.CsvPreference;

public class MainTest {

	@Test
	public void testSuperCSV() throws IOException {
		
		CsvPreference csvFormat = new CsvPreference.Builder('"', '|', "\n")
					.surroundingSpacesNeedQuotes(false)
					.build();
		
		CsvListWriter listWriter = new CsvListWriter(new OutputStreamWriter(System.out),
                 csvFormat);
         
		 String[] record = new String[] {"eggs", "foo, \"bar\"  ", "spam"};
		 listWriter.write(record);
		 listWriter.flush();
		 listWriter.close();
	}
	
	@Test
	public void testDates() throws InvalidFormatException, IOException {
		String[] args = "test_data/dates.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stderr = out.get(1);
		String stdout = out.get(0);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("21/03/20"));
		assertTrue(stdout.contains("18/10/1933 12:36:00"));
		
		args = "-i test_data/dates.xlsx".split(" ");
		out = this.runMain(args);
		stderr = out.get(1);
		stdout = out.get(0);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("2020-03-21T00:00:00Z"));
		assertTrue(stdout.contains("1933-10-18T12:36:00Z"));
	}
	
	@Test
	public void testRequestSheets() throws InvalidFormatException, IOException {
		String[] args = "-s Sheet1 -- test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("Sheet1"));
		assertTrue( ! stdout.contains("Sheet2"));
	
		args = "-s Sheet2 -- test_data/simple01.xlsx".split(" ");
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("Sheet2"));
		assertTrue( ! stdout.contains("Sheet1"));
		
		args = "-s 2 -- test_data/simple01.xlsx".split(" ");
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("Sheet2"));
		assertTrue( ! stdout.contains("Sheet1"));
		
		args = new String[] {"-s", "Sheet2", "Sheet Dates", "FOOBAR", "--", "test_data/dates.xlsx", "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("Sheet2"));
		assertTrue( ! stdout.contains("Sheet1"));
		assertTrue(stdout.contains("Sheet Dates"));
		
		args = new String[] {"-s", "2", "Sheet1", "--", "test_data/dates.xlsx", "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("Sheet2"));
		assertTrue(stdout.contains("Sheet1"));
		assertTrue( ! stdout.contains("Sheet Dates"));
		
		args = new String[] {"-s", "99", "--", "test_data/dates.xlsx", "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertEquals(0, stdout.length());
	}
	
	@Test 
	public void testSize() throws InvalidFormatException, IOException {
		String[] args = "-na NA test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		String[] rows = stdout.split("\n");
		assertEquals(14, rows.length);
		for(String row : rows) {
			if(row.contains("Sheet1")) {
				assertEquals(3+7, row.split("\t").length);
			}
		}
	}
	
	@Test
	public void testDelimiter() throws InvalidFormatException, IOException {
		String[] args = "test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(StringUtils.countMatches(stdout, '\t') > 10);
		
		args = "-d | test_data/simple01.xlsx".split(" ");
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(StringUtils.countMatches(stdout, '|') > 10);	
	}

	@Test
	public void testInvalidDelimiter() throws InvalidFormatException, IOException {
		boolean pass = false;
		try {
			String[] args = "-d foo test_data/simple01.xlsx".split(" ");
			this.runMain(args);
		} catch(RuntimeException e){
			pass = true;
		}
		assertTrue(pass);
	}
	
	@Test
	public void testNAString() throws InvalidFormatException, IOException {
		String[] args = "test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(StringUtils.countMatches(stdout, "\t\t") > 10);
		
		args = new String[] {"-na", "N/A",  "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(StringUtils.countMatches(stdout, "\tN/A\t") > 10);

		args = new String[] {"-na", "",  "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(StringUtils.countMatches(stdout, "\t\t") > 10);
	}
	
	@Test
	public void testQuote() throws InvalidFormatException, IOException {
		String[] args = "test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("\t#HERE!\t"));

		args = new String[] {"-q", "#",  "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("\t###HERE!#\t"));
		
		args = new String[] {"-d", ",", "test_data/quotes.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		System.out.println(stdout);
		assertEquals("test_data/quotes.xlsx,1,Sheet1,eggs,\"Foo, \"\"bar\"\", eggs\",spam with traling space  ,bob", stdout.trim());		
	}
	
	@Test
	public void testInvalidQuote() throws InvalidFormatException, IOException {
		boolean pass = false;
		try {
			String[] args = "-q foo test_data/simple01.xlsx".split(" ");
			this.runMain(args);
		} catch(RuntimeException e){
			pass = true;
		}
		assertTrue(pass);
		
		pass = false;
		try {
			String[] args = "-q '' test_data/simple01.xlsx".split(" ");
			this.runMain(args);
		} catch(RuntimeException e){
			pass = true;
		}
		assertTrue(pass);
	}
	
	@Test
	public void testFileDoesNotExist() throws InvalidFormatException, IOException {
		boolean pass = false;
		try {
			String[] args = "foobar.xls".split(" ");
			this.runMain(args);
		} catch(RuntimeException e){
			pass = true;
		}
		assertTrue(pass);
	}
	
	@Test
	public void testNotExcelFile() throws InvalidFormatException, IOException {
		boolean pass = false;
		try {
			String[] args = "test_data/not_excel.txt".split(" ");
			this.runMain(args);
		} catch(RuntimeException e){
			pass = true;
		}
		assertTrue(pass);
	}
	
	@Test
	public void testRowPrefix() throws InvalidFormatException, IOException {
		String[] args = "test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertEquals(StringUtils.countMatches(stdout, "test_data/simple01.xlsx\t1\tSheet1\t"), 10);
		assertEquals(StringUtils.countMatches(stdout, "test_data/simple01.xlsx\t2\tSheet2\t"), 4);
	
		args = "-p -na NA test_data/simple01.xlsx".split(" ");
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.startsWith("NA\t"));
	}
	
	@Test
	public void testMultipleInputFiles() throws InvalidFormatException, IOException {
		String[] args = "test_data/simple01.xlsx test_data/simple01.xls".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(StringUtils.countMatches(stdout, "test_data/simple01.xlsx\t1\tSheet1\t") > 5);
		assertTrue(StringUtils.countMatches(stdout, "test_data/simple01.xls\t1\tSheet1\t") > 5);
	}
	
	@Test
	public void testEvaluateFormula() throws InvalidFormatException, IOException {
		String[] args = "test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertEquals(1, StringUtils.countMatches(stdout, "\t3.33\t"));
	}
	
	@Test
	public void testEmptyFile() throws InvalidFormatException, IOException {
		String[] args = "test_data/empty.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertEquals(stdout.length(), 0);
	}
	
	@Test
	public void testCanSkipEmptyRows() throws InvalidFormatException, IOException {
		String[] args = "-na NA -r test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue( ! stdout.contains("NA\tNA\tNA\tNA\tNA\tNA\tNA"));
		assertTrue(stdout.contains("NA\tcol1\tcol2"));
	}
	
	@Test
	public void testCanSkipEmptyColumns() throws InvalidFormatException, IOException {
		String[] args = "-d | -c test_data/empty_cols.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("test_data/empty_cols.xlsx|1|Sheet1|a|e|h"));
	}
	
	/** Execute main with the given array of arguments and return a list of length 2 containing 1) stdout and 2) stderr.
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 * */
	private List<String> runMain(String[] args) throws InvalidFormatException, IOException {
		PrintStream stdout= System.out;
		ByteArrayOutputStream baosOut= new ByteArrayOutputStream();
		System.setOut(new PrintStream(baosOut));

		PrintStream stderr= System.err;
		ByteArrayOutputStream baosErr= new ByteArrayOutputStream();
		System.setErr(new PrintStream(baosErr));

		Main.run(args);

		String out= baosOut.toString();
	    System.setOut(stdout);

		String err= baosErr.toString();
	    System.setErr(stderr);

	    List<String> outErr= new ArrayList<String>();
	    outErr.add(out);
	    outErr.add(err);
	    return outErr;
	}
}
