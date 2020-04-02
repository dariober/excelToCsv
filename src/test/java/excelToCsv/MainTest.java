package excelToCsv;

import static org.junit.Assert.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.Test;

public class MainTest {
	
	@Test
	public void testNoFormat() throws InvalidFormatException, IOException {
		
		String[] args = new String[] {"-d",  ",", "-f", "-i", "test_data/format.xlsx"};
		List<String> out = this.runMain(args);
		String stderr = out.get(1);
		String stdout = out.get(0);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains(",60,"));
		assertTrue(stdout.contains(",200.666661562376,"));
		assertTrue(stdout.contains(",21/03/20,"));
		assertTrue(stdout.contains(",TRUE,"));
		assertTrue(stdout.contains(",1.23,"));
	}
	
	@Test
	public void testFormat() throws InvalidFormatException, IOException {
		
		String[] args = new String[] {"-d",  ",", "-i", "test_data/format.xlsx"};
		List<String> out = this.runMain(args);
		String stderr = out.get(1);
		String stdout = out.get(0);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains(",2.66666156237642,"));
		assertTrue(stdout.contains(",5.33332312475284,"));
		assertTrue(stdout.contains(",60,"));
		assertTrue(stdout.contains(",0.0148290494731055,"));
		assertTrue(stdout.contains(",2.7,"));
		assertTrue(stdout.contains(",2.01E+02,"));
		assertTrue(stdout.contains(",TRUE,"));
		assertTrue(stdout.contains(",21/03/20,"));
		assertTrue(stdout.contains(",123.00%,"));
	}
	
	@Test
	public void testDates() throws InvalidFormatException, IOException {
		String[] args = "-i test_data/dates.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stderr = out.get(1);
		String stdout = out.get(0);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("21/03/20"));
		assertTrue(stdout.contains("18/10/1933 12:36:00"));
		
		args = "-I -i test_data/dates.xlsx".split(" ");
		out = this.runMain(args);
		stderr = out.get(1);
		stdout = out.get(0);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("2020-03-21T00:00:00Z"));
		assertTrue(stdout.contains("1933-10-18T12:36:00Z"));
	}
	
	@Test
	public void testRequestSheets() throws InvalidFormatException, IOException {
		String[] args = "-sn Sheet1 -i test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("Sheet1"));
		assertTrue( ! stdout.contains("Sheet2"));
	
		args = "-sn Sheet2 -i test_data/simple01.xlsx".split(" ");
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("Sheet2"));
		assertTrue( ! stdout.contains("Sheet1"));
		
		args = "-si 2 -i test_data/simple01.xlsx".split(" ");
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("Sheet2"));
		assertTrue( ! stdout.contains("Sheet1"));
		
		args = new String[] {"-sn", "Sheet2", "Sheet Dates", "FOOBAR", "-i", "test_data/dates.xlsx", "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("Sheet2"));
		assertTrue( ! stdout.contains("Sheet1"));
		assertTrue(stdout.contains("Sheet Dates"));
		
		args = new String[] {"-si", "2", "-sn", "Sheet1", "-i", "test_data/dates.xlsx", "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains("Sheet2"));
		assertTrue(stdout.contains("Sheet1"));
		assertTrue( ! stdout.contains("Sheet Dates"));
		
		args = new String[] {"-si", "99", "-i", "test_data/dates.xlsx", "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertEquals(0, stdout.length());

		boolean pass = false;
		try {
			args = new String[] {"-si", "0", "-i", "test_data/dates.xlsx", "test_data/simple01.xlsx"};
			this.runMain(args);
		} catch(RuntimeException e){
			pass = true;
		}
		assertTrue(pass);
	}
	
	@Test 
	public void testSize() throws InvalidFormatException, IOException {
		String[] args = "-na NA -i test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		String[] rows = stdout.split("\n");
		assertEquals(14, rows.length);
		for(String row : rows) {
			if(row.contains("Sheet1")) {
				assertEquals(3+7, row.split(",").length);
			}
		}
	}
	
	@Test
	public void testDelimiter() throws InvalidFormatException, IOException {
		String[] args = "-d \t -i test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(StringUtils.countMatches(stdout, '\t') > 10);
		
		args = "-d | -i test_data/simple01.xlsx".split(" ");
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
			String[] args = "-d foo -i test_data/simple01.xlsx".split(" ");
			this.runMain(args);
		} catch(RuntimeException e){
			pass = true;
		}
		assertTrue(pass);
	}
	
	@Test
	public void testNAString() throws InvalidFormatException, IOException {
		String[] args = new String[] {"-d", "\t", "-i", "test_data/simple01.xlsx"};
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(StringUtils.countMatches(stdout, "\t\t") > 10);
		
		args = new String[] {"-na", "N/A", "-d", "\t", "-i", "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(StringUtils.countMatches(stdout, "\tN/A\t") > 10);

		args = new String[] {"-na", "", "-d", "\t", "-i", "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(StringUtils.countMatches(stdout, "\t\t") > 10);
	}
	
	@Test
	public void testQuote() throws InvalidFormatException, IOException {
		String[] args = "-i test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains(",#HERE!,"));

		args = new String[] {"-q", "#",  "-i", "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.contains(",###HERE!#,"));
		
		args = new String[] {"-d", ",", "-i", "test_data/quotes.xlsx"};
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
			String[] args = "-q foo -i test_data/simple01.xlsx".split(" ");
			this.runMain(args);
		} catch(RuntimeException e){
			pass = true;
		}
		assertTrue(pass);
		
		pass = false;
		try {
			String[] args = "-q '' -i test_data/simple01.xlsx".split(" ");
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
			String[] args = "-i foobar.xls".split(" ");
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
			String[] args = "-i test_data/not_excel.txt".split(" ");
			this.runMain(args);
		} catch(RuntimeException e){
			pass = true;
		}
		assertTrue(pass);
	}
	
	@Test
	public void testRowPrefix() throws InvalidFormatException, IOException {
		String[] args = "-i test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertEquals(StringUtils.countMatches(stdout, "test_data/simple01.xlsx,1,Sheet1,"), 10);
		assertEquals(StringUtils.countMatches(stdout, "test_data/simple01.xlsx,2,Sheet2,"), 4);
	
		args = "-p -na NA -i test_data/simple01.xlsx".split(" ");
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(stdout.startsWith("NA,"));
	}
	
	@Test
	public void testMultipleInputFiles() throws InvalidFormatException, IOException {
		String[] args = "-i test_data/simple01.xlsx test_data/simple01.xls".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue(StringUtils.countMatches(stdout, "test_data/simple01.xlsx,1,Sheet1,") > 5);
		assertTrue(StringUtils.countMatches(stdout, "test_data/simple01.xls,1,Sheet1,") > 5);
	}
	
	@Test
	public void testEvaluateFormula() throws InvalidFormatException, IOException {
		String[] args = "-i test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		
		assertEquals(0, stderr.length());
		assertEquals(1, StringUtils.countMatches(stdout, ",3.33,"));
	}
	
	@Test
	public void testEmptyFile() throws InvalidFormatException, IOException {
		String[] args = "-i test_data/empty.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertEquals(stdout.length(), 0);
	}
	
	@Test
	public void testCanSkipEmptyRows() throws InvalidFormatException, IOException {
		String[] args = "-na NA -r -i test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(0, stderr.length());
		assertTrue( ! stdout.contains("NA,NA,NA,NA,NA,NA,NA"));
		assertTrue(stdout.contains("NA,col1,col2"));
	}
	
	@Test
	public void testCanSkipEmptyColumns() throws InvalidFormatException, IOException {
		String[] args = "-d | -c -i test_data/empty_cols.xlsx".split(" ");
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
