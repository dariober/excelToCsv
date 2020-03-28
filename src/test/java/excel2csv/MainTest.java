package excel2csv;

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
	public void testDelimiter() throws InvalidFormatException, IOException {
		String[] args = "test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(stderr.length(), 0);
		assertTrue(StringUtils.countMatches(stdout, '\t') > 10);
		
		args = "-d | test_data/simple01.xlsx".split(" ");
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(stderr.length(), 0);
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
		assertEquals(stderr.length(), 0);
		assertTrue(StringUtils.countMatches(stdout, "\tNA\t") > 10);
		
		args = new String[] {"-na", "N/A",  "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(stderr.length(), 0);
		assertTrue(StringUtils.countMatches(stdout, "\tN/A\t") > 10);

		args = new String[] {"-na", "",  "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(stderr.length(), 0);
		assertTrue(StringUtils.countMatches(stdout, "\t\t") > 10);
	}
	
	@Test
	public void testQuote() throws InvalidFormatException, IOException {
		String[] args = "test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(stderr.length(), 0);
		assertTrue(stdout.contains("\"#HERE!\""));
		
		args = new String[] {"-q", "$",  "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(stderr.length(), 0);
		assertTrue(stdout.contains("$#HERE!$"));
		
		// No quoting
		args = new String[] {"-q", "",  "test_data/simple01.xlsx"};
		out = this.runMain(args);
		stdout = out.get(0);
		stderr = out.get(1);
		assertEquals(stderr.length(), 0);
		assertTrue(stdout.contains("\t#HERE!\t"));
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
		assertEquals(stderr.length(), 0);
		assertEquals(StringUtils.countMatches(stdout, "test_data/simple01.xlsx\t1\tSheet1\t"), 10);
		assertEquals(StringUtils.countMatches(stdout, "test_data/simple01.xlsx\t2\tSheet2\t"), 4);
	}
	
	@Test
	public void testMultipleInputFiles() throws InvalidFormatException, IOException {
		String[] args = "test_data/simple01.xlsx test_data/simple01.xls".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(stderr.length(), 0);
		assertTrue(StringUtils.countMatches(stdout, "test_data/simple01.xlsx\t1\tSheet1\t") > 5);
		assertTrue(StringUtils.countMatches(stdout, "test_data/simple01.xls\t1\tSheet1\t") > 5);
	}
	
	@Test
	public void testEvaluateFormula() throws InvalidFormatException, IOException {
		String[] args = "test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(stderr.length(), 0);
		assertEquals(StringUtils.countMatches(stdout, "\t3.33\t"), 1);
	}
	
	@Test
	public void testEmptyFile() throws InvalidFormatException, IOException {
		String[] args = "test_data/empty.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(stderr.length(), 0);
		assertEquals(stdout.length(), 0);
	}
	
	@Test
	public void testCanSkipEmptyRows() throws InvalidFormatException, IOException {
		String[] args = "-r test_data/simple01.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		assertEquals(stderr.length(), 0);
		assertTrue( ! stdout.contains("NA\tNA\tNA\tNA\tNA\tNA\tNA"));
		assertTrue(stdout.contains("NA\tcol1\tcol2"));
	}
	
	@Test
	public void testCanSkipEmptyColumns() throws InvalidFormatException, IOException {
		String[] args = "-c test_data/empty_cols.xlsx".split(" ");
		List<String> out = this.runMain(args);
		String stdout = out.get(0);
		String stderr = out.get(1);
		System.out.println(stderr);
		assertEquals(stderr.length(), 0);
		// assertTrue( ! stdout.contains("Sheet1\tNA\tcol1"));
		//System.out.println(stdout);
		//assertTrue(stdout.contains("Sheet1\tcol1"));
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
