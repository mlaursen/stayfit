package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Excel {

	public static final String DEFAULT_PATH = System.getProperty("user.home") + "\\Documents";
	public static final String XSSL = ".xlsx";
	public static final String DEFAULT_WORKBOOK_NAME = "\\Nutrition Tracker" + XSSL;
	public static final String DATE_FORMAT = "dd-mmm";
	public static final String NUMBER_FORMAT = "0.00";
	public static final int DATA_START = 1;
	
	
	
	/**
	 * Autosizes the columns on a sheet
	 * @param s	Sheet to autosize
	 */
	public static void autosizeCols(Sheet s) { autosizeCols(s, s.getRow(0).getLastCellNum()); }
	public static void autosizeCols(Sheet s, int numCells) {
		for(int i = 0; i <= numCells; i++) 
			s.autoSizeColumn(i);
	}
	
	public static void writeToExcel(Workbook wb) { writeToExcel(wb, DEFAULT_PATH, DEFAULT_WORKBOOK_NAME); }
	public static void writeToExcel(Workbook wb, String path, String fName) { writeToExcel(wb, path + fName); }
	public static void writeToExcel(Workbook wb, String completeName) {
		try {
			OutputStream o = new FileOutputStream(completeName);
			wb.write(o);
			o.close();
		}
		catch (IOException e) {
			System.err.println(e.getMessage());
		}
	}
	
	
	public static Workbook readExcel(String path) throws FileNotFoundException, IOException { return readExcel(new File(path)); }
	public static Workbook readExcel(File f) throws FileNotFoundException, IOException {
		if(f.getName().charAt(f.getName().length() - 1) == 'x')
			return readExcelXSSFWorkbook(f);
		else
			return readExcelHSSFWorkbook(f);
	}
	
	public static Workbook readExcelHSSFWorkbook(File f) throws FileNotFoundException, IOException {
		return new HSSFWorkbook(new FileInputStream(f));
	}
	public static Workbook readExcelXSSFWorkbook(File f) throws FileNotFoundException, IOException {
		return new XSSFWorkbook(new FileInputStream(f));
	}
	
	public static String rowToLetter(int rowIndex) {
		String s = "";
		while(rowIndex >= 0) {
			s += rowIndex - 26 >= 0 ? "A" : (char) (rowIndex + 65);
			rowIndex -= 26;
		}
		return s;
	}
	
	
	
	
	/*
	public static void createDropDownBox(Sheet s, String[] constraints, int col) {
		createDropDown(s, constraints, Excel.DATA_START, Excel.DATA_START, col, col);
	}
	*/
	/**
	 * Creates a drop down box in Excel through data validation on an entire column
	 * starting with the start of data constant.
	 * @param s				The sheet to add the drop down box to.
	 * @param constraints	An array of String that are the possible values of the drop down box.
	 * @param startCol		The column to create the drop down box in.
	 */
	public static void createDropDown(Sheet s, String[] constraints, int startCol) {
		createDropDown(s, constraints, Excel.DATA_START, startCol);
	}
	
	/**
	 * Creates a drop down box in Excel through data validation on an entire column.  
	 * @param s				The sheet to add the drop down box to.
	 * @param constraints	An array of String that are the possible values of the drop down box.
	 * @param startCell		The cell number to start the validation.  
	 * @param startCol		The column to create the drop down box in.
	 */
	public static void createDropDown(Sheet s, String[] constraints, int startCell, int startCol) {
		int endCell = s.getLastRowNum();
		int endCol = startCol;
		createDropDown(s, constraints, startCell, endCell, startCol, endCol);
	}
	
	/**
	 * Creates a drop down box in Excel through data validation ton a range of cells.  The range is in 
	 * this form: COLUMN_ROW:COLUMN_ROW without an underscore.  Example: A2:B3
	 * 
	 * @param s				The sheet to add the drop down box validation to.
	 * @param constraints	An array of String that are the possible values of the drop down box.
	 * @param startCell		The cell number to start the validation.  This is the number 2 in the above example.
	 * @param endCell		The cell number to stop the validation.  This is the number 3 in the above example.
	 * @param startCol		The column to start the validation.  This is the letter A in the above example.
	 * @param endCol		The column to end the validation.  This is the letter B in the above example.
	 */
	public static void createDropDown(Sheet s, String[] constraints, int startCell, int endCell, int startCol, int endCol) {
		// Create the helper
		DataValidationHelper vHelper = new XSSFDataValidationHelper((XSSFSheet) s);
		
		// Create the range for the data validation to be applied to
		CellRangeAddress vRange = new CellRangeAddress(startCell, endCell, startCol, endCol);
		
		// Add the range to a list of ranges.
		CellRangeAddressList vRanges = new CellRangeAddressList();
		vRanges.addCellRangeAddress(vRange);
		
		// Create the constraint from the list of ranges
		DataValidationConstraint validationConstraint = vHelper.createExplicitListConstraint(constraints);
		
		// Finnaly, create a validation with the constraint over the list of ranges and apply to the sheet
		DataValidation validation = vHelper.createValidation(validationConstraint,  vRanges);
		s.addValidationData(validation);
	}
}
