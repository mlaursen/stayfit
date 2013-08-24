package excel.sheet;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import enums.Macro;
import excel.CellBuilder;
import excel.Excel;

public class Intake {

	public static final int MACROS = 0;
	public static final int DATE = MACROS + 1;
	public static final int MEALS = DATE + 1;
	public static final int DEFAULT_NUM_MEALS = 6;
	public static final int EXPECTED_TOTAL = MEALS + DEFAULT_NUM_MEALS;
	public static final int MY_TOTAL = EXPECTED_TOTAL + 1;
	public static final int REMAINING_TOTAL = MY_TOTAL + 1;
	
	private Workbook wb;
	public Intake(Workbook wb) {
		this.wb = wb;
	}
	
	public void addMacroChart(Sheet s, int row, Date d) {
		Row title = s.createRow(row);
		
		makeIntakeDateCell(wb, title, d);
		//for(int i = 1; i <= DEFAULT_NUM_MEALS; i++)
		//	CellBuilder.makeTitleCell(wb, title, i + DATE, "Meal 0" + i);
		//CellBuilder.makeTitleCell(wb, title, EXPECTED_TOTAL, "Expected Total");
		//CellBuilder.makeTitleCell(wb, title, MY_TOTAL, "My Total");
		//CellBuilder.makeTitleCell(wb, title, REMAINING_TOTAL, "Remaining Total");
		
		int macroCounter = 0;
		int goTo = Macro.values().length + 1 + row;
		for(int i = row + 1; i < goTo; i++) {
			Row macro = s.createRow(i);
			
			//CellBuilder.makeTitleCell(wb, macro, MACROS, Macro.values()[macroCounter].getMacroName());
			macroCounter++;
			
			for(int j = MEALS; j < MEALS + DEFAULT_NUM_MEALS; j++) {
				//CellBuilder.makeNumberCell(wb, macro, j, 0);
			}
			
			//CellBuilder.makeNumberCell(wb, macro, EXPECTED_TOTAL, 0);
			//CellBuilder.makeNumberCell(wb, macro, MY_TOTAL, 0);
			//CellBuilder.makeNumberCell(wb, macro, REMAINING_TOTAL, 0);
			
						
			//Row refRow = settings.getRow();
		}
		/*
		Row r = s.createRow(Macro.values().length + 1);
		Cell c = r.createCell(EXPECTED_TOTAL);
		c.setCellType(Cell.CELL_TYPE_FORMULA);
		//c.setCellFormula("VLOOKUP(B1,Settings!A:B,2,TRUE)");
		c.setCellFormula(getCalorieFormula(1));
		FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
		System.out.println(eval.evaluate(c).formatAsString());
		System.out.println(c.getCellFormula());
		*/
	}
	
	public static Cell updateExpectedCalories(Workbook wb, int row) {
		Sheet s = wb.getSheet("Intake");
		Row r = s.getRow(row);
		Cell expectedCalsCell = r.getCell(EXPECTED_TOTAL);
		expectedCalsCell.setCellValue(Settings.extractSettings(wb, row).getExpectedCalories());
		return expectedCalsCell;
	}
	
	/**
	 * AL = get activity multiplier from Settings sheet
	 * W =get weight of the day
	 * US = get the unit system
	 * H = get the height
	 * A = get the age
	 * 
	 * TDEE = AL * (10 * (W * US) + 6.5 * (US * H) + 5 * A)
	 * 
	 * DT = get day type
	 * CS = get Calorie split
	 * TDEE * CS.DT
	 * 
	 * @param row
	 * @return
	 */
	public String getCalorieFormula(int row) {
		
		/*
		 * 
		 */
		
		
		return "";
	}
	
	public static Sheet updateSheet(Workbook wb) {
		Sheet s = wb.getSheet("Intake");
		
		for(int i = Excel.DATA_START; i < s.getLastRowNum(); i+=6) {
			updateExpectedCalories(wb, i);
		}
		return s;
	}
	
	public Sheet createIntakeSheet() {
		Sheet s = wb.createSheet("Intake");
		//Row titles = s.createRow(0);
		
		Calendar c = Calendar.getInstance();
		Date d = new Date();
		int chartCopies = Settings.DEFAULT_NUM_ROWS * Macro.values().length;
		for(int i = 0; i < chartCopies; i=i+6) {
			addMacroChart(s, i, d);
			updateExpectedCalories(wb, i + 1);
		}
		
		
		Excel.autosizeCols(s);
		return s;
	}
	
	public Cell makeIntakeDateCell(Workbook wb, Row r, Date d) {
		Cell c = CellBuilder.makeCell(r, DATE, d);
		c.setCellStyle(intakeDateStyle(wb));
		return c;
	}
	
	public static CellStyle intakeDateStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		Font f = wb.createFont();
		f.setBoldweight(Font.BOLDWEIGHT_BOLD);
		f.setUnderline(Font.U_SINGLE);
		cs.setFont(f);
		cs.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat(Excel.DATE_FORMAT));
		return cs;	
	}
}
