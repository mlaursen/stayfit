package excel.sheet;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.LocalDate;

import user.User;
import excel.CellBuilder;
import excel.CellStyles;
import excel.Excel;
import excel.sheet.cell.Formulas;

public class Settings {
	
	public static final Map<String, Integer> COLS = new HashMap<String, Integer>();
	public static final Map<String, Integer> ROWS = new HashMap<String, Integer>();
	static {
		int i = 0;
		COLS.put("DOW", i++);
		COLS.put("Date", i++);
		COLS.put("Weight", i++);
		COLS.put("EWMA5", i++);
		COLS.put("EWMA7", i++);
		COLS.put("Smoothed", i++);
		COLS.put("Forecast", i++);
		COLS.put("Residual", i++);
		COLS.put("Lost/wk", i++);
		COLS.put("Trend", i++);
		COLS.put("Slope", i++);
		COLS.put("C/F Split", i++);
		COLS.put("Change", i++);
		COLS.put("BMR", i++);
		COLS.put("TDEE", i++);
		COLS.put("BMR?", i++);
		COLS.put("Calories", i++);
		COLS.put("P*", i++);
		COLS.put("Protein", i++);
		COLS.put("Assumed Constants", ++i);	// inc before
		COLS.put("Birthday", i++);
		COLS.put("Height", i++);
		COLS.put("Activity", i++);
		COLS.put("Units", i++);
		COLS.put("Gender", i++);
		COLS.put("Alpha", i++);
		COLS.put("Activity Names", ++i);	// table for storing mult vals
		COLS.put("Split", i++);
		COLS.put("Activity Multipliers", i);
		COLS.put("Carbs", i++);
		COLS.put("Fat", i++);
		
		ROWS.put("Constants", 1);
		ROWS.put("Constant Values", 2);
		
		ROWS.put("Sedentary", 1);
		ROWS.put("Lightly Active", 2);
		ROWS.put("Moderately Active", 3);
		ROWS.put("Very Active", 4);
		ROWS.put("Extremely Active", 5);
		
		ROWS.put("Split Table", 7);
	}
	
	public static final int DEFAULT_NUM_ROWS = 12*7; // 12 weeks
	public static final double DEFAULT_HEIGHT = 71.0;
	public static final int DEFAULT_AGE = 22;
	
	public static String REF_SETTINGS = "Settings!";
	
	private Workbook wb;
	public Settings(Workbook wb) {
		this.wb = wb;
	}

	
	public Sheet createSettingsSheet() {
		Sheet s = this.wb.createSheet("Settings");
		Row titles = s.createRow(0);
		createCell(titles, "DOW");
		createCell(titles, "Date");
		createCell(titles, "Weight");
		createCell(titles, "EWMA5");
		createCell(titles, "EWMA7");
		createCell(titles, "Smoothed");
		createCell(titles, "Forecast");
		createCell(titles, "Residual");
		createCell(titles, "Lost/wk");
		createCell(titles, "Trend");
		createCell(titles, "Slope");
		createCell(titles, "C/F Split");
		createCell(titles, "Change");
		createCell(titles, "BMR");
		createCell(titles, "TDEE");
		createCell(titles, "BMR?");
		createCell(titles, "Calories");
		createCell(titles, "P*");
		createCell(titles, "Protein");
		createCell(titles, "Assumed Constants");
		createCell(titles, "Activity Names");
		
		
		LocalDate d = new LocalDate();
		for(int i = Excel.DATA_START; i <= DEFAULT_NUM_ROWS; i++) {
			Row r = s.createRow(i);
			createFormulaCell(r, "DOW", Formulas.dowFormula(r.getRowNum()));
			createCell(r, "Date", d);
			d = d.plusDays(1);
			createCell(r, "Weight", "");
			createCell(r, "EWMA5", "");
			createCell(r, "EWMA7", "");
			createCell(r, "Smoothed", "");
			createCell(r, "Forecast", "");
			createCell(r, "Residual", "");
			createCell(r, "Lost/wk", "");
			createCell(r, "Trend", "");
			createCell(r, "Slope", "");
			createCell(r, "C/F Split", "");
			createCell(r, "Change", "");
			createCell(r, "BMR", "");
			createCell(r, "TDEE", "");
			createCell(r, "BMR?", "");
			createCell(r, "Calories", "");
			createCell(r, "P*", "");
			createCell(r, "Protein", "");
		
		}
		/*
		 * just commented this 8/23
		
		Row r = s.getRow(ROWS.get("Split Table"));
		createCell(r, "Split");
		createCell(r, "Carbs");
		createCell(r, "Fat");
		
		r = s.getRow(ROWS.get("Constants"));
		CellBuilder.makeCell(r, COLS.get("Birthday"), "Birthday");
		CellBuilder.makeCell(r, COLS.get("Height"), "Height");
		CellBuilder.makeCell(r, COLS.get("Activity"), "Activity Multiplier");
		CellBuilder.makeCell(r, COLS.get("Units"), "Units");
		CellBuilder.makeCell(r, COLS.get("Gender"), "Gender");
		CellBuilder.makeCell(r, COLS.get("Alpha"), "alpha");
		*/
		//makeDateCell(r, COLS.get("Birthday"));
		//CellBuilder.makeNumberCell(wb, r, COLS.get("Height"));
		//Excel.createDropDown(s, enums.ActivityMultiplier.ACTIVITY_MULTIPLIER, CONSTS, CONSTS, COLS.get("Activity"), COLS.get("Activity"));
		
		/*
		old comment
		
		Calendar c = Calendar.getInstance();
		Date d = new Date();
		for(int i = Excel.DATA_START; i <= DEFAULT_NUM_ROWS; i++) {
			Row r = s.createRow(i);
			makeDateCell(wb, r, DATE, d);
			makeWeightCell(wb, r);
			makeDayTypeCell(r);
			makeCalorieSplitCell(r);
			makeCarbFatSplitCell(r);
			c.setTime(d);
			c.add(Calendar.DATE, 1);
			d = c.getTime();
		}
		
		Row r = s.getRow(Excel.DATA_START);
		CellBuilder.makeNumberCell(wb, r, HEIGHT, DEFAULT_HEIGHT);
		CellBuilder.makeNumberCell(wb, r, AGE, DEFAULT_AGE);
		makeUnitSystemCell(r);
		makeActivityMultiplierCell(r);
		makeFormulaCell(r);
		makeGenderCell(r);
		
		// Create drop down boxes
		Excel.createDropDown(s, DayType.DAY_TYPES, DAY_TYPE);
		Excel.createDropDown(s, CalorieSplit.CALORIE_SPLIT, CALORIE_SPLIT);
		Excel.createDropDown(s, CarbFatSplit.CARB_FAT_SPLIT, CARB_FAT_SPLIT);
		Excel.createDropDownBox(s, Formula.FORMULA, FORMULA);
		Excel.createDropDownBox(s, UnitSystem.UNIT_SYSTEM, UNIT_SYSTEM);
		Excel.createDropDownBox(s, ActivityMultiplier.ACTIVITY_MULTIPLIER, ACTIVITY_MULTIPLIER);
		Excel.createDropDownBox(s, Sex.SEX, SEX);
		*/
		
		Excel.autosizeCols(s);
		return s;
	}
	
	
	public static User extractSettings(Workbook wb, int row) {
		Sheet s = wb.getSheet("Settings");
		
		Row constantsRow = s.getRow(Excel.DATA_START);
		row = row / 6 + 1;
		
		/*
		Row updatingRow = s.getRow(row);
		double weight = updatingRow.getCell(WEIGHT).getNumericCellValue();
		double height = constantsRow.getCell(HEIGHT).getNumericCellValue();
		double age = constantsRow.getCell(AGE).getNumericCellValue();
		UnitSystem units = UnitSystem.parseRichTextString(constantsRow.getCell(UNIT_SYSTEM).getRichStringCellValue());
		ActivityMultiplier activity = ActivityMultiplier.parseRichTextString(constantsRow.getCell(ACTIVITY_MULTIPLIER).getRichStringCellValue());
		CalorieSplit calSplit = CalorieSplit.parseRichTextString(updatingRow.getCell(CALORIE_SPLIT).getRichStringCellValue());
		DayType type = DayType.parseRichTextString(updatingRow.getCell(DAY_TYPE).getRichStringCellValue());
		CarbFatSplit carbFatSplit = CarbFatSplit.parseRichTextString(updatingRow.getCell(CARB_FAT_SPLIT).getRichStringCellValue());
		Formula f = Formula.parseRichTextString(constantsRow.getCell(FORMULA).getRichStringCellValue());
		Sex sex = Sex.parseRichTextString(constantsRow.getCell(SEX).getRichStringCellValue());
		*/
		return null; //new User(weight, height, (int) age, units, type, carbFatSplit, activity, calSplit, f, sex);
	}
	
	public Cell createFormulaCell(Row r, String n, String f) {
		return createCell(CellBuilder.makeFormulaCell(r, COLS.get(n), f), n, r.getRowNum());
	}
	public Cell createCell(Row r, String n) { return createCell(r, n, n); }
	public Cell createCell(Row r, String n, Object v) {
		int rn = r.getRowNum();
		int i = COLS.get(n);
		Cell c = CellBuilder.makeCell(r, i, v);
		if(n == "Date" && rn >= Excel.DATA_START) {
			c = CellBuilder.makePrevPlus1Cell(r, i);
			if(rn == Excel.DATA_START) {
				c = r.createCell(i); 
				c.setCellValue(((LocalDate) v).toDate());
			}
		}
		else if(n.contains("EWMA")) {
			int x = Integer.parseInt(n.substring(n.length()-1));
			if(rn > x)
				c = CellBuilder.makeFormulaCell(r, i, Formulas.ewmaFormula(rn, x));
		}
		return createCell(c, n, rn);
	}
	public Cell createCell(Cell c, String n, int rn) {
		Set<Short> styles = addStyles(n, rn);
		CellStyles.applyStyles(styles, c);
		return c;
	}
	
	public Set<Short> addStyles(String n, int rn) {
		Set<Short> styles = new HashSet<Short>();
		if(COLS.containsKey(n) && rn == Excel.DATA_START - 1) {
			styles.add(CellStyles.BOLD);
			styles.add(CellStyles.BORDER_BOTTOM);
		}
		
		if(n == "Weight" || n == "EWMA7" || n == "Residual" || n == "Lost/wk" || n == "Slope" || n == "C/F Split"
		|| n == "Change"|| n == "TDEE"|| n == "BMR?"|| n == "Calories"|| n == "P*" || n == "Protein") {
			styles.add(CellStyles.BORDER_RIGHT);
		}
		
		if(rn % 7 == 0)
			styles.add(CellStyles.BORDER_BOTTOM);
		
		if(n.contains("EWMA")) {
			int e = Integer.parseInt(n.substring(n.length()-1));
			//System.out.println("ewma");
			//System.out.println(e);
			if(rn >= Excel.DATA_START && (rn <= 5 || (e == 7 && rn <= e))) 
				styles.add(CellStyles.GRAY_FILL);
			if(e == 5 && rn > 5 && rn < 8)
				styles.add(CellStyles.BORDER_RIGHT_THIN);
			
			if(e == 7)
				styles.add(CellStyles.BORDER_RIGHT);
			
			if(rn == e && e == 5)
				styles.add(CellStyles.BORDER_BOTTOM_THIN);
		}
		else if(n.equals("Smoothed") || n.equals("Forecast") || n.equals("Residual")) {
			if(rn <= 5 && rn >= Excel.DATA_START)
				styles.add(CellStyles.GRAY_FILL);
			
			if(rn == 5)
				styles.add(CellStyles.BORDER_BOTTOM_THIN);
			
			if(n.equals("Residual"))
				styles.add(CellStyles.BORDER_RIGHT);
		}
		else if(n.equals("Date")) {
			styles.add(CellStyles.DATE);
		}
		//System.out.println(styles);
		return styles;
	}
	
	
	
	
	public Cell makeSmoothedCell(Row r) {
		Cell c = CellBuilder.makeCell(r, COLS.get("Smoothed"), "");
		int rn = r.getRowNum();
		List<String> borders = new ArrayList<String>();
		if(rn == 5)
			borders.add("b");
		
		if(rn <= 5) { 
			//c.setCellStyle(CellStyles.grayFillBorderStyle(wb, borders));
		}
		else {
			String f = "(" + getConst("Alpha") + "*" + getCol("Weight") + rn + ")";
			f += "+((1-" + getConst("Alpha") + ")*" + getCol("Forecast") + rn + ")";
			String f2 = "IF(" + getCol("Weight") + (rn+1) + "=\"\",NA()," + f + ")";
			c.setCellFormula(f2);
			//c.setCellStyle(CellStyles.numberStyle(wb, true, 4));
		}
		
		return c;
	}
	
	public Cell makeForecastCell(Row r) {
		Cell c = CellBuilder.makeCell(r, COLS.get("Forecast"), "");
		int rn = r.getRowNum();
		List<String> borders = new ArrayList<String>();
		if(rn == 5)
			borders.add("b");
		
		if(rn <= 5) {
			//c.setCellStyle(CellStyles.grayFillBorderStyle(wb, borders));
		}
		else {
			
		}
		return c;
	}
	
	public static String getCol(String c) {
		return "$" + Excel.rowToLetter(COLS.get(c));
	}
	
	public static String getConst(String c) {
		return "$" + Excel.rowToLetter(COLS.get(c)) + ROWS.get("Constants");
	}
	
	/*
	public static Cell makeDayTypeCell(Row r) { 
		return CellBuilder.makeCell(r, DAY_TYPE, DayType.WORKOUT.toString());
	}
	
	
	public static Cell makeCalorieSplitCell(Row r) {
		return CellBuilder.makeCell(r, CALORIE_SPLIT, CalorieSplit.WEIGHT_LOSS.getShorthand());
	}
	
	public static Cell makeCarbFatSplitCell(Row r) {
		return CellBuilder.makeCell(r, CARB_FAT_SPLIT, CarbFatSplit.DEFAULT.getSplitName());
	}
	
	public static Cell makeUnitSystemCell(Row r) {
		return CellBuilder.makeCell(r, UNIT_SYSTEM, UnitSystem.IMPERIAL.getUnitSystemName());
	}
	
	public static Cell makeActivityMultiplierCell(Row r) {
		return CellBuilder.makeCell(r, ACTIVITY_MULTIPLIER, ActivityMultiplier.SEDENTARY.getActivityMultiplierName());
	}
	
	public static Cell makeFormulaCell(Row r) {
		return CellBuilder.makeCell(r,  FORMULA, Formula.MIFFLIN_ST_JOER.getFormulaName());
	}
	
	public static Cell makeGenderCell(Row r) {
		return CellBuilder.makeCell(r, SEX, Sex.MALE.getSexName());
	}
	
	
	public static String dateColumn() {
		return "$" + DATE_LETTER;
	}
	
	public static String dayTypeColumn() {
		return "$" + DAY_TYPE_LETTER;
	}
	
	public static String calorieSplitColumn() {
		return "$" + CALORIE_SPLIT_LETTER;
	}
	
	public static String carbFatSplitColumn() {
		return "$" + CARB_FAT_SPLIT_LETTER;
	}
	
	public static String heightColumn() {
		return "$" + HEIGHT_LETTER;
	}
	
	public static String ageColumn() {
		return "$" + AGE_LETTER;
	}
	
	public static String formulaColumn() {
		return "$" + FORMULA_LETTER;
	}
	
	public static String sexColumn() {
		return "$" + SEX_LETTER;
	}
	
	public static String unitSystemColumn() {
		return "$" + UNIT_SYSTEM_LETTER;
	}
	
	public static String activityMultiplierColumn() {
		return "$" + ACTIVITY_MULTIPLIER_LETTER;
	}
	*/
}
