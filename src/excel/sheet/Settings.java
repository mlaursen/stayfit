package excel.sheet;

import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.LocalDate;

import user.User;
import enums.CalorieSplit;
import enums.CarbFatSplit;
import enums.DayType;
import enums.Formula;
import enums.Sex;
import enums.UnitSystem;
import enums.ActivityMultiplier;
import excel.CellBuilder;
import excel.CellStyles;
import excel.Excel;

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
		
		createTitleCell(titles, "DOW");
		createTitleCell(titles, "DOW");
		createTitleCell(titles, "Date");
		createTitleCell(titles, "Weight");
		createTitleCell(titles, "EWMA5");
		createTitleCell(titles, "EWMA7");
		createTitleCell(titles, "Smoothed");
		createTitleCell(titles, "Forecast");
		createTitleCell(titles, "Residual");
		createTitleCell(titles, "Lost/wk");
		createTitleCell(titles, "Trend");
		createTitleCell(titles, "Slope");
		createTitleCell(titles, "C/F Split");
		createTitleCell(titles, "Change");
		createTitleCell(titles, "BMR");
		createTitleCell(titles, "TDEE");
		createTitleCell(titles, "BMR?");
		createTitleCell(titles, "Calories");
		createTitleCell(titles, "P*");
		createTitleCell(titles, "Protein");
		createTitleCell(titles, "Assumed Constants");
		createTitleCell(titles, "Activity Names");
		
		
		
		LocalDate d = new LocalDate();
		for(int i = Excel.DATA_START; i <= DEFAULT_NUM_ROWS; i++) {
			Row r = s.createRow(i);
			makeDOWCell(r, d);
			makeDateCell(r, d);
			
			d = d.plusDays(1);
			makeWeightCell(r);
			
			/*
			makeEWMA5Cell(r);
			makeEWMA7Cell(r);
			makeSmoothedCell(r);
			makeForecastCell(r);
			makeResidualCell(r);
			makeLostWeekCell(r);
			makeTrendCell(r);
			makeSlopeCell(r);
			makeCFSplitCell(r);
			makeCalChangeCell(r);
			makeBMRCell(r);
			makeTDEECell(r);
			makeUseBMRCell(r);
			makeCaloriesCell(r);
			makePMCell(r);
			makeProteinCell(r);
			*/
		}
		
		Row r = s.getRow(ROWS.get("Split Table"));
		createTitleCell(r, "Split");
		createTitleCell(r, "Carbs");
		createTitleCell(r, "Fat");
		
		r = s.getRow(ROWS.get("Constants"));
		CellBuilder.makeCell(r, COLS.get("Birthday"), "Birthday");
		CellBuilder.makeCell(r, COLS.get("Height"), "Height");
		CellBuilder.makeCell(r, COLS.get("Activity"), "Activity Multiplier");
		CellBuilder.makeCell(r, COLS.get("Units"), "Units");
		CellBuilder.makeCell(r, COLS.get("Gender"), "Gender");
		CellBuilder.makeCell(r, COLS.get("Alpha"), "alpha");
		
		//makeDateCell(r, COLS.get("Birthday"));
		//CellBuilder.makeNumberCell(wb, r, COLS.get("Height"));
		//Excel.createDropDown(s, enums.ActivityMultiplier.ACTIVITY_MULTIPLIER, CONSTS, CONSTS, COLS.get("Activity"), COLS.get("Activity"));
		
		/*
		
		
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

	public Cell createTitleCell(Row r, String n) {
		return CellBuilder.makeTitleCell(this.wb, r, COLS.get(n), n);
	}
	
	public Cell makeDOWCell(Row r, LocalDate d) {
		String formula = "TEXT(" + dateCol() + (r.getRowNum()+1) + ", \"ddd\")";
		return CellBuilder.makeFormulaCell(r, COLS.get("DOW"), formula);
	}
	
	
	
	
	/**
	 * Creates an empty cell with a date format
	 * @param r			The row to create the cell
	 * @param index		The index of the cell
	 * @return			A cell with a date format.
	 */
	public Cell makeDateCell(Row r, int index) { 
		Cell c = r.createCell(index);
		c.setCellStyle(CellStyles.dateStyle(this.wb));
		return c;
	}
	public Cell makeDateCell(Row r, LocalDate d) { return makeDateCell(r, COLS.get("Date"), d); }
	public Cell makeDateCell(Row r, int index, LocalDate d) {
		Cell c = r.getRowNum() == Excel.DATA_START ? r.createCell(index) : CellBuilder.makePrevPlus1Cell(r, index);
		if(r.getRowNum() == Excel.DATA_START) c.setCellValue(d.toDate());
		c.setCellStyle(CellStyles.dateStyle(this.wb));
		return c;
	}
	
	public Cell makeWeightCell(Row r) {
		return CellBuilder.makeNumberCell(this.wb, r, COLS.get("Weight"));
	}
	
	public Cell makeWeightCell(Row r, double weight) {
		return CellBuilder.makeNumberCell(this.wb, r, COLS.get("Weight"), weight);
	}
	
	
	
	public static String dowCol() {
		return "$" + Excel.rowToLetter(COLS.get("DOW"));
	}
	
	public static String dateCol() {
		return "$" + Excel.rowToLetter(COLS.get("Date"));
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
