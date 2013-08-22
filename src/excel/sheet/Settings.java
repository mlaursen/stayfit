package excel.sheet;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
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
		CellBuilder.boldBottomBorderCell(createCell(titles, "DOW"));
		CellBuilder.boldBottomBorderCell(createCell(titles, "Date"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "Weight"));
		CellBuilder.boldBottomBorderCell(createCell(titles, "EWMA5"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "EWMA7"));
		CellBuilder.boldBottomBorderCell(createCell(titles, "Smoothed"));
		CellBuilder.boldBottomBorderCell(createCell(titles, "Forecast"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "Residual"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "Lost/wk"));
		CellBuilder.boldBottomBorderCell(createCell(titles, "Trend"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "Slope"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "C/F Split"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "Change"));
		CellBuilder.boldBottomBorderCell(createCell(titles, "BMR"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "TDEE"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "BMR?"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "Calories"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "P*"));
		CellBuilder.boldBottomRightBorderCell(createCell(titles, "Protein"));
		CellBuilder.boldAllBorderCell(createCell(titles, "Assumed Constants"));
		CellBuilder.boldAllBorderCell(createCell(titles, "Activity Names"));
		
		LocalDate d = new LocalDate();
		for(int i = Excel.DATA_START; i <= DEFAULT_NUM_ROWS; i++) {
			Row r = s.createRow(i);
			makeDOWCell(r, d);
			makeDateCell(r, d);
			
			d = d.plusDays(1);
			makeWeightCell(r);
			
			
			makeEWMACell(r, 5);
			makeEWMACell(r, 7);
			makeSmoothedCell(r);
			/*
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
	
	
	public Cell createCell(Row r, String n) {
		return CellBuilder.makeCell(r, COLS.get(n), n);
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
		Cell c = CellBuilder.makeCell(r, COLS.get("Weight"), "");
		c.setCellStyle(CellStyles.numberStyle(wb, false, 1));
		return c;
	}

	
	public String ewma5Formula(int rn) {
		String avg = "AVERAGE(0.25*" + weightCol() + (rn-4) + ",";
		avg += "0.5*" + weightCol() + (rn-3) + ",";
		avg += weightCol() + (rn-2) + ",";
		avg += "1.25*" + weightCol() + (rn-1) + ",";
		avg += "2*" + weightCol() + rn + ")";
		return avg;
	}
	
	public String ewma7Formula(int rn) {
		String avg = "AVERAGE(0.25*" + weightCol() + (rn-6) + ",";
		avg += "0.5*" + weightCol() + (rn-5) + ",";
		avg += "0.75*" + weightCol() + (rn-4) + ",";
		avg += weightCol() + (rn-3) + ",";
		avg += "1.25*" + weightCol() + (rn-2) + ",";
		avg += "1.5*" + weightCol() + (rn-2) + ",";
		avg += "1.75*" + weightCol() + rn + ")";
		return avg;
	}
	
	public Cell makeEWMACell(Row r, int n) {
		Cell c = CellBuilder.makeCell(r, COLS.get("EWMA" + n), "");
		int rn = r.getRowNum();
		List<String> borders = new ArrayList<String>();
		if(n == 5 || (n == 7 && rn > 5)) {
			borders.add("l");
		}
		
		if(n == 7) {
			borders.add("r");
		}
		
		if(n == rn) {
			borders.add("b");
		}
		String avg = n == 5 ? ewma5Formula(rn) : ewma7Formula(rn);
		if(rn <= n) { 
			c.setCellStyle(CellStyles.grayFillBorderStyle(wb, borders));
		}
		else {
			String f = "IF(" + weightCol() + (r.getRowNum()+1) + "=\"\",NA()," + avg + ")";
			c.setCellFormula(f);
			c.setCellStyle(CellStyles.numberStyle(wb, true, 4));
		}
		
		return c;
	}
	
	public Cell makeSmoothedCell(Row r) {
		Cell c = CellBuilder.makeCell(r, COLS.get("Smoothed"), "");
		int rn = r.getRowNum();
		List<String> borders = new ArrayList<String>();
		if(rn == 5)
			borders.add("b");
		
		if(rn <= 5) { 
			c.setCellStyle(CellStyles.grayFillBorderStyle(wb, borders));
		}
		else {
			String f = "(" + alphaCol() + ROWS.get("Constants") + "*" + weightCol() + rn + ")";
			f += "+((1-" + alphaCol() + ROWS.get("Constants") + ")*" + forecastCol() + rn + ")";
			String f2 = "IF(" + weightCol() + (r.getRowNum()+1) + "=\"\",NA()," + f + ")";
			c.setCellFormula(f2);
			c.setCellStyle(CellStyles.numberStyle(wb, true, 4));
		}
		
		return c;
	}
	
	public static String dowCol() {
		return "$" + Excel.rowToLetter(COLS.get("DOW"));
	}
	
	public static String dateCol() {
		return "$" + Excel.rowToLetter(COLS.get("Date"));
	}
	
	public static String weightCol() {
		return "$" + Excel.rowToLetter(COLS.get("Weight"));
	}
	
	public static String ewma5Col() {
		return "$" + Excel.rowToLetter(COLS.get("EWMA5"));
	}
	
	public static String ewma7Col() {
		return "$" + Excel.rowToLetter(COLS.get("EWMA7"));
	}
	
	public static String smoothedCol() {
		return "$" + Excel.rowToLetter(COLS.get("Smoothed"));
	}
	
	public static String forecastCol() {
		return "$" + Excel.rowToLetter(COLS.get("Forecast"));
	}
	
	public static String alphaCol() {
		return "$" + Excel.rowToLetter(COLS.get("Alpha"));
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
