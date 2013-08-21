package excel.sheet;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import user.User;

import enums.CalorieSplit;
import enums.CarbFatSplit;
import enums.DayType;
import enums.Formula;
import enums.Sex;
import enums.UnitSystem;
import enums.ActivityMultiplier;
import excel.CellBuilder;
import excel.Excel;

public class Settings {
	
	
	
	public static final int DATE = 0;
	public static final int WEIGHT = DATE + 1;
	public static final int DAY_TYPE = WEIGHT + 1;
	public static final int CALORIE_SPLIT = DAY_TYPE + 1;
	public static final int CARB_FAT_SPLIT = CALORIE_SPLIT + 1;
	public static final int HEIGHT = CARB_FAT_SPLIT + 2;
	public static final int AGE = HEIGHT + 1;
	public static final int UNIT_SYSTEM = AGE + 1;
	public static final int ACTIVITY_MULTIPLIER = UNIT_SYSTEM + 1;
	public static final int FORMULA = ACTIVITY_MULTIPLIER + 1;
	public static final int SEX = FORMULA + 1;
	
	
	public static final char DATE_LETTER = 'A';
	public static final char WEIGHT_LETTER = DATE_LETTER + WEIGHT;
	public static final char DAY_TYPE_LETTER = DATE_LETTER + DAY_TYPE;
	public static final char CALORIE_SPLIT_LETTER = DATE_LETTER + CALORIE_SPLIT;
	public static final char CARB_FAT_SPLIT_LETTER = DATE_LETTER + CARB_FAT_SPLIT;
	public static final char HEIGHT_LETTER = DATE_LETTER + HEIGHT;
	public static final char AGE_LETTER = DATE_LETTER + AGE;
	public static final char UNIT_SYSTEM_LETTER = DATE_LETTER + UNIT_SYSTEM;
	public static final char ACTIVITY_MULTIPLIER_LETTER = DATE_LETTER + ACTIVITY_MULTIPLIER;
	public static final char FORMULA_LETTER = DATE_LETTER + ACTIVITY_MULTIPLIER;
	public static final char SEX_LETTER = DATE_LETTER + FORMULA;

	public static final int DEFAULT_NUM_ROWS = 31; // it's really 31 - 1 rows.
	public static final double DEFAULT_WEIGHT = 0.0;
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
		
		CellBuilder.makeTitleCell(wb, titles, DATE, "Date: ");
		CellBuilder.makeTitleCell(wb, titles, WEIGHT, "Weight: ");
		CellBuilder.makeTitleCell(wb, titles, DAY_TYPE, "Day Type: ");
		CellBuilder.makeTitleCell(wb, titles, CALORIE_SPLIT, "Rest / Workout Calorie Split: ");
		CellBuilder.makeTitleCell(wb, titles, CARB_FAT_SPLIT, "Carbs / Fat Split (Rest - Workout): ");
		CellBuilder.makeTitleCell(wb, titles, HEIGHT, "Height: ");
		CellBuilder.makeTitleCell(wb, titles, AGE, "Age: ");
		CellBuilder.makeTitleCell(wb, titles, UNIT_SYSTEM, "Units: ");
		CellBuilder.makeTitleCell(wb, titles, ACTIVITY_MULTIPLIER, "Activity Multiplier: ");
		CellBuilder.makeTitleCell(wb, titles, FORMULA, "Forumla: ");
		CellBuilder.makeTitleCell(wb, titles, SEX, "Sex: ");
		
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
		
		Excel.autosizeCols(s);
		return s;
	}
	
	public static User extractSettings(Workbook wb, int row) {
		Sheet s = wb.getSheet("Settings");
		
		Row constantsRow = s.getRow(Excel.DATA_START);
		row = row / 6 + 1;
		
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
		return new User(weight, height, (int) age, units, type, carbFatSplit, activity, calSplit, f, sex);
	}
	
	
	/**
	 * Creates a Cell in the weight column with the default weight.
	 * 
	 * @param wb	The workbook being used.
	 * @param r		The row to add the weight cell to.
	 * @return		A Cell formatted to be the weight of the user with the default decimal precision.
	 */
	public static Cell makeWeightCell(Workbook wb, Row r) {
		return CellBuilder.makeNumberCell(wb, r, WEIGHT, DEFAULT_WEIGHT);
	}
	

	/**
	 * Creates a Cell in the weight column with the specified weight.
	 * 
	 * @param wb		The workbook being used.
	 * @param r			The row to add the weight cell to.
	 * @param weight	The weight to be the cell value.
	 * @return			A Cell formatted to be the weight of the user with the default decimal precision.
	 */
	public static Cell makeWeightCell(Workbook wb, Row r, double weight) {
		return CellBuilder.makeNumberCell(wb, r, WEIGHT, weight);
	}
	
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
	

	public static Cell makeDateCell(Workbook wb, Row r, int index, Date d) {
		Cell c = CellBuilder.makeCell(r, index, d);
		c.setCellStyle(Excel.dateStyle(wb));
		return c;
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
}
