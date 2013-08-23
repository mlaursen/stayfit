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
	
	public static final short DOW, DATE, EWMA5, EWMA7, WEIGHT, SMOOTHED, FORECAST, RESIDUAL, LOST_WEEK;
	public static final short TREND, SLOPE, CFSPLIT, CHANGE, BMR, TDEE, BMR_, CALORIES, PMULT, PROTEIN;
	
	public static final short CONSTANTS, BIRTHDAY, HEIGHT, ACTIVITY, UNITS, GENDER, ALPHA;
	public static final short ACTIVITIES, SPLITS, FAT, CARBS, ACTIVITY_VAL;
	
	public static final short CONSTANTS_ROW_NAMES = 1;
	public static final short CONSTANTS_ROW = 2;
	public static final short SEDENTARY = 2;
	public static final short LIGHTLY_ACTIVE = 3;
	public static final short MODERATELY_ACTIVE = 4;
	public static final short VERY_ACTIVE = 5;
	public static final short EXTREMELY_ACTIVE = 6;
	
	public static final List<Short> TITLES_RIGHT_BORDER = new ArrayList<Short>();
	public static final List<Short> TITLES = new ArrayList<Short>();
	static {
		short i = 0;
		DOW = i++;
		DATE = i++;
		WEIGHT = i++;
		EWMA5 = i++;
		EWMA7 = i++;
		SMOOTHED = i++;
		FORECAST = i++;
		RESIDUAL = i++;
		LOST_WEEK = i++;
		TREND = i++;
		SLOPE = i++;
		CFSPLIT = i++;
		CHANGE = i++;
		BMR = i++;
		TDEE = i++;
		BMR_ = i++;
		CALORIES = i++;
		PMULT = i++;
		PROTEIN = i++;
		i++;
		CONSTANTS = i;
		BIRTHDAY = i++;
		HEIGHT = i++;
		ACTIVITY = i++;
		UNITS = i++;
		GENDER = i++;
		ALPHA = i++;
		i++;
		ACTIVITIES = i;
		SPLITS = i++;
		ACTIVITY_VAL = i;
		CARBS = i++;
		FAT = i++;
		
		TITLES_RIGHT_BORDER.add(WEIGHT, WEIGHT);
		TITLES_RIGHT_BORDER.add(EWMA7, EWMA7);
		TITLES_RIGHT_BORDER.add(RESIDUAL, RESIDUAL);
		TITLES_RIGHT_BORDER.add(LOST_WEEK, LOST_WEEK);
		TITLES_RIGHT_BORDER.add(SLOPE, SLOPE);
		TITLES_RIGHT_BORDER.add(CFSPLIT, CFSPLIT);
		TITLES_RIGHT_BORDER.add(CHANGE, CHANGE);
		TITLES_RIGHT_BORDER.add(TDEE, TDEE);
		TITLES_RIGHT_BORDER.add(BMR_, BMR_);
		TITLES_RIGHT_BORDER.add(CALORIES, CALORIES);
		TITLES_RIGHT_BORDER.add(PMULT, PMULT);
		TITLES_RIGHT_BORDER.add(PROTEIN, PROTEIN);
		
		TITLES.addAll(TITLES_RIGHT_BORDER);
		TITLES.add(DOW, DOW);
		TITLES.add(DATE, DATE);
		TITLES.add(EWMA5, EWMA5);
		TITLES.add(SMOOTHED, SMOOTHED);
		TITLES.add(FORECAST, FORECAST);
		TITLES.add(BMR, BMR);
		TITLES.add(TREND, TREND);
		TITLES.add(CONSTANTS, CONSTANTS);
		TITLES.add(ACTIVITIES, ACTIVITIES);
		TITLES.add(SPLITS, SPLITS);
		TITLES.add(CARBS, CARBS);
		TITLES.add(FAT, FAT);
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
		createCell(titles, DOW, "DOW");
		createCell(titles, DATE, "Date");
		createCell(titles, WEIGHT, "Weight");
		createCell(titles, EWMA5, "EWMA5");
		createCell(titles, EWMA7, "EWMA7");
		createCell(titles, SMOOTHED, "Smoothed");
		createCell(titles, FORECAST, "Forecast");
		createCell(titles, RESIDUAL, "Residual");
		createCell(titles, LOST_WEEK, "Lost/wk");
		createCell(titles, TREND, "Trend");
		createCell(titles, SLOPE, "Slope");
		createCell(titles, CFSPLIT, "C/F Split");
		createCell(titles, CHANGE, "Change");
		createCell(titles, BMR, "BMR");
		createCell(titles, TDEE, "TDEE");
		createCell(titles, BMR_, "BMR?");
		createCell(titles, CALORIES, "Calories");
		createCell(titles, PMULT, "P*");
		createCell(titles, PROTEIN, "Protein");
		createCell(titles, CONSTANTS, "Assumed Constants");
		createCell(titles, ACTIVITIES, "Activity Names");
		
		
		LocalDate d = new LocalDate();
		for(int i = Excel.DATA_START; i <= DEFAULT_NUM_ROWS; i++) {
			Row r = s.createRow(i);
			createFormulaCell(r, DOW, Formulas.dowFormula(r.getRowNum()));
			createCell(r, DATE, d);
			d = d.plusDays(1);
			createCell(r, WEIGHT, "");
			createCell(r, EWMA5, "");
			createCell(r, EWMA7, "");
			createCell(r, SMOOTHED, "");
			createCell(r, FORECAST, "");
			createCell(r, RESIDUAL, "");
			createCell(r, LOST_WEEK, "");
			createCell(r, TREND, "");
			createCell(r, SLOPE, "");
			createCell(r, CFSPLIT, "");
			createCell(r, CHANGE, "");
			createCell(r, BMR, "");
			createCell(r, TDEE, "");
			createCell(r, BMR_, "");
			createCell(r, CALORIES, "");
			createCell(r, PMULT, "");
			createCell(r, PROTEIN, "");
		
		}
		
		Row cNames = s.getRow(CONSTANTS_ROW_NAMES);
		Row cVals = s.getRow(CONSTANTS_ROW);
		createCell(cNames, BIRTHDAY, "Birthday");
		createCell(cNames, HEIGHT, "Height");
		createCell(cNames, ACTIVITY, "Activity Multiplier");
		createCell(cNames, UNITS, "Units");
		createCell(cNames, GENDER, "Gender");
		createCell(cNames, ALPHA, "alpha");
		createCell(cVals, BIRTHDAY, "");
		createCell(cVals, HEIGHT, "");
		createCell(cVals, ACTIVITY, "");
		createCell(cVals, UNITS, "");
		createCell(cVals, GENDER, "");
		createCell(cVals, ALPHA, "0.3");
		
		Row r = s.getRow(SEDENTARY);
		r.createCell(ACTIVITY_VAL);
		//Row cNames = s.getRow(ROWS.get("Constants") - 1);
		//Row cVals = s.getRow(ROWS.get("Constant Values") - 1);
		//createCell(cNames, "Birthday");
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
	
	public Cell createFormulaCell(Row r, short n, String f) {
		return createCell(CellBuilder.makeFormulaCell(r, n, f), n, r.getRowNum());
	}
	
	public Cell createCell(Row r, short n, Object v) {
		int rn = r.getRowNum();
		Cell c = CellBuilder.makeCell(r, n, v);
		if(n == DATE && rn >= Excel.DATA_START) {
			c = CellBuilder.makePrevPlus1Cell(r, n);
			if(rn == Excel.DATA_START) {
				c = r.createCell(n); 
				c.setCellValue(((LocalDate) v).toDate());
			}
		}
		//else if(n.contains("EWMA")) {
		if(n == EWMA5 || n == EWMA7) {
			int x = n == EWMA5 ? 5 : 7;
			if(rn > x)
				c = CellBuilder.makeFormulaCell(r, n, Formulas.ewmaFormula(rn, x));
		}
		else if(rn > 5) {	// grayed otherwise with empty string as value
			//if(n == "Smoothed") {
			if(n == SMOOTHED) {
				c = CellBuilder.makeFormulaCell(r, n, Formulas.smoothed(rn));
			}
		}
		return createCell(c, n, rn);
	}
	public Cell createCell(Cell c, short n, int rn) {
		Set<Short> styles = addStyles(n, rn);
		CellStyles.applyStyles(styles, c);
		return c;
	}
	
	public Set<Short> addStyles(short n, int rn) {
		Set<Short> styles = new HashSet<Short>();
		if(TITLES.contains(n) && rn == Excel.DATA_START - 1) {
			styles.add(CellStyles.BOLD);
			styles.add(CellStyles.BORDER_BOTTOM);
		}
		
		if(TITLES_RIGHT_BORDER.contains(n)) {
			styles.add(CellStyles.BORDER_RIGHT);
		}
		
		if(rn % 7 == 0)
			styles.add(CellStyles.BORDER_BOTTOM);
		
		if(n == EWMA5 || n == EWMA7) {
			int e = n == EWMA5 ? 5 : 7;
			if(rn >= Excel.DATA_START && (rn <= 5 || (e == 7 && rn <= e))) 
				styles.add(CellStyles.GRAY_FILL);
			if(e == 5 && rn > 5 && rn < 8)
				styles.add(CellStyles.BORDER_RIGHT_THIN);
			
			if(e == 7)
				styles.add(CellStyles.BORDER_RIGHT);
			
			if(rn == e && e == 5)
				styles.add(CellStyles.BORDER_BOTTOM_THIN);
		}
		else if(n == SMOOTHED || n == FORECAST || n == RESIDUAL) {
			if(rn <= 5 && rn >= Excel.DATA_START)
				styles.add(CellStyles.GRAY_FILL);
			
			if(rn == 5)
				styles.add(CellStyles.BORDER_BOTTOM_THIN);
			
			if(n == RESIDUAL)
				styles.add(CellStyles.BORDER_RIGHT);
		}
		else if(n == DATE) {
			styles.add(CellStyles.DATE);
		}
		return styles;
	}
	
	public static String getCol(short r) {
		return "$" + Excel.rowToLetter(r);
	}
	
	public static String getConst(short c) {
		return "$" + Excel.rowToLetter(c) + "$" + CONSTANTS_ROW;
	}
}
