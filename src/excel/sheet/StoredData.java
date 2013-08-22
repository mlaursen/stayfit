package excel.sheet;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import user.User;

import enums.ActivityMultiplier;
import enums.CalorieSplit;
import enums.CarbFatSplit;
import excel.CellBuilder;
import excel.Data;
import excel.Excel;

public class StoredData {

	public static final int REST_CAL = 0;
	public static final int WORKOUT_CAL = REST_CAL + 1;
	public static final int FAT_REST = WORKOUT_CAL + 1;
	public static final int CARB_REST = FAT_REST + 1;
	public static final int FAT_WORKOUT = CARB_REST + 1;
	public static final int CARB_WORKOUT = FAT_WORKOUT + 1;
	public static final int ACTIVITY_MULT = CARB_WORKOUT + 1;
	
	private Workbook wb;
	public StoredData(Workbook wb) {
		this.wb = wb;
	}
	
	public Sheet createStoredDataSheet() {
		Sheet s = wb.createSheet("StoredData");
		Row r = s.createRow(0);
		//CellBuilder.makeTitleCell(wb, r, REST_CAL, "Rest Calorie Multiplier: ");
		//CellBuilder.makeTitleCell(wb, r, WORKOUT_CAL, "Workout Calorie Multiplier: ");
		//CellBuilder.makeTitleCell(wb, r, FAT_REST, "Fat (Rest Day): ");
		//CellBuilder.makeTitleCell(wb, r, CARB_REST, "Carb (Rest Day): ");
		//CellBuilder.makeTitleCell(wb, r, FAT_WORKOUT, "Fat (Workout Day): ");
		//CellBuilder.makeTitleCell(wb, r, CARB_WORKOUT, "Carb (Workout Day): ");
		//CellBuilder.makeTitleCell(wb, r, WORKOUT_CAL, "Workout Calorie Multiplier: ");
		//CellBuilder.makeTitleCell(wb, r, ACTIVITY_MULT, "Activity Multiplier: ");
		
		int numRows = Math.max(CarbFatSplit.values().length, Math.max(ActivityMultiplier.values().length, CalorieSplit.values().length));
		for(int i = 1; i <= numRows; i++)
			s.createRow(i);
		
		for(int i = 0; i < CalorieSplit.values().length; i++) {
			//CellBuilder.makeNumberCell(wb, s.getRow(i+1), REST_CAL, CalorieSplit.values()[i].getRest());
			//CellBuilder.makeNumberCell(wb, s.getRow(i+1), WORKOUT_CAL, CalorieSplit.values()[i].getWorkout());
		}
		
		for(int i = 0; i < CarbFatSplit.values().length; i++) {
		//	CellBuilder.makeNumberCell(wb, s.getRow(i+1), FAT_REST, CarbFatSplit.values()[i].getRestFatMultiplier());
			//CellBuilder.makeNumberCell(wb, s.getRow(i+1), CARB_REST, CarbFatSplit.values()[i].getRestCarbMultiplier());
		//	CellBuilder.makeNumberCell(wb, s.getRow(i+1), FAT_WORKOUT, CarbFatSplit.values()[i].getWorkoutFatMultiplier());
		//	CellBuilder.makeNumberCell(wb, s.getRow(i+1), CARB_WORKOUT, CarbFatSplit.values()[i].getWorkoutCarbMultiplier());
		}
		
		for(int i = 0; i < ActivityMultiplier.values().length; i++) {
			//CellBuilder.makeNumberCell(wb, s.getRow(i+1), ACTIVITY_MULT, ActivityMultiplier.values()[i].getActivityMultiplier(), 3);
		}
		
		Excel.autosizeCols(s);
		return s;
	}
	
	public static Data extractData(Workbook wb) {
		// 
		User u = Settings.extractSettings(wb, 1);
		
		return new Data(0,0,0);
	}
}
