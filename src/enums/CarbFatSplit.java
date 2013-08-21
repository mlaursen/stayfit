package enums;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.RichTextString;

public enum CarbFatSplit {
	DEFAULT("(50/50) - (50/50)", .5, .5, .5, .5),
	REST_LESS_FAT("(25/75) - (75/25)", .25, .75, .75, .25),
	REST_LESS_FAT2("(20/80) - (80/20)", .2, .8, .8, .2),
	REST_LESS_FAT3("(15/85) - (85/15)", .15, .85, .85, .15),
	REST_LESS_FAT4("(10/90) - (90/10)", .1, .9, .9, .1),
	WORKOUT_MORE_FAT("(50/50) - (75-25)", .5, .5, .75, .25);
	
	public static final String[] CARB_FAT_SPLIT = new String[CarbFatSplit.values().length];
	static {
		for(int i = 0; i < CarbFatSplit.values().length; i++)
			CARB_FAT_SPLIT[i] = CarbFatSplit.values()[i].name;
	}
	
	private String name;
	private double restFat, restCarb, workoutFat, workoutCarb;
	private CarbFatSplit(String name, double restFat, double restCarb, double workoutFat, double workoutCarb) {
		this.name = name;
		this.restFat = restFat;
		this.restCarb = restCarb;
		this.workoutFat = workoutFat;
		this.workoutCarb = workoutCarb;
	}
	
	public String toString() { 
		return "Split name: " + this.name + "\n"
			 + "Rest Fat Multiplier: " + this.restFat + "\n"
			 + "Rest Carb Multiplier: " + this.restCarb + "\n"
			 + "Workout Fat Multiplier: " + this.workoutFat + "\n"
			 + "Workout Carb Multiplier: " + this.workoutCarb;
	}
	
	public String getSplitName() { return this.name; }
	public double getRestFatMultiplier() { return this.restFat; }
	public double getRestCarbMultiplier() { return this.restCarb; }
	public double getWorkoutFatMultiplier() { return this.workoutFat; }
	public double getWorkoutCarbMultiplier() { return this.workoutCarb; }
	
	public double getCarbMultiplier(DayType dt) {
		if(dt.equals(DayType.REST))
			return this.restCarb;
		else
			return this.workoutCarb;
	}
	
	public double getFatMultiplier(DayType dt) {
		if(dt.equals(DayType.REST))
			return this.restFat;
		else
			return this.workoutFat;
	}
	
	public static CarbFatSplit parseRichTextString(RichTextString rts) { return parseString(rts.getString()); }
	public static CarbFatSplit parseString(String s) {
		for(CarbFatSplit split : CarbFatSplit.values())
			if(split.name.equals(s))
				return split;
		throw new NoSuchElementException();
	}
}
