package enums;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.RichTextString;

public enum CalorieSplit {
	STANDARD_RECOMP("Standard Recomp", "(-20 / +20)", .8, 1.2),
	WEIGHT_LOSS("Weight Loss", "(-20 / +00)", .8, 1),
	WEIGHT_LOSS2("Weight Loss #2", "(-40 / +20)", .6, 1.2),
	FASTER_WEIGHT_LOSS("Faster Weight Loss", "(-30 / -10)", .7, .9),
	LEAN_MASSING("Lean Massing", "(-10 / +20)", .9, 1.2);
	
	public static final String[] CALORIE_SPLIT = new String[CalorieSplit.values().length];
	static {
		for(int i = 0; i < CalorieSplit.values().length; i++)
			CALORIE_SPLIT[i] = CalorieSplit.values()[i].shorthand;
	}
	
	private String name, shorthand;
	private double rest, workout;
	private CalorieSplit(String name, String representation, double rest, double workout) {
		this.name = name;
		this.shorthand = representation;
		this.rest = rest;
		this.workout = workout;
	}
	
	public String toString() { 
		return "Split name: " + this.name + "\n"
			 + "Value: " + this.shorthand + "\n"
			 + "Rest multiplier: " + this.rest + "\n"
			 + "Workout multiplier: " + this.workout;
	}
	public String getSplitName() { return this.name; }
	public String getShorthand() { return this.shorthand; }
	
	public double getRest() { return this.rest; }
	public double getWorkout() { return this.workout; }
	
	public static CalorieSplit parseRichTextString(RichTextString rts) { return parseString(rts.getString()); }
	public static CalorieSplit parseString(String s) {
		for(CalorieSplit split : CalorieSplit.values())
			if(split.name.equals(s) || split.shorthand.equals(s))
				return split;
		throw new NoSuchElementException();
	}
	
	public double getCalorieMultiplier(DayType dt) {
		if(dt.equals(DayType.REST))
			return this.rest;
		else
			return this.workout;
	}
}
