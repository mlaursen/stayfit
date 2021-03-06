package enums;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.RichTextString;

import excel.Excel;

public enum ActivityMultiplier {
	SEDENTARY("Sedentary", 1.2),
	LIGHTLY_ACTIVE("Lightly Active", 1.375),
	MODERATELY_ACTIVE("Moderately Active", 1.55),
	VERY_ACTIVE("Very Active", 1.725),
	EXTREMELY_ACTIVE("Extremely Active", 1.9);
	
	public static final String[] ACTIVITY_MULTIPLIER = new String[ActivityMultiplier.values().length];
	static {
		for(int i = 0; i < ActivityMultiplier.values().length; i++)
			ACTIVITY_MULTIPLIER[i] = ActivityMultiplier.values()[i].name;
	}
	
	private String name;
	private double multiplier;
	private ActivityMultiplier(String name, double multiplier) {
		this.name = name;
		this.multiplier = multiplier;
	}
	
	public int asRowNum() {
		return this.ordinal() + Excel.DATA_START;
	}
	
	public String toString() { 
		return "Activity Multiplier name: " + this.name + "\n"
			 + "Activity Multiplier: " + this.multiplier;
	}
	public String getActivityMultiplierName() { return this.name; }
	public double getActivityMultiplier() { return this.multiplier; }
	
	public static ActivityMultiplier parseRichTextString(RichTextString rts) { return parseString(rts.getString()); }
	public static ActivityMultiplier parseString(String s) {
		for(ActivityMultiplier activity : ActivityMultiplier.values())
			if(activity.name.equals(s))
				return activity;
		throw new NoSuchElementException();
	}
}
