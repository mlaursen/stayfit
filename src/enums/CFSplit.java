package enums;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.RichTextString;

import excel.sheet.Settings;

public enum CFSplit {
	A("(0 / 100)", 0, 1),
	B("(25 / 75)", .25, .75),
	C("(75 / 25)", .75, .25),
	D("(60 / 40)", .6, .4),
	E("(40 / 60)", .4, .6),
	F("(20 / 80)", .2, .8),
	G("(80 / 20)", .8, .2),
	H("(10 / 90)", .1, .9),
	I("(90 / 10)", .9, .1),
	J("(50 / 50)", .5, .5);
	
	public static final String[] CARB_FAT_SPLIT = new String[CFSplit.values().length];
	static {
		for(int i = 0; i < CFSplit.values().length; i++)
			CARB_FAT_SPLIT[i] = CFSplit.values()[i].name;
	}
	
	private String name;
	private double c, f;
	private CFSplit(String name, double c, double f) {
		this.name = name;
		this.c = c;
		this.f = f;
	}
	
	public String toString() { 
		return "Split name: " + this.name + "\n"
			 + "Carb multiplier: " + this.c + "\n"
			 + "Fat multiplier: " + this.f;
	}
	
	public String getSplitName() { return this.name; }
	public double getFat() { return this.f; }
	public double getCarb() { return this.c; }
	public int asRowNum() {
		return this.ordinal() + Settings.SPLITS_START + 1;
	}
	public static CFSplit parseRichTextString(RichTextString rts) { return parseString(rts.getString()); }
	public static CFSplit parseString(String s) {
		for(CFSplit split : CFSplit.values())
			if(split.name.equals(s))
				return split;
		throw new NoSuchElementException();
	}
}
