package enums;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.RichTextString;

public enum Macro {
	CALORIES("Calories"),
	FAT("Fat", 9),
	CARBOHYDRATES("Carbohydrates", 4),
	PROTEIN("Protein", 4);
	
	public static final String[] MACRO = new String[Macro.values().length];
	static {
		for(int i = 0; i < Macro.values().length; i++)
			MACRO[i] = Macro.values()[i].name;
	}
	
	private String name;
	private int toCal;
	private Macro(String name) {
		this.name = name;
		this.toCal = 1;
	}
	private Macro(String name, int toCal) {
		this.name = name;
		this.toCal = toCal;
	}
	
	public String toString() {
		return "Macro name: " + this.name + "\n"
			 + "To Calorie Multiplier " + this.toCal;
	}
	
	public String getMacroName() { return this.name; }
	public int getToCalorie() { return this.toCal; }
	
	public static Macro parseRichTextString(RichTextString rts) { return parseString(rts.getString()); }
	public static Macro parseString(String s) {
		for(Macro macro : Macro.values())
			if(macro.name.equals(s))
				return macro;
		throw new NoSuchElementException();
	}
}
