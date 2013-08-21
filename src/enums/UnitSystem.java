package enums;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.RichTextString;

public enum UnitSystem {
	IMPERIAL("Imperial"),
	METRIC("Metric");
	
	public static final String[] UNIT_SYSTEM = new String[UnitSystem.values().length];
	static {
		for(int i = 0; i < UnitSystem.values().length; i++)
			UNIT_SYSTEM[i] = UnitSystem.values()[i].name;
	}
	
	private String name;
	private UnitSystem(String name) {
		this.name = name;
	}
	
	public String toString() { return this.name; }
	public String getUnitSystemName() {
		return this.name;
	}
	
	public static UnitSystem parseRichTextString(RichTextString rts) { return parseString(rts.getString()); }
	public static UnitSystem parseString(String s) {
		for(UnitSystem units : UnitSystem.values())
			if(units.name.equals(s))
				return units;
		throw new NoSuchElementException();
	}
}
