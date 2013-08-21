package enums;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.RichTextString;

public enum DayType {
	WORKOUT("Workout"),
	REST("Rest");
	
	public static final String[] DAY_TYPES = new String[DayType.values().length];
	static {
		for(int i = 0; i < DayType.values().length; i++)
			DAY_TYPES[i] = DayType.values()[i].name;
	}
	
	private String name;
	private DayType(String name) {
		this.name = name;
	}
	
	public String getTypeName() {
		return this.name;
	}
	public String toString() { return this.name; }
	public static DayType parseRichTextString(RichTextString rts) { return parseString(rts.getString()); }
	public static DayType parseString(String s) {
		for(DayType type : DayType.values())
			if(type.name.equals(s))
				return type;
		throw new NoSuchElementException();
	}
	
}
