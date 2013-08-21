package enums;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.RichTextString;

public enum Sex {
	MALE("Male"),
	FEMALE("Female");
	
	public static final String[] SEX = new String[Sex.values().length];
	static {
		for(int i = 0; i < Sex.values().length; i++)
			SEX[i] = Sex.values()[i].name;
	}
	
	private String name;
	private Sex(String name) {
		this.name = name;
	}
	
	public String getSexName() {
		return this.name;
	}
	

	public static Sex parseRichTextString(RichTextString rts) { return parseString(rts.getString()); }
	public static Sex parseString(String s) {
		for(Sex sex : Sex.values())
			if(sex.name.equals(s))
				return sex;
		throw new NoSuchElementException();
	}
}
