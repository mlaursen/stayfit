package enums;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.RichTextString;

public enum Gender {
	MALE("Male"),
	FEMALE("Female");
	
	public static final String[] GENDER = new String[Gender.values().length];
	static {
		for(int i = 0; i < Gender.values().length; i++)
			GENDER[i] = Gender.values()[i].name;
	}
	
	private String name;
	private Gender(String name) {
		this.name = name;
	}
	
	public String getGenderName() {
		return this.name;
	}
	

	public static Gender parseRichTextString(RichTextString rts) { return parseString(rts.getString()); }
	public static Gender parseString(String s) {
		for(Gender g : Gender.values())
			if(g.name.equals(s))
				return g;
		throw new NoSuchElementException();
	}
}
