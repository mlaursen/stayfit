package enums;

import java.util.NoSuchElementException;

import org.apache.poi.ss.usermodel.RichTextString;

public enum Formula {
	MIFFLIN_ST_JOER("Mifflin-St Joer", UnitSystem.METRIC),
	HARRIS_BENEDICT("Harris-Benedict", UnitSystem.IMPERIAL),
	SIMPLE_MULTIPLIER("Simple Multiplier", UnitSystem.IMPERIAL);
	
	public static final String[] FORMULA = new String[Formula.values().length];
	static {
		for(int i = 0; i < Formula.values().length; i++)
			FORMULA[i] = Formula.values()[i].name;
	}
	
	private String name;
	private UnitSystem us;	
	private Formula(String name, UnitSystem us) {
		this.name = name;
		this.us = us;
	}
	
	public String getFormulaName() {
		return this.name;
	}
	
	public UnitSystem getUnitSystem() {
		return this.us;
	}
	public double getFormula(Sex s, double w, double h, int a) {
		if(this.equals(MIFFLIN_ST_JOER)) {
			return mifflinFormula(s, w, h, a);
		}
		else if(this.equals(HARRIS_BENEDICT)) {
			return harrisFormula(s, w, h , a);
		}
		else {
			return simpleFormula(s, w, h, a);
		}
	}
	
	private double mifflinFormula(Sex s, double w, double h, int a) {
		double weight = 10 * w;
		double age = 5 * a;
		double height, end;
		if(s.equals(Sex.MALE)) {
			height = 6.5 * h;
			end = 5;
		}
		else {
			height = 6.25 * h;
			end = -161;
		}
		return weight + height - age + end;
	}
	
	private double harrisFormula(Sex s, double w, double h, int a) {
		double add, wMult, hMult, aMult;
		if(s.equals(Sex.MALE)) {
			add = 66;
			wMult = 6.23;
			hMult = 12.7;
			aMult = 6.76;
		}
		else {
			add = 665;
			wMult = 4.35;
			hMult = 4.7;
			aMult = 4.7;
		}
		return add + (wMult * w) + (hMult * h) - (aMult * a);
	}
	
	private double simpleFormula(Sex s, double w, double h, int a) {
		
		return 0;
	}
	
	public static Formula parseRichTextString(RichTextString rts) { return parseString(rts.getString()); }
	public static Formula parseString(String s) {
		for(Formula f : Formula.values())
			if(f.name.equals(s))
				return f;
		throw new NoSuchElementException();
	}
}
