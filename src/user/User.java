package user;

import enums.ActivityMultiplier;
import enums.CarbFatSplit;
import enums.Formula;
import enums.Sex;
import enums.UnitSystem;

public class User {

	private UnitSystem us;
	private CarbFatSplit cfs;
	private ActivityMultiplier am;
	private double w, h;
	private int a;
	private Formula f;
	private Sex s;
	public User(double w, double h, int a, UnitSystem us, CarbFatSplit cfs, ActivityMultiplier am, Formula f, Sex s) {
		this.w = w;
		this.h = h;
		this.a = a;
		this.us = us;
		this.cfs = cfs;
		this.am = am;
		this.f = f;
		this.s = s;
	}
	
	public double getExpectedCalories() {
		Formula f = Formula.MIFFLIN_ST_JOER;
		if(!us.equals(f.getUnitSystem())) {
			w = w / 2.2;
			h = h * 2.54;
		}
		
		double cals = this.getMultiplierActivity() * Formula.MIFFLIN_ST_JOER.getFormula(Sex.MALE, w, h, a);
		return cals;
	}
	
	public UnitSystem getUnitSystem() {
		return this.us;
	}
	
	public CarbFatSplit getCarbFatSplit() {
		return this.cfs;
	}
	
	public ActivityMultiplier getActivityMultiplier() {
		return this.am;
	}
	
	public double getWeight() {
		return this.w;
	}
	
	public double getHeight() {
		return this.h;
	}
	
	public int getAge() {
		return this.a;
	}
	
	public Formula getFormula() {
		return this.f;
	}
	
	public Sex getSex() {
		return this.s;
	}
	
	public double getMultiplierActivity() {
		return this.am.getActivityMultiplier();
	}
	
}
