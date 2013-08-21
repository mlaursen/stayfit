package user;

import enums.ActivityMultiplier;
import enums.CalorieSplit;
import enums.CarbFatSplit;
import enums.DayType;
import enums.Formula;
import enums.Sex;
import enums.UnitSystem;

public class User {

	private UnitSystem us;
	private DayType dt;
	private CarbFatSplit cfs;
	private ActivityMultiplier am;
	private CalorieSplit cs;
	private double w, h;
	private int a;
	private Formula f;
	private Sex s;
	public User(double w, double h, int a, UnitSystem us, DayType dt, CarbFatSplit cfs, ActivityMultiplier am, CalorieSplit cs, Formula f, Sex s) {
		this.w = w;
		this.h = h;
		this.a = a;
		this.us = us;
		this.dt = dt;
		this.cfs = cfs;
		this.am = am;
		this.cs = cs;
		this.f = f;
		this.s = s;
	}
	
	public double getExpectedCalories() {
		Formula f = Formula.MIFFLIN_ST_JOER;
		if(!us.equals(f.getUnitSystem())) {
			w = w / 2.2;
			h = h * 2.54;
		}
		
		double cals = this.getMultiplierActivity() * Formula.MIFFLIN_ST_JOER.getFormula(Sex.MALE, w, h, a) * this.cs.getCalorieMultiplier(this.dt);
		return cals;
	}
	
	public UnitSystem getUnitSystem() {
		return this.us;
	}
	
	public DayType getDayType() {
		return this.dt;
	}
	
	public CarbFatSplit getCarbFatSplit() {
		return this.cfs;
	}
	
	public ActivityMultiplier getActivityMultiplier() {
		return this.am;
	}
	
	public CalorieSplit getCalorieSplit() {
		return this.cs;
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
	
	public double getMultiplierCalorie() {
		return this.cs.getCalorieMultiplier(this.dt);
	}
	
	public double getMultiplierCarb() {
		return this.cfs.getCarbMultiplier(this.dt);
	}
	
	public double getMultiplierFat() {
		return this.cfs.getFatMultiplier(this.dt);
	}
}
