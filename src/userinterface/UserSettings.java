package userinterface;

import enums.ActivityMultiplier;
import enums.UnitSystem;

public class UserSettings {

	private double weight;
	private double height;
	private int age;
	private UnitSystem unit;
	private ActivityMultiplier multiplier;
	
	public UserSettings(double weight, double height, int age, UnitSystem unit, ActivityMultiplier multiplier) {
		this.weight = weight;
		this.height = height;
		this.age = age;
		this.unit = unit;
		this.multiplier = multiplier;
	}
	
	public String toString() {
		String s = "User weight: " + this.weight + "\n"
				 + "User height: " + this.height + "\n"
				 + "User age: " + this.age + "\n"
				 + "Units of measure: " + this.unit + "\n"
				 + "Activity multiplier: " + multiplier;
		return s;
	}
	
	
}
