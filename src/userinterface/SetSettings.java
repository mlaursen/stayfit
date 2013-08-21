package userinterface;

import java.util.Scanner;
import enums.Formula;
import enums.Sex;

public class SetSettings {

	public static void main(String[] args) {
		//Formula.MIFFLIN_ST_JOER.getFormula(Sex.MALE, 244, 71, 22);
		System.out.println((10*244/2.2) + (6.5 * 2.54*71) - (5*22) + 5);
		/*
		Scanner s = new Scanner(System.in);
		System.out.print("What is the units of measure (imperial or metric)? ");
		Unit mes = Unit.fromString(s.nextLine());
		System.out.print("What is your weight? ");
		double weight = s.nextDouble();
		System.out.print("What is your height? ");
		double height = s.nextDouble();
		System.out.print("What is your age? ");
		int age = s.nextInt();
		System.out.print("What is your activity level? ");
		ActivityMultiplier al = ActivityMultiplier.fromString(s.next());
		
		UserSettings us = new UserSettings(weight, height, age, mes, al);
		
		s.close();
		
		System.out.println();
		System.out.println(us);
		*/
	}

}
