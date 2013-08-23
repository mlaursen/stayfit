package excel.sheet.cell;

import excel.CellBuilder;
import excel.sheet.Settings;

public class Formulas {

	public static String ewmaFormula(int rn, int x) {
		return "IF(" + Settings.getCol("Weight") + (rn+1) + "=\"\",NA()," + (x == 5 ? ewma5Formula(rn) : ewma7Formula(rn)) + ")";
	}
	
	private static String ewma7Formula(int rn) {
		String w = Settings.getCol("Weight");
		String avg = "AVERAGE(0.25*" + w + (rn-6) + ",";
		avg += "0.5*" + w + (rn-5) + ",";
		avg += "0.75*" + w + (rn-4) + ",";
		avg += w + (rn-3) + ",";
		avg += "1.25*" + w + (rn-2) + ",";
		avg += "1.5*" + w + (rn-2) + ",";
		avg += "1.75*" + w + rn + ")";
		return avg;
	}
	
	private static String ewma5Formula(int rn) {
		String w = Settings.getCol("Weight");
		String avg = "AVERAGE(0.25*" + w + (rn-4) + ",";
		avg += "0.5*" + w + (rn-3) + ",";
		avg += w + (rn-2) + ",";
		avg += "1.25*" + w + (rn-1) + ",";
		avg += "2*" + w + rn + ")";
		return avg;
	}
	
	public static String dowFormula(int rn) {
		return "TEXT(" + Settings.getCol("Date") + (rn+1) + ", \"ddd\")";
	}
}
