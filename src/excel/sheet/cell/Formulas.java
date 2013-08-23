package excel.sheet.cell;

import excel.CellBuilder;
import excel.sheet.Settings;

public class Formulas {
	
	/**
	 * Creates an excel IF formula
	 * @param b	Boolean expression
	 * @param t	Then case
	 * @param e	Else case
	 * @return	Complete excel IF formula
	 */
	private static String ifFormula(String b, String t, String e) {
		return "IF(" + b + "," + t + "," + e + ")";
	}
	
	private static String avgFormula(String[] args) {
		StringBuilder a = new StringBuilder("AVERAGE(");
		int l = args.length;
		for(int i = 0; i < l; i++) {
			a.append(args[i]);
			if(i+1 != l)
				a.append(",");
		}
		return a + ")";
	}
	
	public static String NA() { return "NA()"; }

	public static String ewmaFormula(int rn, int x) {
		String w = Settings.getCol("Weight");
		String b = w + (rn+1) + "=\"\"";
		double[] ms;
		if(x == 5)
			ms = new double[]{.25, .5, 1, 1.25, 2};
		else
			ms = new double[]{.25, .5, .75, 1, 1.25, 1.5, 1.75};
		String[] as = new String[x];
		for(int i = 0, j = x; i < x; i++, j--) {
			as[i] = ms[i] + "*" + w + (rn-j);
		}
		String e = avgFormula(as);
		return ifFormula(b, NA(), e);
	}
	
	public static String dowFormula(int rn) {
		return "TEXT(" + Settings.getCol("Date") + (rn+1) + ", \"ddd\")";
	}
	
	public static String smoothed(int rn) {
		String a = Settings.getConst("Alpha");
		String w = Settings.getCol("Weight");
		String f = Settings.getCol("Forecast");
		String form = "(" + a + "*" + w + rn + ")";
		form += "+((1-" + a + ")*" + f + rn + ")";
		return "IF(" + w + (rn+1) + "=\"\",NA()," + form + ")";
	}
}
