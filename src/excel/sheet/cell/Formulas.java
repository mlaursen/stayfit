package excel.sheet.cell;

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
		return "IF" + paren(b + "," + t + "," + e);
	}
	
	/**
	 * Creates an excel text format formula
	 * @param t	Text
	 * @param f	Format (without "")
	 * @return	TEXT(t, "f")
	 */
	private static String textFormula(String t, String f) {
		return "TEXT" + paren(t + "," + "\"" + f + "\"");
	}
	
	/**
	 * Creates a parenthesis for formulas.
	 * @param s	The string to go inside the parenthesis
	 * @return
	 */
	private static String paren(String s) {
		return "(" + s + ")";
	}
	
	/**
	 * Creates an excel average formula.
	 * example: args = {1,2,3,4,5} would return AVERAGE(1,2,3,4,5)
	 * @param args	The arguments to be averaged
	 * @return	A string for the excel average formula
	 */
	private static String avgFormula(String[] args) {
		StringBuilder a = new StringBuilder("");
		int l = args.length;
		for(int i = 0; i < l; i++) {
			a.append(args[i]);
			if(i+1 != l)
				a.append(",");
		}
		return "AVERAGE" + paren(a.toString());
	}
	
	public static String NA() { return "NA()"; }

	/**
	 * Creates a Estimated weekly moving average formula
	 * @param rn	The row number
	 * @param x		Either 5 or 7. Number of days for average
	 * @return		String formula
	 */
	public static String ewmaFormula(int rn, int x) {
		String w = Settings.getCol(Settings.WEIGHT);
		String b = w + (rn+1) + "=\"\"";
		double[] ms;
		if(x == 5)
			ms = new double[]{.25, .5, 1, 1.25, 2};
		else
			ms = new double[]{.25, .5, .75, 1, 1.25, 1.5, 1.75};
		String[] as = new String[x];
		for(int i = 0, j = x-1; i < x; i++, j--) {
			as[i] = ms[i] + "*" + w + (rn-j);
		}
		String e = avgFormula(as);
		return ifFormula(b, NA(), e);
	}
	
	/**
	 * Creates the formula to convert a date into a 3 letter Day of Week string
	 * @param rn	The row number
	 * @return		String of the formula
	 */
	public static String dowFormula(int rn) {
		return textFormula(Settings.getCol(Settings.DATE) + "" + (rn+1), "ddd");
	}
	
	public static String smoothed(int rn) {
		String a = Settings.getConst(Settings.ALPHA);
		String w = Settings.getCol(Settings.WEIGHT) + (rn+1);
		String f = Settings.getCol(Settings.FORECAST) + (rn+1);
		String b = w + "=\"\"";
		String e = paren(a + "*" + w) + "+" + paren(paren("1-" + a) + "*"+f);
		return ifFormula(b, NA(), e);
	}
}
