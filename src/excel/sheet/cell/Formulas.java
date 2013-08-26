package excel.sheet.cell;

import excel.sheet.Settings;

/**
 * Class for generating formulas to put in formula cells
 * 
 * @author Mikkel Laursen
 *
 */
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
	
	private static String ifFormula(String b, String e) {
		return ifFormula(b, NA(), e);
	}
	
	private static String isEmpty() { return "=\"\""; }
	
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
	private static String avg(String[] args) {
		StringBuilder a = new StringBuilder("");
		int l = args.length;
		for(int i = 0; i < l; i++) {
			a.append(args[i]);
			if(i+1 != l)
				a.append(",");
		}
		return "AVERAGE" + paren(a.toString());
	}
	
	private static String avg(String c1, String c2) {
		return "AVERAGE" + paren(c1 + ":" + c2);
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
		String b = w + (rn+1) + isEmpty();
		double[] ms;
		if(x == 5)
			ms = new double[]{.25, .5, 1, 1.25, 2};
		else
			ms = new double[]{.25, .5, .75, 1, 1.25, 1.5, 1.75};
		String[] as = new String[x];
		for(int i = 0, j = x-1; i < x; i++, j--) {
			as[i] = ms[i] + "*" + w + (rn-j);
		}
		String e = avg(as);
		return ifFormula(b, e);
	}
	
	/**
	 * Creates the formula to convert a date into a 3 letter Day of Week string
	 * @param rn	The row number
	 * @return		String of the formula
	 */
	public static String dowFormula(int rn) {
		return textFormula(Settings.getCol(Settings.DATE) + (rn+1), "ddd");
	}
	
	/**
	 * Creates the formula for a "smoothed" estimation of the weight.
	 * @param rn	Row number
	 * @return	Formula as a string
	 */
	public static String smoothed(int rn) {
		String a = Settings.getConst(Settings.ALPHA);
		String w = Settings.getCol(Settings.WEIGHT) + (rn+1);
		String f = Settings.getCol(Settings.FORECAST) + (rn+1);
		String b = w + isEmpty();
		String e = paren(a + "*" + w) + "+" + paren(paren("1-" + a) + "*"+f);
		return ifFormula(b, e);
	}
	
	/**
	 * Creates the formlua for a forecasted weight
	 * @param rn	Row number
	 * @return		Formula as a string.
	 */
	public static String forecast(int rn) {
		String w = Settings.getCol(Settings.WEIGHT);
		String b = w + (rn+1) + isEmpty();
		String e = rn == 6 ? avg(w+2, w+7) : Settings.getCol(Settings.SMOOTHED) + rn;
		return ifFormula(b, e);
	}
	
	/**
	 * Creates a residual formula
	 * @param rn	Row number
	 * @return		Formula as a string.
	 */
	public static String residual(int rn) {
		String w = Settings.getCol(Settings.WEIGHT) + (rn+1);
		String b = w + isEmpty();
		String e = w + "-" + Settings.getCol(Settings.FORECAST) + (rn+1);
		return ifFormula(b, e);
	}
	
	/**
	 * Creates a lost/wk formula
	 * @param rn	Row number
	 * @return		Formula as a string.
	 */
	public static String lostPerWeek(int rn) {
		String w = Settings.getCol(Settings.WEIGHT);
		String b = w + (rn+1) + isEmpty();
		String e = w + (rn+8) + "-" + w + (rn+1);
		return ifFormula(b, e);
	}
	
	/**
	 * Creates the trend formula
	 * @param rn Row number
	 * @return	 Formula as a string
	 */
	public static String trend(int rn) {
		if(rn == 1)
			return Settings.getCol(Settings.WEIGHT) + "$2";
		else {
			String w = Settings.getCol(Settings.WEIGHT) + (rn+1);
			String t = Settings.getCol(Settings.TREND) + rn;
			String b = w + isEmpty();
			String e = t + "+" + paren(paren(w + "-" + t) + "/10");
			return ifFormula(b, e);
		}
	}
	
	public static String slope(int rn) {
		String w = Settings.getCol(Settings.WEIGHT) + (rn+1);
		String t = Settings.getCol(Settings.TREND);
		String b = w + isEmpty();
		String e = t + (rn+1) + "-" + t + rn;
		return ifFormula(b, e);
	}
}
