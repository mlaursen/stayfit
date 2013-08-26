package excel;

public class Data {
	
	private double calMult, fatMult, carbMult;
	public Data(double calMult, double fatMult, double carbMult) {
		this.calMult = calMult;
		this.fatMult = fatMult;
		this.calMult = carbMult;
	}
	
	public double getCalMult() { return this.calMult; }
	public double getFatMult() { return this.fatMult; }
	public double getCarbMult() { return this.carbMult; }
}
