package excel;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class CellBuilder {
	
	public static Cell makeCell(Row r, int index, Object o) {
		Cell c = r.createCell(index);
		c.setCellValue(o.toString());
		return c;
	}
	
	public static Cell makeFormulaCell(Row r, int index, String f) {
		Cell c = r.createCell(index);
		c.setCellFormula(f);
		return c;
	}
	

	/**
	 * Creates a cell with its value equal to the cell above plus one
	 * @param r		The row to create the cell
	 * @param index	The column index to create the cell.
	 * @return		A formula cell in the form of "=$CELL1 + 1"
	 */
	public static Cell makePrevPlus1Cell(Row r, int index) {
		return makePrevPlusXCell(r, index, 1);
	}
	
	public static Cell makePrevPlusXCell(Row r, int index, int amt) {
		String f = "$" + Excel.rowToLetter(index) + r.getRowNum() + "+" + amt;
		return makeFormulaCell(r, index, f);
	}
}
