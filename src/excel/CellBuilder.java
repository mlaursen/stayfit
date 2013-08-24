package excel;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class CellBuilder {
	
	/**
	 * Creates a cell.
	 * @param r		Row to create cell on
	 * @param index	Cell index (column letter as number)
	 * @param o		Object to put into the cell (toString called)
	 * @return
	 */
	public static Cell makeCell(Row r, int index, Object o) {
		Cell c = r.createCell(index);
		c.setCellValue(o.toString());
		return c;
	}
	
	/**
	 * Create a cell that has a formula
	 * @param r	Row to create cell
	 * @param i	Index (column letter as number)
	 * @param f	String representation of the formula
	 * @return	a Cell
	 */
	public static Cell makeFormulaCell(Row r, int i, String f) {
		Cell c = r.createCell(i);
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
