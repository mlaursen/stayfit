package excel;


import java.util.Arrays;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class CellBuilder {
	
	/**
	 * Creates a cell.
	 * @param r		Row to create cell on
	 * @param index	Cell index (column letter as number)
	 * @param o		Object to put into the cell (toString called)
	 * @return
	 */
	public static Cell createCell(Row r, int index, Object o) {
		Cell c = r.createCell(index);
		if(Integer.class.isAssignableFrom(o.getClass()))
			c.setCellValue((Integer) o);
		else if(Double.class.isAssignableFrom(o.getClass()))
			c.setCellValue((Double) o);
		else 
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
	public static Cell createFormulaCell(Row r, int i, String f) {
		Cell c = r.createCell(i);
		c.setCellFormula(f);
		return c;
	}
	
	/**
	 * Creates a formula cell for copying the contents of the previous verticle cell
	 * @param r	Row
	 * @param i	Index
	 * @return	Cell
	 */
	public static Cell createPrevCell(Row r, int i) {
		return createFormulaCell(r, i, "$" + Excel.rowToLetter(i) + r.getRowNum());
	}

	/**
	 * Creates a cell with its value equal to the cell above plus one
	 * @param r		The row to create the cell
	 * @param index	The column index to create the cell.
	 * @return		A formula cell in the form of "=$CELL1 + 1"
	 */
	public static Cell createPrevPlus1Cell(Row r, int index) {
		return createPrevPlusXCell(r, index, 1);
	}
	
	public static Cell createPrevPlusXCell(Row r, int index, int amt) {
		String f = "$" + Excel.rowToLetter(index) + r.getRowNum() + "+" + amt;
		return createFormulaCell(r, index, f);
	}
	
}
