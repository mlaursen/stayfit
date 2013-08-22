package excel;

import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import enums.CalorieSplit;
import enums.CarbFatSplit;
import enums.DayType;
import enums.UnitSystem;
import excel.sheet.Settings;

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
	
	/**
	 * Creates a cell with the bold cell style
	 * @param wb	The workbook to create the bold cell style for.
	 * @param r		The row to add the cell to.
	 * @param index	The index of the cell.
	 * @param title	The value to be put into the cell.
	 * @return		A new cell that has the bold style and the title value.
	 */
	public static Cell makeTitleCell(Workbook wb, Row r, int index, String title) {
		Cell c = r.createCell(index);
		c.setCellStyle(CellStyles.boldCell(wb));
		c.setCellValue(title);
		return c;
	}
	
	public static Cell makeNumberCell(Workbook wb, Row r, int index) {
		Cell c = makeCell(r, index, "");
		c.setCellStyle(CellStyles.numberCell(wb));
		return c;
	}
	public static Cell makeNumberCell(Workbook wb, Row r, int index, double val) {
		Cell c = makeCell(r, index, val);
		c.setCellStyle(CellStyles.numberCell(wb));
		return c;
	}
	
	public static Cell makeNumberCell(Workbook wb, Row r, int index, double val, int precision) {
		Cell c = makeCell(r, index, val);
		c.setCellStyle(CellStyles.numberCell(wb, precision));
		return c;
	}
	
}
