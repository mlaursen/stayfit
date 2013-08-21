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

public class CellBuilder2 {
	

	/**
	 * Creates a cell with a value.
	 * @param r			The row to create the cell in.
	 * @param index		The cell index.
	 * @param s			The String value to be placed in the cell.
	 * @return			A Cell in with a value.
	 */
	public static Cell makeCell(Row r, int index, String s) {
		Cell c = r.createCell(index);
		c.setCellValue(s);
		return c;
	}
	
	/**
	 * Creates a cell with a value.
	 * @param r			The row to create the cell in.
	 * @param index		The cell index.
	 * @param d			The double value to be placed in the cell.
	 * @return			A Cell in with a value.
	 */
	public static Cell makeCell(Row r, int index, double d) {
		Cell c = r.createCell(index);
		c.setCellValue(d);
		return c;
	}
	
	/**
	 * Creates a cell with a date value.
	 * @param r			The row to create the cell in.
	 * @param index		The cell index.
	 * @param d			The date to be placed in the cell
	 * @return			A Cell with a date value.
	 */
	public static Cell makeCell(Row r, int index, Date d) {
		Cell c = r.createCell(index);
		c.setCellValue(d);
		return c;
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
		c.setCellStyle(Excel.boldCell(wb));
		c.setCellValue(title);
		return c;
	}
	
	public static Cell makeNumberCell(Workbook wb, Row r, int index, double val) {
		Cell c = makeCell(r, index, val);
		c.setCellStyle(Excel.numberCell(wb));
		return c;
	}
	
	public static Cell makeNumberCell(Workbook wb, Row r, int index, double val, int precision) {
		Cell c = makeCell(r, index, val);
		c.setCellStyle(Excel.numberCell(wb, precision));
		return c;
	}
	
}
