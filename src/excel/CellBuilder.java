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
	
	public static Cell boldBottomBorderCell(Row r, int index, Object o) {
		return boldBottomBorderCell(makeCell(r, index, o));
	}
	
	public static Cell boldBottomBorderCell(Cell c) {
		CellStyle cs = c.getCellStyle();
		//CellStyles.applyBorderStyle(cs, CellStyles.BORDER_BOTTOM);
		c.setCellStyle(cs);
		//c.setCellStyle(CellStyles.boldBottomBorderStyle(c.getSheet().getWorkbook()));
		return c;
	}
	
	public static Cell boldBottomRightBorderCell(Row r, int index, Object o) {
		return boldBottomRightBorderCell(makeCell(r, index, o));
	}
	
	public static Cell boldBottomRightBorderCell(Cell c) {
		CellStyle cs = c.getCellStyle();
		//CellStyles.applyBorderStyle(cs, CellStyles.BORDER_BOTTOM);
		//CellStyles.applyBorderStyle(cs, CellStyles.BORDER_RIGHT);
		//CellStyles.applyBoldStyle(cs, c.getSheet().getWorkbook());
		//c.setCellStyle(CellStyles.boldBottomRightBorderStyle(c.getSheet().getWorkbook()));
		return c;
	}
	
	public static Cell bottomRightBorderCell(Row r, int i, Object o) {
		return bottomRightBorderCell(makeCell(r, i, o));
	}
	
	public static Cell bottomRightBorderCell(Cell c) {
		//c.setCellStyle(CellStyles.bottomRightBorderStyle(c.getSheet().getWorkbook()));
		return c;
	}
	
	public static Cell boldAllBorderCell(Cell c) {
		//c.setCellStyle(CellStyles.boldAllBorderStyle(c.getSheet().getWorkbook()));
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
	
	/*
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
	*/
}
