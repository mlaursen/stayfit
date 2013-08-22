package excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public class CellStyles {
	public static final String DATE_FORMAT = "dd-mmm";
	public static final String NUMBER_FORMAT = "0.00";
	
	public static CellStyle dateStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		cs.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat(DATE_FORMAT));
		return cs;
	}
	
	public static CellStyle sideThickBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		cs.setBorderLeft(CellStyle.BORDER_THICK);
		cs.setBorderRight(CellStyle.BORDER_THICK);
		return cs;
	}
	
	public static CellStyle topThickBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		cs.setTopBorderColor(CellStyle.BORDER_THICK);
		return cs;
	}
	
	
	/**
	 * Creates a cell style with the default bold font size and bolds the font.
	 * @param wb	The workbook to add the bold style to
	 * @return		The bold cell style.
	 */
	public static CellStyle boldCell(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		Font f = wb.createFont();
		f.setBoldweight(Font.BOLDWEIGHT_BOLD);
		cs.setFont(f);
		return cs;
	}
	
	public static CellStyle numberCell(Workbook wb, String format) {
		CellStyle cs = wb.createCellStyle();
		cs.setDataFormat(wb.createDataFormat().getFormat(format));
		return cs;
	}
	
	public static CellStyle numberCell(Workbook wb) { return numberCell(wb, NUMBER_FORMAT); }
	public static CellStyle numberCell(Workbook wb, int precision) {
		String format = "0.";
		for(int i = 0; i < precision; i++)
			format += "0";
		return numberCell(wb, format);
	}
}
