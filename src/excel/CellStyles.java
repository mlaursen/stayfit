package excel;


import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

public class CellStyles {
	public static final String DATE_FORMAT = "dd-mmm";
	public static final String NUMBER_FORMAT = "0.00";
	public static final short BORDER_LEFT = 0;
	public static final short BORDER_RIGHT = 1;
	public static final short BORDER_BOTTOM = 2;
	public static final short BORDER_TOP = 3;
	public static final short BORDER_ALL = 4;
	public static final short BORDER_LEFT_THIN = 5;
	public static final short BORDER_RIGHT_THIN = 6;
	public static final short BORDER_BOTTOM_THIN = 7;
	public static final short BORDER_TOP_THIN = 8;
	public static final short BORDER_ALL_THIN = 9;
	public static final short BOLD = 10;
	public static final short GRAY_FILL = 11;
	public static final short NUMBER = 12;
	public static final short NUMBER_1 = 13;
	public static final short NUMBER_2 = 14;
	public static final short NUMBER_4 = 15;
	public static final short DATE = 16;

	public static Cell applyStyles(Set<Short> styles, Cell c) {
		Workbook wb = c.getSheet().getWorkbook();
		CellStyle cs = wb.createCellStyle();
		for(Short s : styles) {
			if(s == BOLD) {
				applyBoldStyle(cs, wb);
			}
			else if(s >= BORDER_LEFT && s <= BORDER_ALL_THIN) {
				applyBorderStyle(cs, s);
			}
			else if(s == GRAY_FILL) {
				applyGrayFillStyle(cs);
			}
			else if(s >= NUMBER && s <= NUMBER_4) {
				applyNumberStyle(cs, wb, s - NUMBER);
			}
			else if(s == DATE) {
				applyDateStyle(cs, wb);
			}
			else {
				//System.out.println("Invalid choice");
			}
		}
		c.setCellStyle(cs);
		return c;
	}
	
	private static CellStyle applyBoldStyle(CellStyle cs, Workbook wb) {
		Font f = wb.createFont();
		f.setBoldweight(Font.BOLDWEIGHT_BOLD);
		cs.setFont(f);
		return cs;
	}
	
	private static CellStyle applyBorderStyle(CellStyle cs, short dir) {
		short type = (dir >= BORDER_LEFT_THIN && dir <= BORDER_ALL_THIN) ? CellStyle.BORDER_THIN : CellStyle.BORDER_MEDIUM;
		if(dir == BORDER_LEFT || dir == BORDER_LEFT_THIN || dir == BORDER_ALL)
			cs.setBorderLeft(type);
		else if(dir == BORDER_RIGHT || dir == BORDER_RIGHT_THIN || dir == BORDER_ALL)
			cs.setBorderRight(type);
		else if(dir == BORDER_BOTTOM || dir == BORDER_BOTTOM_THIN || dir == BORDER_ALL)
			cs.setBorderBottom(type);
		else if(dir == BORDER_TOP || dir == BORDER_TOP_THIN || dir == BORDER_ALL)
			cs.setBorderTop(type);
		return cs;
	}
	
	private static CellStyle applyGrayFillStyle(CellStyle cs) {
		cs.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
		return cs;
	}
	
	private static CellStyle applyNumberStyle(CellStyle cs, Workbook wb, int prec) {
		String format = "0.";
		for(int i = 0; i < prec; i++)
			format += 0;
		return applyNumberStyle(cs, wb, format);
	}
	private static CellStyle applyNumberStyle(CellStyle cs, Workbook wb, String format) {
		cs.setDataFormat(wb.createDataFormat().getFormat(format));
		return cs;
	}
	
	private static CellStyle applyDateStyle(CellStyle cs, Workbook wb) {
		cs.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat(DATE_FORMAT));
		return cs;
	}
	
}
