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
				//System.out.println("Applying bold");
				applyBoldStyle(cs, wb);
			}
			else if(s >= BORDER_LEFT && s <= BORDER_ALL_THIN) {
				//System.out.println("Applying borders");
				applyBorderStyle(cs, s);
			}
			else if(s == GRAY_FILL) {
				//System.out.println("Applying gray fill");
				applyGrayFillStyle(cs);
			}
			else if(s >= NUMBER && s <= NUMBER_4) {
				//System.out.println("Applying number format");
				applyNumberStyle(cs, wb, s - NUMBER);
			}
			else if(s == DATE) {
				//System.out.println("Applying date format");
				applyDateStyle(cs, wb);
			}
			else {
				//System.out.println("Invalid choice");
			}
		}
		//System.out.println(cs.getBorderBottom());
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
	
	/*
	private static CellStyle applyBottomBorderStyle(CellStyle cs) { return applyBottomBorderStyle(cs, false); }
	private static CellStyle applyBottomBorderStyle(CellStyle cs, boolean thin) {
		cs.setBorderBottom((thin ? CellStyle.BORDER_THIN : CellStyle.BORDER_MEDIUM));
		return cs;
	}
	
	private static CellStyle applyTopBorderStyle(CellStyle cs) {
		cs.setBorderTop(CellStyle.BORDER_MEDIUM);
		return cs;
	}
	
	private static CellStyle applyLeftBorderStyle(CellStyle cs) {
		cs.setBorderLeft(CellStyle.BORDER_MEDIUM);
		return cs;
	}
	
	private static CellStyle applyRightBorderStyle(CellStyle cs) {
		cs.setBorderRight(CellStyle.BORDER_MEDIUM);
		return cs;
	}
	
	
	private static CellStyle applyAllBorderStyle(CellStyle cs) { return applyAllBorderStyle(cs, BORDER_MEDIUM); }
	private static CellStyle applyAllBorderStyle(CellStyle cs, short type) {
		applyBorderStyle(cs, BORDER_ALL, type);
		return cs;
	}
	*/
	
	
	
	
	/*
	public static CellStyle boldBottomBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyBoldStyle(cs, wb);
		applyBorderStyle(cs, BORDER_BOTTOM);
		return cs;
	}
	
	public static CellStyle boldAllBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyBoldStyle(cs, wb);
		applyBorderStyle(cs, BORDER_ALL);
		return cs;
	}
	
	public static CellStyle boldBottomRightBorderStyle(Workbook wb) {
		CellStyle cs = boldBottomBorderStyle(wb);
		applyBorderStyle(cs, BORDER_RIGHT);
		return cs;
	}

	public static CellStyle bottomBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyBorderStyle(cs, BORDER_BOTTOM);
		return cs;
	}
	
	
	public static CellStyle bottomRightBorderStyle(Workbook wb) {
		CellStyle cs = bottomBorderStyle(wb);
		applyBorderStyle(cs, BORDER_RIGHT);
		return cs;
	}
	
	public static CellStyle leftBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyBorderStyle(cs, BORDER_LEFT);
		return cs;
	}
	
	public static CellStyle rightBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyBorderStyle(cs, BORDER_RIGHT);
		return cs;
	}
	
	public static CellStyle grayFillStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyGrayFillStyle(cs);
		return cs;
	}
	
	public static CellStyle grayFillBorderStyle(Workbook wb, Collection<Short> bs) {
		CellStyle cs = grayFillStyle(wb);
		for(Short b : bs) {
			applyBorderStyle(cs, b);
		}
		
		return cs;
	}
	
	public static CellStyle grayFillLeftBorderStyle(Workbook wb) {
		CellStyle cs = grayFillStyle(wb);
		applyBorderStyle(cs, BORDER_LEFT);
		return cs;
	}
	
	public static CellStyle grayFillRightBorderStyle(Workbook wb) {
		CellStyle cs = grayFillStyle(wb);
		applyBorderStyle(cs, BORDER_RIGHT);
		return cs;
	}
	
	public static CellStyle grayFillBottomLeftBorderStyle(Workbook wb) {
		CellStyle cs = grayFillLeftBorderStyle(wb);
		applyBottomBorderStyle(cs);
		return cs;
	}
	public static CellStyle numberStyle(Workbook wb, boolean leftBorder) { return numberStyle(wb, leftBorder, 2); }
	public static CellStyle numberStyle(Workbook wb, boolean leftBorder, int prec) {
		CellStyle cs = wb.createCellStyle();
		CellStyles.applyNumberStyle(cs, wb, prec);
		return cs;
	}
	
	public static CellStyle dateStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyDateStyle(cs, wb);
		return cs;
	}
	
	
	*/
	////////////////////////////////////////
	// OLD NON-UPDATED VERSIONS
	/*
	public static CellStyle dateStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		cs.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat(DATE_FORMAT));
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
	
	public static CellStyle grayFillCell(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		cs.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
		return cs;
	}
	*/
}
