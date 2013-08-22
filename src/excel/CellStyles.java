package excel;

import java.awt.Color;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

public class CellStyles {
	public static final String DATE_FORMAT = "dd-mmm";
	public static final String NUMBER_FORMAT = "0.00";

	
	private static CellStyle applyBoldStyle(CellStyle cs, Workbook wb) {
		Font f = wb.createFont();
		f.setBoldweight(Font.BOLDWEIGHT_BOLD);
		cs.setFont(f);
		return cs;
	}
	
	private static CellStyle applyBottomBorderStyle(CellStyle cs) {
		cs.setBorderBottom(CellStyle.BORDER_MEDIUM);
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
	
	private static CellStyle applyAllBorderStyle(CellStyle cs) {
		applyTopBorderStyle(cs);
		applyBottomBorderStyle(cs);
		applyLeftBorderStyle(cs);
		applyRightBorderStyle(cs);
		return cs;
	}
	
	private static CellStyle applyGrayFillStyle(CellStyle cs) {
		cs.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
		return cs;
	}
	
	private static CellStyle applyNumberStyle(CellStyle cs, Workbook wb) { return applyNumberStyle(cs, wb, NUMBER_FORMAT); }
	private static CellStyle applyNumberStyle(CellStyle cs, Workbook wb, int prec) {
		String format = "0.";
		for(int i = 0; i < prec; i++)
			format += "0";
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
	
	
	/**
	 * Bolds, and puts a medium border on the bottom of a cell
	 * @param wb
	 * @return
	 */
	public static CellStyle boldBottomBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyBoldStyle(cs, wb);
		applyBottomBorderStyle(cs);
		return cs;
	}
	
	public static CellStyle boldAllBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyBoldStyle(cs, wb);
		applyAllBorderStyle(cs);
		return cs;
	}
	
	public static CellStyle boldBottomRightBorderStyle(Workbook wb) {
		CellStyle cs = boldBottomBorderStyle(wb);
		applyRightBorderStyle(cs);
		return cs;
	}

	public static CellStyle bottomBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyBottomBorderStyle(cs);
		return cs;
	}
	
	
	public static CellStyle bottomRightBorderStyle(Workbook wb) {
		CellStyle cs = bottomBorderStyle(wb);
		applyRightBorderStyle(cs);
		return cs;
	}
	
	public static CellStyle leftBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyLeftBorderStyle(cs);
		return cs;
	}
	
	public static CellStyle rightBorderStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyRightBorderStyle(cs);
		return cs;
	}
	
	public static CellStyle grayFillStyle(Workbook wb) {
		CellStyle cs = wb.createCellStyle();
		applyGrayFillStyle(cs);
		return cs;
	}
	
	public static CellStyle grayFillBorderStyle(Workbook wb, List<String> bs) {
		CellStyle cs = grayFillStyle(wb);
		if(bs.contains("l"))
			applyLeftBorderStyle(cs);
		
		if(bs.contains("r"))
			applyRightBorderStyle(cs);
		
		if(bs.contains("b"))
			applyBottomBorderStyle(cs);
		
		if(bs.contains("t"))
			applyTopBorderStyle(cs);
		
		return cs;
	}
	
	public static CellStyle grayFillLeftBorderStyle(Workbook wb) {
		CellStyle cs = grayFillStyle(wb);
		applyLeftBorderStyle(cs);
		return cs;
	}
	
	public static CellStyle grayFillRightBorderStyle(Workbook wb) {
		CellStyle cs = grayFillStyle(wb);
		applyRightBorderStyle(cs);
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
