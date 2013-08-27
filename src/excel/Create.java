package excel;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import javax.swing.JFileChooser;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import excel.sheet.Settings;

public class Create {
	
	
	public static void main(String[] _) throws IOException {
		JFileChooser fc = new JFileChooser();
		fc.setSelectedFile(new File(Excel.DEFAULT_WORKBOOK_NAME));
		int r = fc.showOpenDialog(null);
		String wb = (r == JFileChooser.CANCEL_OPTION ? Excel.DEFAULT_PATH : fc.getSelectedFile().getAbsolutePath());
		wb = wb.endsWith(".xlsx") ? wb : wb + Excel.DEFAULT_WORKBOOK_NAME;
		Excel.writeToExcel(createWorkbook(), wb);
		Desktop.getDesktop().open(new File(wb));
	}
	
	private static Workbook createWorkbook() {
		Workbook wb = new XSSFWorkbook();
		Settings s = new Settings(wb);
		s.createSettingsSheet();
		return wb;
	}
	
}
