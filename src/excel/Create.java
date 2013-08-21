package excel;

import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import excel.sheet.Intake;
import excel.sheet.Settings;
import excel.sheet.StoredData;

public class Create {
	
	
	public static void main(String[] _) throws IOException {
		//System.out.println(Settings.dateColumn());
		String workbook = Excel.DEFAULT_PATH + Excel.DEFAULT_WORKBOOK_NAME;
		File wb = new File(workbook);
		Excel.writeToExcel(createWorkbook());
		Desktop.getDesktop().open(wb);
		/*
		
		if(!wb.exists()) {
			Excel.writeToExcel(createWorkbook());
		}
		try {
			Workbook w = Excel.readExcel(wb);
			Intake.updateSheet(w);
			Excel.writeToExcel(w);
			Desktop.getDesktop().open(wb);
		} catch (IOException e) {
			e.printStackTrace();
		}
		*/
	}
	
	private static Workbook createWorkbook() {
		Workbook wb = new XSSFWorkbook();
		Settings s = new Settings(wb);
		s.createSettingsSheet();
		Intake i = new Intake(wb);
		i.createIntakeSheet();
		StoredData sd = new StoredData(wb);
		sd.createStoredDataSheet();
		return wb;
	}
	
}
