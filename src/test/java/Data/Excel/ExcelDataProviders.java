package Data.Excel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class ExcelDataProviders {
	DataFormatter dataformat = new DataFormatter();
	@Test(dataProvider="datadriver")
	public void method1(String name, String email, String gender, String department) {
		System.out.println(name + " "+ email + " " + gender + " " + department);
	}

	@DataProvider(name = "datadriver")
	public Object[][] getData() throws IOException {
		
		FileInputStream file = new FileInputStream("N:\\Book1.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(file);
		XSSFSheet sheet = wb.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int cellCount = row.getLastCellNum();
		Object data[][] = new Object[rowCount - 1][cellCount];

		for (int i = 0; i < rowCount-1; i++) {
			row = sheet.getRow(i + 1);
			for (int j = 0; j < cellCount; j++) {
				XSSFCell cell = row.getCell(j);
				data[i][j] = dataformat.formatCellValue(cell);
			}
		}
		return data;
	}

}
