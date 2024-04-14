package Data.Excel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class DataP {
	
	@Test(dataProvider="data")
	public void dataMethod(String a, String b, String c, String d) {
		System.out.println(a + b + c + d);
	}
	@DataProvider(name="data")
	public Object[][] getData() throws IOException{
		DataFormatter format = new DataFormatter();
		FileInputStream fis = new FileInputStream("N:\\Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int colCount = row.getLastCellNum();
		Object data[][] = new Object[rowCount-1][colCount];
		for(int i = 0; i<rowCount-1;i++) {
			row = sheet.getRow(i+1);
			for(int j = 0; j<colCount; j++) {
				XSSFCell cell = row.getCell(j);
				data[i][j] = format.formatCellValue(cell);
			}
		}
		
		return data;		
	}
	
}
