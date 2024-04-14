package Data.Excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {

	public static void main(String[] args) throws IOException {
		
		FileInputStream fis = new FileInputStream("N:\\Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		
		int numOfSheet = workbook.getNumberOfSheets();
		for(int i=0; i < numOfSheet; i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("Data1")) {
				XSSFSheet sheet = workbook.getSheetAt(i);
				
				Iterator<Row> rows =  sheet.iterator(); //sheet is collection of rows
				Row firstRow = rows.next();
				Iterator<Cell> ce = firstRow.cellIterator();   //row is collection of cells
				int k=0; int column = 0;
				while(ce.hasNext()) {
					Cell value = ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("TestCase")) {
						column = k;
					}
					k++;
				}
				System.out.println(column);
				while(rows.hasNext()) {
					Row r = rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("password")) {
						Iterator <Cell> c = r.cellIterator();
						while(c.hasNext()) {
							System.out.println(c.next().getStringCellValue());
						}
					}
				}
			}
		}

	}

}
