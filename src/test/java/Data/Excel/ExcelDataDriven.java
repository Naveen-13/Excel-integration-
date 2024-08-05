package Data.Excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ExcelDataDriven {
	
	@Test
	public void test() throws IOException {
		ExcelDataDriven obj = new ExcelDataDriven();
		ArrayList<String> a = obj.method();
		System.out.println(a.get(1));
		System.out.println(a.get(2));
		System.out.println(a.get(3));
		
	}
	
	public ArrayList<String> method() throws IOException {
		ArrayList<String> arr = new ArrayList<>();
		FileInputStream file = new FileInputStream("N:\\Book1.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(file);
		XSSFSheet sheet  = wb.getSheet("sheet2");
		Iterator<Row> rows = sheet.rowIterator();
		Row firstRow = rows.next();
		Iterator<Cell> cells = firstRow.cellIterator();
		int k=0;
		int column=0;
		while(cells.hasNext()) {
			Cell cellValue=cells.next();
			if(cellValue.getStringCellValue().equalsIgnoreCase("testcases")) {
				column=k;
			}
			k++;
		}
		
		while(rows.hasNext()) {
			Row r =rows.next();
			if(r.getCell(column).getStringCellValue().equalsIgnoreCase("purchase")) {
				Iterator<Cell> c = r.cellIterator();
				while(c.hasNext()) {
					Cell ce = c.next();
					switch(ce.getCellType()) {
						case STRING:
							arr.add(ce.getStringCellValue());
							break;
					case NUMERIC:
						arr.add(NumberToTextConverter.toText(ce.getNumericCellValue()));
						break;
					}
					
				}
			}
		}
		return arr;
	}

}

