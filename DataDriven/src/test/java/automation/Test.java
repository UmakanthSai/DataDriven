package automation;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {

	public static void main(String[] args) throws IOException {

		//FileInputStream is a class which as power to read any file
		FileInputStream file = new FileInputStream("D://uma//Study//DemoData.xlsx");
		
		//XSSFWorkbook is a class which get the access to workbook of excel file
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		int sheets = workbook.getNumberOfSheets();//getting number of sheets available
		
		//Iterating to number of sheets to find required sheet
		for(int i=0;i<sheets;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase("DataDriven")) {
				XSSFSheet sheet = workbook.getSheetAt(i);//we have taken sheet access
				
				Iterator<Row> rows = sheet.iterator();//Iterating through row
				Row row = rows.next();
				Iterator<Cell> cell =row.cellIterator();//Iterating through column
				int colIndex = 0;
				while(cell.hasNext()) {
					Cell value = cell.next();
					if(value.getStringCellValue().equalsIgnoreCase("StepDefinition")) {
					colIndex =	value.getColumnIndex();
					}
				}
				System.out.println(colIndex);
				
				while(rows.hasNext()) {
					
				Row r = rows.next();
				if(r.getCell(colIndex).getStringCellValue().equalsIgnoreCase("Group B")) {
					
					Iterator<Cell> c = r.cellIterator();
					while(c.hasNext()) {
						System.out.println( c.next().getStringCellValue());
					}
				}
				}
			}
		}
		
	}

}
