package automation;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.util.SystemOutLogger;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {

	public ArrayList<String> ExcelData(String sheetName, String columnName, String rowName ) throws IOException {
ArrayList<String> data = new ArrayList<String>();
		
		//FileInputStream is a class which as power to read any file
		FileInputStream file = new FileInputStream("D://uma//Study//DemoData.xlsx");
		
		//XSSFWorkbook is a class which get the access to workbook of excel file
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		
		int sheets = workbook.getNumberOfSheets();//getting number of sheets available
		
		//Iterating to number of sheets to find required sheet
		for(int i=0;i<sheets;i++) {
			if(workbook.getSheetName(i).equalsIgnoreCase(sheetName)) {
				XSSFSheet sheet = workbook.getSheetAt(i);//we have taken sheet access
				
				Iterator<Row> rows = sheet.iterator();//Iterating through row
				Row row = rows.next();
				Iterator<Cell> cell =row.cellIterator();//Iterating through column
				int colIndex = 0;
				while(cell.hasNext()) {
					Cell value = cell.next();
					if(value.getStringCellValue().equalsIgnoreCase(columnName)) {
					colIndex =	value.getColumnIndex();
					}
				}
				System.out.println(colIndex);
				
				while(rows.hasNext()) {
					
				Row r = rows.next();
				if(r.getCell(colIndex).getStringCellValue().equalsIgnoreCase(rowName)) {
					
					Iterator<Cell> c = r.cellIterator();
					while(c.hasNext()) {
						Cell ce=c.next();
						if(ce.getCellType()==CellType.STRING) {
							data.add(ce.getStringCellValue());//passing into Array List
						}
						else {
							//USing NumberToTextConverter from Apache.poi developers to convert numeric to string  
							data.add(NumberToTextConverter.toText(ce.getNumericCellValue()));
						}
					}
				}
				}
			}
		}
		return data;
	}
	
	public static void main(String[] args) throws IOException {


	
		
	}

}
