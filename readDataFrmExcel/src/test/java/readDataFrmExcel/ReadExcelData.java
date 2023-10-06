package readDataFrmExcel;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelData {
	public static void main (String [] args) throws IOException{
		//Path of the excel file
		FileInputStream fs = new FileInputStream("C:\\Users\\bhava\\OneDrive\\Documents\\DemoData.xlsx");
		//Creating a workbook
		XSSFWorkbook workbook = new XSSFWorkbook(fs);
		XSSFSheet sheet = workbook.getSheetAt(0);
		//row0
		Row row = sheet.getRow(0);
		Cell cell = row.getCell(0);
		System.out.println(sheet.getRow(0).getCell(0));
		Row row1 = sheet.getRow(1);
		Cell cell1 = row1.getCell(1);
		System.out.println(sheet.getRow(0).getCell(1));
		Row row2=sheet.getRow(2);
		Cell cell2=row2.getCell(2);
		System.out.println(sheet.getRow(0).getCell(2));
		
		//row1
		Row row3 = sheet.getRow(1);
		Cell cell3 = row3.getCell(0);
		System.out.println(sheet.getRow(1).getCell(0));
		Row row4 = sheet.getRow(1);
		Cell cell4 = row4.getCell(1);
		System.out.println(sheet.getRow(1).getCell(1));
		Row row5 = sheet.getRow(1);
		Cell cell5 = row5.getCell(1);
		System.out.println(sheet.getRow(1).getCell(2));
		
		//row2 
		Row row6 = sheet.getRow(2);
		Cell cell6 = row6.getCell(0);
		System.out.println(sheet.getRow(2).getCell(0));
		Row row7 = sheet.getRow(1);
		Cell cell7 = row7.getCell(1);
		System.out.println(sheet.getRow(2).getCell(1));
		Row row8 = sheet.getRow(1);
		Cell cell8 = row8.getCell(1);
		System.out.println(sheet.getRow(2).getCell(2));
		}

}
