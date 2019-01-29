// YouTube link: https://www.youtube.com/watch?v=mlasUKBC_Y4

package test;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Test3 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		
		try (Workbook wb =WorkbookFactory.create(new File ("C:\\Users\\PC User1\\Desktop\\SeleniumFiles\\ExcelTestData\\data.xls"))){
			
			Sheet sheet = wb.getSheetAt(0);
			
			int rowStart = sheet.getFirstRowNum();
			int rowEnd = sheet.getLastRowNum();
			
			for (int i = rowStart; i < rowEnd; i++) {
				Row row = sheet.getRow(i);
				
				for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
					Cell cell = row.getCell(j);
					System.out.println(cell.getStringCellValue());
				}
				
				System.out.println("-----------------------");
			}
			
		} catch (Exception e) {
			e.printStackTrace();
			
		}
		
	}

}
