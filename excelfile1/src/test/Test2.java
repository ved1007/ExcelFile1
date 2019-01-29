
// YouTube link: https://www.youtube.com/watch?v=go3SsnvSpOo

package test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test2 {

	public static void main(String[] args) throws IOException {

		FileInputStream fis = new FileInputStream (new File("C:\\Users\\PC User1\\Desktop\\SeleniumFiles\\ExcelTestData.xls"));
		
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		
		XSSFSheet sheet = wb.getSheetAt(0);
		
		FormulaEvaluator forlulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
		
		for (Row row : sheet) {
			for (Cell cell : row) {
				switch(forlulaEvaluator.evaluateInCell(cell).getCellType())
				{
				
		//		case Cell.CELL_TYPE_NUMERIC:
					
				}System.out.println(cell.getNumericCellValue() + "\t\t");
				}
				}
			}
	}

