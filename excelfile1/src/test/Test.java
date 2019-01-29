package test;

import java.io.File;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Test {
	
	private static XSSFWorkbook wb;
	private static XSSFSheet sh;
	private static FileInputStream fis;
	private static FileOutputStream fos;
	private static XSSFRow row;
	private static XSSFCell cell;
	
	public static void main(String[] args) throws Exception {
		
		File src = new File ("C:\\Users\\PC User1\\Desktop\\SeleniumFiles\\ExcelTestData\\TestData.xlsx");
		FileInputStream fis = new FileInputStream(src);
		wb  = new XSSFWorkbook(fis);
		sh = wb.getSheet("Sheet1");
		row = sh.getRow(1);
		cell = row.getCell(1);
		
		int noOfRows = sh.getLastRowNum();
		System.out.println(noOfRows);

	}

}
