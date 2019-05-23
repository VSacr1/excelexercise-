package com.qa.testexcelexersice;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

public class TestExcel {

	
	@Test
	public void loginTest() throws Exception {
		FileInputStream file = new FileInputStream("C:\\Users\\Admin\\Desktop/example.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file); 
		XSSFSheet sheet = workbook.getSheetAt(0);
		//reading 
		for(int rowNum = 0; rowNum < sheet.getPhysicalNumberOfRows(); rowNum++)
		{
			for(int colNum = 0; colNum < sheet.getRow(rowNum).getPhysicalNumberOfCells(); colNum++)
			{
				XSSFCell cell = sheet.getRow(rowNum).getCell(colNum);
				String userCell = cell.getStringCellValue(); 
				System.out.println(userCell);
			}
		}
		
		file.close(); 
		
		
	}
	
	@Test
	public void writeTest() throws Exception {
		//Write
			FileInputStream file = new FileInputStream("C:\\Users\\Admin\\Desktop/example.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file); 
			XSSFSheet sheet = workbook.getSheetAt(0);
		
				XSSFRow row = sheet.createRow(1); 
				XSSFCell cell = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); 
				
				if(cell == null)
				{
					cell = row.createCell(1); 
				}
				cell.setCellValue("hello");
				
				FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Admin\\Desktop\\example.xlsx"); 
				
				workbook.write(fileOut);
				fileOut.flush(); 
				fileOut.close(); 
				file.close(); 
	}
	
}
