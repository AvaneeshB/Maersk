package excelOperations;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class writing {
	
	public static void merge(File a, File b,int k) throws IOException
	{
		
		//old
		String pathA=".\\data\\"+a.getName();
		String pathB=".\\data\\"+b.getName();
		FileInputStream inputStreamA = new FileInputStream(pathA);
		FileInputStream inputStreamB = new FileInputStream(pathB);
		
		XSSFWorkbook oldworkbookA = new XSSFWorkbook(inputStreamA);
		XSSFSheet oldsheetA = oldworkbookA.getSheetAt(0);
		
		XSSFWorkbook oldworkbookB = new XSSFWorkbook(inputStreamB);
		XSSFSheet oldsheetB = oldworkbookB.getSheetAt(0);
		
		//new
		@SuppressWarnings("resource")
		XSSFWorkbook newworkbook = new XSSFWorkbook();
		XSSFSheet newsheet = newworkbook.createSheet("Merged Sheet");
		
		
		
		int rows = oldsheetA.getLastRowNum()+1;
		int cols = oldsheetA.getRow(1).getLastCellNum();
		
		//for A
		for(int r=0;r<rows;r++) 
		{
		
			XSSFRow rowA = oldsheetA.getRow(r);
			XSSFRow newrow = newsheet.createRow(r);
			
			for(int c=0;c<cols;c++)
			{
				XSSFCell cellA = rowA.getCell(c);
				XSSFCell newcell = newrow.createCell(c);
				switch(cellA.getCellType()) 
				{
				case STRING: String x = cellA.getStringCellValue();
					newcell.setCellValue(x);
							 break;
				
				case NUMERIC: int y = (int) cellA.getNumericCellValue();
					newcell.setCellValue(y); 
					
 							  break;
				}
			}
			
		}
		
		//for B		
		for(int f=1;f<rows-1;f++) 
		{
			XSSFRow rowB = oldsheetB.getRow(f);
			XSSFRow newrow = newsheet.createRow(f+rows-1);
			for(int c=0;c<cols;c++)
			{
				XSSFCell cellB = rowB.getCell(c);
				XSSFCell newcell = newrow.createCell(c);
				switch(cellB.getCellType()) 
				{
				case STRING: String x= cellB.getStringCellValue(); 
							
					newcell.setCellValue(x);;
							 break;
				
				case NUMERIC: int y = (int) cellB.getNumericCellValue();
					newcell.setCellValue(y); 
					
 							  break;
				}
			}
		}
		String name = ".\\data\\new"+Integer.toString(k)+".xlsx";
		FileOutputStream outputStream=new FileOutputStream(name);
		newworkbook.write(outputStream);
		outputStream.close();
		System.out.println("Created a new excel sheet ...");
		
	}
}	