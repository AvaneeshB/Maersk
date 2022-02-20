package excelOperations;

import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

public class execution {

public static ArrayList<String> header(File a) throws IOException{
		
		ArrayList<String> res = new ArrayList<String> () ;
		String path=".\\data\\"+a.getName();
		
		FileInputStream inputStream = new FileInputStream(path);
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		
		int rows = sheet.getLastRowNum();
		int cols = sheet.getRow(0).getLastCellNum();
		
		for(int r=0;r<1;r++) {
			XSSFRow row = sheet.getRow(r);
			for(int c=0;c<cols;c++)
			{
				XSSFCell cell = row.getCell(c);
				
				res.add(cell.getStringCellValue());
				
			}
		}
		return res;
}
}
