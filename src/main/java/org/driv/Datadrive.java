package org.driv;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Datadrive {
	public static void main(String[] args) throws IOException {
		
	
	File f = new File("C:\\Users\\kavin\\eclipse-workspace\\Datadriven\\excel\\inmakes.xlsx");
	 
	FileInputStream fis = new FileInputStream(f);
	
	Workbook fb = new XSSFWorkbook(fis);
	Sheet mySheet = fb.getSheet("inmakes");
	
	for (int i = 0; i < mySheet.getPhysicalNumberOfRows(); i++) {
	Row r	=  mySheet.getRow(i);
	
	for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
	Cell c = r.getCell(j);
	int cellType = c.getCellType();
	if (cellType == 1) {
		String value= c.getStringCellValue();
		System.out.println(value);
		
	}
		
	else if (DateUtil.isCellDateFormatted(c)) {
		Date datecell = c.getDateCellValue();
		SimpleDateFormat s = new SimpleDateFormat("dd-MMM-yy"); 
		String value = s.format(datecell);
		System.out.println(value);
		
	}
	else {
		double d = c.getNumericCellValue();
		long l =(long)d;
		String value = String.valueOf(l);
		System.out.println(value);
	}	
	}
	}
}
}

