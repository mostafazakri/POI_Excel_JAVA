package com.mostafa.introPOI;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class IntroPOI {

	public static void main(String[] args) throws IOException {

		String filePath = "src//ExcelXLSX.xlsx";
        File excelFile = new File(filePath);
        FileInputStream fis = new FileInputStream(excelFile);
        Workbook workbook;
       
        if(filePath.endsWith(".xls")) {
             workbook = new HSSFWorkbook(fis);
        }else {
             workbook = new XSSFWorkbook(fis);
        }
       
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(3);
        Cell cell = row.getCell(2);
        String cellValue =  cell.getStringCellValue();
  
        System.out.println(cellValue);
        workbook.close();
	}

}
