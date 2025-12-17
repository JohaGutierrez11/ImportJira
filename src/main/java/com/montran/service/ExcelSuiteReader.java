package com.montran.service;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.montran.model.TestSuite;

public class ExcelSuiteReader {
	
	public List<TestSuite> readSuites(String excelPath) throws Exception{
		List<TestSuite> suites = new ArrayList<>();
		
		try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelPath))) {
            for (Sheet sheet : workbook) {
            	String sheetName = sheet.getSheetName();

                if (sheetName.startsWith("Test Suite")) {
                    suites.add(new TestSuite(sheetName));
                }
            }
        }
		
		return suites;		
	}

}
