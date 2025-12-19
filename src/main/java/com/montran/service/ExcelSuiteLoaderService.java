package com.montran.service;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.montran.model.TestSuite;

public class ExcelSuiteLoaderService {

    public List<TestSuite> loadTestSuites(String excelFilePath) throws Exception {

        List<TestSuite> suites = new ArrayList<>();

        try (Workbook workbook =
                     new XSSFWorkbook(new FileInputStream(excelFilePath))) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                String sheetName = workbook.getSheetName(i);

                if (sheetName.toLowerCase().startsWith("test suite")) {
                    suites.add(new TestSuite(sheetName));
                }
            }
        }
        return suites;
    }
}
