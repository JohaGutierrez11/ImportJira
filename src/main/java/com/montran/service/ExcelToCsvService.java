package com.montran.service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.montran.model.TestSuite;
import com.opencsv.CSVWriter;

public class ExcelToCsvService {

    public String generateCsv(
            String excelFilePath,
            String projectName,
            String outputCsv,
            List<TestSuite> selectedSuites
    ) throws Exception {

        StringBuilder advertencias = new StringBuilder();

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFilePath));
             CSVWriter csvWriter = new CSVWriter(
                     new OutputStreamWriter(
                             new FileOutputStream(outputCsv),
                             StandardCharsets.UTF_8))) {

            // Headers CSV
            csvWriter.writeNext(new String[]{
                    "Issue Key",
                    "Test case name",
                    "Description",
                    "Test Step",
                    "Test Result",
                    "Test Data"
            });

            for (TestSuite suite : selectedSuites) {
                if (!suite.isSelected()) continue;

                Sheet sheet = workbook.getSheet(suite.getName());
                if (sheet == null) continue;

                processSheet(sheet, workbook, projectName, csvWriter, advertencias);
            }
        }

        return advertencias.toString();
    }

    private void processSheet(
            Sheet sheet,
            Workbook workbook,
            String projectName,
            CSVWriter csvWriter,
            StringBuilder advertencias
    ) throws Exception {

        String lastFuncionality = "";

        Iterator<Row> rowIterator = sheet.iterator();
        if (rowIterator.hasNext()) rowIterator.next(); // header
        if (rowIterator.hasNext()) rowIterator.next(); // segunda fila (como tu Excel real)

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Validación fuerte (igual a tu Main original)
            if (row.getCell(0) == null ||
                row.getCell(1) == null ||
                row.getCell(2) == null ||
                row.getCell(3) == null) {
                continue;
            }

            String funcionality = getCellAsString(row.getCell(3));
            if (funcionality.isEmpty()) {
                funcionality = lastFuncionality;
            } else {
                lastFuncionality = funcionality;
            }

            String feature = getCellAsString(row.getCell(4));
            String testScenario = getCellAsString(row.getCell(5));
            String testCaseNo = getCellAsString(row.getCell(7));
            String syntax = getCellAsString(row.getCell(9));
            String testSteps = getCellAsString(row.getCell(10));
            String testOutput = getCellAsString(row.getCell(11));
            String dataSheetName = getCellAsString(row.getCell(12));

            String testCaseName =
                    projectName + "/" +
                    funcionality + "/" +
                    feature + "/" +
                    testCaseNo + "_" + testScenario;

            if (testCaseName.length() > 255) {
                advertencias.append(feature)
                        .append(" ")
                        .append(testCaseNo)
                        .append(" truncado a 255 caracteres.\n");
                testCaseName = testCaseName.substring(0, 254);
            }

            String testStep = syntax + " " + testSteps;
            String testData = resolveTestData(workbook, dataSheetName, testCaseNo);

            csvWriter.writeNext(new String[]{
                    "",
                    testCaseName,
                    "",
                    testStep,
                    testOutput,
                    testData
            });
        }
    }

    private String resolveTestData(
            Workbook workbook,
            String dataSheetName,
            String testCaseNo
    ) {

    	if (dataSheetName != null && !dataSheetName.isEmpty()) {
            XSSFSheet dataSheet = (XSSFSheet) workbook.getSheet(dataSheetName);
       
            if (dataSheet != null) {
                Iterator<Row> dataRows = dataSheet.iterator();
                Row headerRow = dataRows.next();
                List<String> headersData = new ArrayList<>();
                for (Cell cell : headerRow)
                    headersData.add(cell.getStringCellValue().trim());
                
              
                while (dataRows.hasNext()) {
                    Row dataRow = dataRows.next();
                    
                    Cell idCell = dataRow.getCell(0);
                    
                    if (idCell == null) {
                    	continue;
                    } else if (!testCaseNo.equals(idCell.getStringCellValue().trim())){
                    	continue;
                    }
                   

                    StringBuilder sb = new StringBuilder();
                    for (int i = 1; i < headersData.size(); i++) {
                        Cell c = dataRow.getCell(i);
                        String val = (c != null) ? c.toString().trim() : "";
                        sb.append(headersData.get(i)).append(": ").append(val);
                        if (i < headersData.size() - 1)
                            sb.append(" | ");
                    }
                    return sb.toString();
                }
            } else {
            	
            	String dataValue = dataSheetName.replace("\n", " ").replace("\r", " ");
            	  System.out.println(dataValue);
            	return dataValue;
            }
        }
		return "";
    }

    private String getCellAsString(Cell cell) {
        if (cell == null) return "";

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                    ? cell.getDateCellValue().toString()
                    : String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                try {
                    yield cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    yield String.valueOf(cell.getNumericCellValue());
                }
            }
            default -> "";
        };
    }
}
