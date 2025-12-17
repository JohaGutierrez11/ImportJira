package com.montran.service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
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

            // Headers
            csvWriter.writeNext(new String[]{
                "Issue Key", "Test case name", "Description",
                "Test Step", "Test Result", "Test Data"
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
        var rowIterator = sheet.iterator();
        if (rowIterator.hasNext()) rowIterator.next(); // encabezado
        if (rowIterator.hasNext()) rowIterator.next();

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            if (row.getCell(3) == null) continue;

            String funcionality = getCellAsString(row.getCell(3));
            if (funcionality.isEmpty()) funcionality = lastFuncionality;
            lastFuncionality = funcionality;

            String feature = getCellAsString(row.getCell(4));
            String testCaseNo = getCellAsString(row.getCell(7));
            String testScenario = getCellAsString(row.getCell(5));
            String syntax = getCellAsString(row.getCell(9));
            String testSteps = getCellAsString(row.getCell(10));
            String testOutput = getCellAsString(row.getCell(11));
            String dataSheetName = getCellAsString(row.getCell(12));

            String testCaseName = projectName + "/" + funcionality + "/" +
                    feature + "/" + testCaseNo + "_" + testScenario;

            if (testCaseName.length() > 255) {
                advertencias.append(feature)
                        .append(" ")
                        .append(testCaseNo)
                        .append(" truncado a 255 caracteres\n");
                testCaseName = testCaseName.substring(0, 254);
            }

            String testData = resolveTestData(
                    workbook, dataSheetName, testCaseNo);

            csvWriter.writeNext(new String[]{
                    "",
                    testCaseName,
                    "",
                    syntax + " " + testSteps,
                    testOutput,
                    testData
            });
        }
    }

    private String resolveTestData(
            Workbook workbook,
            String sheetName,
            String testCaseNo
    ) {
        if (sheetName == null || sheetName.isEmpty()) return "";

        Sheet dataSheet = workbook.getSheet(sheetName);
        if (dataSheet == null) return "";

        var rows = dataSheet.iterator();
        if (!rows.hasNext()) return "";

        Row header = rows.next();
        int columns = header.getLastCellNum();

        while (rows.hasNext()) {
            Row r = rows.next();
            Cell id = r.getCell(0);
            if (id == null || !testCaseNo.equals(id.toString().trim()))
                continue;

            StringBuilder sb = new StringBuilder();
            for (int i = 1; i < columns; i++) {
                sb.append(header.getCell(i).getStringCellValue())
                  .append(": ")
                  .append(getCellAsString(r.getCell(i)));
                if (i < columns - 1) sb.append(" | ");
            }
            return sb.toString();
        }
        return "";
    }

    private String getCellAsString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.toString();
            default -> "";
        };
    }
}
