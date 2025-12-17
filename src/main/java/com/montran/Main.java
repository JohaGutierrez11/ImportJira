package com.montran;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.OutputStreamWriter;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.opencsv.CSVWriter;

public class Main {

    public String processExcelAndGenerateCsv(String excelFilePath, String projectName, String outputCsv) throws Exception {
        String lastFuncionality = "";
        StringBuilder advertencias = new StringBuilder();

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFilePath));
             CSVWriter csvWriter = new CSVWriter(new OutputStreamWriter(new FileOutputStream(outputCsv), StandardCharsets.UTF_8))) {

            XSSFSheet mainSheet = (XSSFSheet) workbook.getSheet("Test Suite");  //siempre con el nombre de Test Suite
            Iterator<Row> rowIterator = mainSheet.iterator();
            rowIterator.next(); // Saltar encabezado

            // Encabezados CSV
            String[] headers = { "Issue Key", "Test case name", "Description", "Test Step", "Test Result", "Test Data" };
            csvWriter.writeNext(headers);

            if (rowIterator.hasNext()) rowIterator.next();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Evitar procesar la fila de encabezados
                if (row.getCell(0) == null || row.getCell(1) == null || row.getCell(2) == null || row.getCell(3) == null)
                    continue;

                String funcionality = "";		 
                Cell funcionalityCell = row.getCell(3); 
                funcionality = (funcionalityCell != null && !funcionalityCell.toString().trim().isEmpty())
                        ? funcionalityCell.toString().trim()
                        : lastFuncionality;

                if (!funcionality.equals(lastFuncionality)) {
                    lastFuncionality = funcionality;
                }

                // Leer valores de la fila
                String feature = getCellAsString(row.getCell(4));
                String testCaseNo = getCellAsString(row.getCell(7));
                String testScenario = getCellAsString(row.getCell(5));
                String syntax = getCellAsString(row.getCell(9));
                String testSteps = getCellAsString(row.getCell(10));
                String testOutput = getCellAsString(row.getCell(11));
                String dataSheetName = getCellAsString(row.getCell(12));

                // Concatenar Test case name
                String testCaseName = projectName + "/" + funcionality + "/" + feature + "/" + testCaseNo + "_" + testScenario;
                
                if(testCaseName.length()>255) {
                	 advertencias.append( feature +" " + testCaseNo + " Tiene texto truncado a 255 caracteres.\n");
                	 testCaseName = testCaseName.substring(0,254);
                }
                

                // Test Step y Test Result
                String testStep = syntax + " " + testSteps;
                String testResult = testOutput;
                String testData = "";

                if (dataSheetName != null && !dataSheetName.isEmpty()) {
                    XSSFSheet dataSheet = (XSSFSheet) workbook.getSheet(dataSheetName);
                    System.out.println(dataSheetName);
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
                            testData = sb.toString();
                            break;
                        }
                    } else {
                        System.err.println("Hoja '" + dataSheetName + "' no encontrada.");
                    }
                }

                String[] record = { "", testCaseName, "", testStep, testResult, testData };
                csvWriter.writeNext(record);
            }

        } catch (Exception e) {
            e.printStackTrace();
            throw new Exception("Error al procesar el archivo: " + e.getMessage());
        }
        
        return advertencias.toString();
    }

    public String getCellAsString(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BLANK:
                return "";
            default:
                return "";
        }
    }
}
