package com.montran;

import java.io.FileInputStream;
import java.io.FileWriter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.opencsv.CSVWriter;

public class Main2 {

	public static void main(String[] args) {

		Scanner scanner = new Scanner(System.in);
		System.out.print("Ingresa el nombre del archivo Excel (con extensión .xlsx): ");
		String excelFilePath = scanner.nextLine();

		System.out.print("Ingresa el nombre del proyecto: ");
		String projectName = scanner.nextLine().trim().replace(" ", "_");

		System.out.print("Ingresa el nombre del archivo CSV de salida (con .csv): ");
		String outputCsv = scanner.nextLine();
		
		Main2 main = new Main2();
		String lastFuncionality ="";

		try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFilePath));
				CSVWriter csvWriter = new CSVWriter(new FileWriter(outputCsv))) {

			XSSFSheet mainSheet = (XSSFSheet) workbook.getSheetAt(5);
			Iterator<Row> rowIterator = mainSheet.iterator();
			rowIterator.next(); // Saltar encabezado

			// Encabezados CSV
			String[] headers = { "Issue Key", "Test case name", "Description", "Test Step", "Test Result",
					"Test Data" };
			csvWriter.writeNext(headers);

			if (rowIterator.hasNext()) rowIterator.next();
			while (rowIterator.hasNext()) {
// C:\Users\Johanna Gutierrez\eclipse-workspace\2021.12\Jira\TestCase_Import\src\main\resources\CAVAPY_CSD_1 1.xlsx
				Row row = rowIterator.next();

				// Evitar procesar la fila de encabezados
				if (row.getCell(0) == null || row.getCell(1) == null || row.getCell(2) == null
						|| row.getCell(3) == null)
					continue;

				
				//para que se repita si esta vacio 
				String funcionality = "";		 
		
				Cell funcionalityCell = row.getCell(3); 

			    funcionality = (funcionalityCell != null && !funcionalityCell.toString().trim().isEmpty())
			            ? funcionalityCell.toString().trim()
			            : lastFuncionality;

			    // Actualiza solo si había un nuevo valor
			    if (!funcionality.equals(lastFuncionality)) {
			        lastFuncionality = funcionality;
			    }
				
				// Leer valores de la fila
				String feature = main.getCellAsString(row.getCell(4));
				String testCaseNo = main.getCellAsString(row.getCell(7));
				String testScenario = main.getCellAsString(row.getCell(5));
				String syntax = main.getCellAsString(row.getCell(9));
				String testSteps = main.getCellAsString(row.getCell(10));
				String testOutput = main.getCellAsString(row.getCell(11));
				String dataSheetName = main.getCellAsString(row.getCell(12));

				// Concatenar Test case name
				String testCaseName = projectName + " / " + funcionality + " / " + feature + "/" + testCaseNo + "_"
						+ testScenario;

				// Test Step y Test Result
				String testStep = syntax + " "+testSteps; // Columna Test Step
				String testResult = testOutput; // Columna Test Result
				String testData = "";

				if (dataSheetName != null && !dataSheetName.isEmpty()) {
					XSSFSheet  dataSheet = (XSSFSheet) workbook.getSheet(dataSheetName);
					if (dataSheet != null) {
						Iterator<Row> dataRows = dataSheet.iterator();
						Row headerRow = dataRows.next();
						List<String> headersData = new ArrayList<>();
						for (Cell cell : headerRow)
							headersData.add(cell.getStringCellValue().trim().toLowerCase());

						while (dataRows.hasNext()) {
							Row dataRow = dataRows.next();
							Cell idCell = dataRow.getCell(0);
							if (idCell == null || !testCaseNo.equals(idCell.getStringCellValue().trim()))
								continue;

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

			System.out.println("Archivo CSV generado correctamente: " + outputCsv);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public String getCellAsString(Cell cell) {
	    if (cell == null) return "";

	    switch (cell.getCellType()) {
	        case STRING:
	            return cell.getStringCellValue().trim();

	        case NUMERIC:
	            if (DateUtil.isCellDateFormatted(cell)) {
	                return cell.getDateCellValue().toString(); // Or format with SimpleDateFormat
	            } else {
	                return String.valueOf(cell.getNumericCellValue());
	            }

	        case BOOLEAN:
	            return String.valueOf(cell.getBooleanCellValue());

	        case FORMULA:
	            // Try to evaluate it as string first
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
