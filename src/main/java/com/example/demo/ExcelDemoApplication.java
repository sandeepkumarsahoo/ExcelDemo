package com.example.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ExcelDemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(ExcelDemoApplication.class, args);
		System.out.println("Reading excel sheet");
		writeFormulaInExcelFile();
	}

	static void writeFormulaInExcelFile() {
		try {
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet spreadsheet = workbook.createSheet("formula");
			XSSFRow row = spreadsheet.createRow(1);
			XSSFCell cell = row.createCell(1);

			cell.setCellValue("A = ");
			cell = row.createCell(2);
			cell.setCellValue(2);
			row = spreadsheet.createRow(2);
			cell = row.createCell(1);
			cell.setCellValue("B = ");
			cell = row.createCell(2);
			cell.setCellValue(4);
			row = spreadsheet.createRow(3);
			cell = row.createCell(1);
			cell.setCellValue("Total = ");
			cell = row.createCell(2);

			// Create SUM formula
			cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
			cell.setCellFormula("SUM(C2:C3)");
			cell = row.createCell(3);
			cell.setCellValue("SUM(C2:C3)");
			row = spreadsheet.createRow(4);
			cell = row.createCell(1);
			cell.setCellValue("POWER =");
			cell = row.createCell(2);

			// Create POWER formula
			cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
			cell.setCellFormula("POWER(C2,C3)");
			cell = row.createCell(3);
			cell.setCellValue("POWER(C2,C3)");
			row = spreadsheet.createRow(5);
			cell = row.createCell(1);
			cell.setCellValue("MAX = ");
			cell = row.createCell(2);

			// Create MAX formula
			cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
			cell.setCellFormula("MAX(C2,C3)");
			cell = row.createCell(3);
			cell.setCellValue("MAX(C2,C3)");
			row = spreadsheet.createRow(6);
			cell = row.createCell(1);
			cell.setCellValue("FACT = ");
			cell = row.createCell(2);

			// Create FACT formula
			cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
			cell.setCellFormula("FACT(C3)");
			cell = row.createCell(3);
			cell.setCellValue("FACT(C3)");
			row = spreadsheet.createRow(7);
			cell = row.createCell(1);
			cell.setCellValue("SQRT = ");
			cell = row.createCell(2);

			// Create SQRT formula
			cell.setCellType(XSSFCell.CELL_TYPE_FORMULA);
			cell.setCellFormula("SQRT(C5)");
			cell = row.createCell(3);
			cell.setCellValue("SQRT(C5)");
			workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
			FileOutputStream out = new FileOutputStream(new File("C:\\Users\\AC21398\\Desktop\\test\\Book2.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("fromula.xlsx written successfully");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	static void readDataFromExcelFile() {
		try {
			FileInputStream file = new FileInputStream(new File("C:\\Users\\AC21398\\Desktop\\test\\Book1.xlsx"));
			Workbook workbook = new XSSFWorkbook(file);
			Sheet datatypeSheet = workbook.getSheetAt(0);
			Iterator<Row> iterator = datatypeSheet.iterator();
			while (iterator.hasNext()) {

				Row currentRow = iterator.next();
				Iterator<Cell> cellIterator = currentRow.iterator();

				while (cellIterator.hasNext()) {

					Cell currentCell = cellIterator.next();
					// getCellTypeEnum shown as deprecated for version 3.15
					// getCellTypeEnum ill be renamed to getCellType starting from version 4.0
					if (currentCell.getCellTypeEnum() == CellType.STRING) {
						System.out.print(currentCell.getStringCellValue());
					} else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
						System.out.print(currentCell.getNumericCellValue());
					}

				}
				System.out.println();
			}
			workbook.close();
		} catch (Exception exception) {
			exception.printStackTrace();
		}
	}

}
