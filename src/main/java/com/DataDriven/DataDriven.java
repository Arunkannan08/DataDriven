package com.DataDriven;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {
	public static void getParticularData() throws InvalidFormatException, IOException {
		File file = new File("C:\\Arunkumar\\SeleniumData\\TestData.xlsx");
		Workbook book = new XSSFWorkbook(file);
		Sheet sheet1 = book.getSheetAt(0);
		Row row = sheet1.getRow(0);
		Cell cell = row.getCell(0);

		int rowSize = sheet1.getLastRowNum();
		System.out.println(rowSize);
		int columnSize = sheet1.getRow(0).getLastCellNum();
		System.out.println(columnSize);

		System.out.println("The Row is ");
		System.out.println("========================");
		// To Read first row value
		for (int i = 0; i < columnSize; i++) {
			DataFormatter format = new DataFormatter();
			cell = row.getCell(i);
			String cellValue = format.formatCellValue(cell);
			System.out.print(cellValue);
			System.out.print(" ");
		}

		System.out.println();
		System.out.println("The Read ALL Data ");
		System.out.println("========================");

		// To Read all data
		for (int j = 0; j <= rowSize; j++) {
			row = sheet1.getRow(j);
			for (int i = 0; i < columnSize; i++) {
				DataFormatter format = new DataFormatter();
				cell = row.getCell(i);
				String cellValue = format.formatCellValue(cell);
				System.out.print(cellValue);
				System.out.print(" ");
			}
			System.out.println();
		}

		// To Format the given cell we can use dataFormat class
		/*
		 * DataFormatter format = new DataFormatter(); String cellValue String cellValue
		 * = format.formatCellValue(cell); System.out.println(cellValue);
		 */

	}

	public static void main(String[] args) throws InvalidFormatException, IOException {
		getParticularData();
	}

}
