package com.chicmic.JExcel2Pdf.gen;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;




public class GenApplication {

	private static final String FILE_NAME = "For Bank (August).xlsx";
	private static final String path = System.getProperty("user.dir");

	public static void main(String[] args) {
		try {
			ClassLoader classLoader = GenApplication.class.getClassLoader();
			InputStream inputStream = classLoader.getResourceAsStream("application.properties");

			Properties properties = new Properties();
			properties.load(inputStream);


			File excelFile = new File(path + "/" +FILE_NAME); // Excel File Read in current Directory
			ExcelSorter excelSorter = new ExcelSorter();
			File sortedExcelFile = excelSorter.excelManager(excelFile); // Excel File Sort acc to column D then F

			ExcelPerformOperations excelPerformOperations = new ExcelPerformOperations();
			excelPerformOperations.excelPerformOperations(sortedExcelFile);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	public static File searchForFile(File directory, String targetFileName) {
		File[] files = directory.listFiles();
		if (files != null) {
			for (File file : files) {
				if (file.isDirectory()) {
					File found = searchForFile(file, targetFileName);
					if (found != null) {
						return found; // Return the found file if it's in a subdirectory
					}
				} else {
					if (file.getName().equals(targetFileName)) {
						return file; // Return the found file
					}
				}
			}
		}
		return null; // File not found
	}

}
