package com.chicmic.JExcel2Pdf.gen;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

import static com.chicmic.JExcel2Pdf.gen.ExcelRead.excelReadAndSort;
import static com.chicmic.JExcel2Pdf.gen.ExcelSorter.excelReadAndSort2;


public class GenApplication {

	public static void main(String[] args) {
		try {
			ClassLoader classLoader = GenApplication.class.getClassLoader();
			InputStream inputStream = classLoader.getResourceAsStream("application.properties");

			Properties properties = new Properties();
			properties.load(inputStream);
			String path = properties.getProperty("folder_path");

			String targetFileName = "For Bank (August).xlsx";
//			System.out.println("\u001B[31m " + targetFileName + " \u001B[0m");
			File excelFile = searchForExcelFile(new File(path), targetFileName);
			excelReadAndSort2(excelFile); // Read and sort cells based on column D
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	public static File searchForExcelFile(File directory, String targetFileName) {
		File[] files = directory.listFiles();
		if (files != null) {
			for (File file : files) {
				if (file.isDirectory()) {
					File found = searchForExcelFile(file, targetFileName);
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
