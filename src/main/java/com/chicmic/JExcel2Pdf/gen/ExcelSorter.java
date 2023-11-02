package com.chicmic.JExcel2Pdf.gen;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

public class ExcelSorter {
    public static void excelReadAndSort2(File file) {
        try {
            FileInputStream fis = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0); // Assuming it's the first sheet

            // Create a custom comparator for sorting by column D (0-based index)
            int columnIndexToSort = 3; // Column D is index 3 (0-based index)

            Comparator<Row> comparator = (r1, r2) -> {
                Cell cell1 = r1.getCell(columnIndexToSort);
                Cell cell2 = r2.getCell(columnIndexToSort);
                return cell1.toString().compareTo(cell2.toString());
            };

            // Convert the sheet's rows to a list for sorting, skipping the first row
            List<Row> rows = new ArrayList<>();
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                rows.add(row);
            }

            // Sort the rows using the custom comparator
            rows.sort(comparator);

            // Create a new Excel workbook and sheet for the sorted data
            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("Sorted Data");

            int rowIndex = 0;
            double sum = 0.0; // Initialize the sum

            for (int i = 0; i < rows.size(); i++) {
                Row sortedRow = rows.get(i);
                Row newRow = newSheet.createRow(rowIndex++);

                for (int j = 0; j < sortedRow.getLastCellNum(); j++) {
                    Cell cell = newRow.createCell(j);
                    Cell originalCell = sortedRow.getCell(j);

                    if (originalCell != null) {
                        cell.setCellValue(originalCell.toString());
                    }
                }

                // Check if the values in column D are the same as the next row and update column G
                if (i < rows.size() - 1) {
                    Row nextRow = rows.get(i + 1);
                    Cell currentCellD = sortedRow.getCell(columnIndexToSort);
                    Cell nextCellD = nextRow.getCell(columnIndexToSort);

                    if (currentCellD != null && nextCellD != null) {
                        if (currentCellD.toString().equals(nextCellD.toString())) {
                            // Get the cells in column G and update their values
                            Cell currentCellG = newRow.createCell(6); // Assuming G is column 7 (0-based index)
                            Cell nextCellG = nextRow.getCell(6);
                            double currentCellValueG = currentCellG.getNumericCellValue();
                            double nextCellValueG = nextCellG.getNumericCellValue();
                            sum += currentCellValueG;

                        } else {
                            // Values in column D are different; set the sum and reset it
                            System.out.println("sum = " + sum);
                            sum = 0; // Reset the sum
                        }
                    }
                }
            }

            // Write the new workbook to an output file
            FileOutputStream fos = new FileOutputStream("output.xlsx"); // Replace with your output file name
            newWorkbook.write(fos);
            fos.close();

            // Close the input file
            fis.close();

            System.out.println("Excel sheet sorted, column G updated, and new file generated based on column D.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


}
