package com.chicmic.JExcel2Pdf.gen;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

import static com.chicmic.JExcel2Pdf.gen.DateConverter.getTodaysDate;
import static com.chicmic.JExcel2Pdf.gen.FolderCreate.pathBefore;

public class ExcelSorter {


    public File excelManager(File file) {
        file = excelSortByColumn(file, 3);
        file = excelSortByColumn(file, 5);
        return file;
    }
    public File excelSortByColumn(File file, int columnIndex) {
        File sortedFile = null;
        try {
            FileInputStream fis = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0); // Assuming it's the first sheet

            // Create a custom comparator for sorting by the specified column
            Comparator<Row> comparator = (r1, r2) -> {
                Cell cell1 = r1.getCell(columnIndex);
                Cell cell2 = r2.getCell(columnIndex);
                return cell1.toString().compareTo(cell2.toString());
            };

            // Convert the sheet's rows to a list for sorting, skipping the first row
            List<Row> rows = new ArrayList<>();
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                rows.add(row);
            }
            rows.sort(comparator);

            // Create a new Excel workbook and sheet for the sorted data
            Workbook newWorkbook = new XSSFWorkbook();
            Sheet newSheet = newWorkbook.createSheet("Sorted Data");

            int rowIndex = 0;
            for (Row sortedRow : rows) {
                Row newRow = newSheet.createRow(rowIndex++);

                for (int j = 0; j < sortedRow.getLastCellNum(); j++) {
                    Cell cell = newRow.createCell(j);
                    Cell originalCell = sortedRow.getCell(j);

                    if (originalCell != null) {
                        cell.setCellValue(originalCell.toString());
                    }
                }
            }

            // Generate a unique file name for the sorted Excel file
            String originalFileName = file.getName();
            String sortedFileName = "sorted_" + originalFileName;
            sortedFile = new File(sortedFileName);

            // Write the new workbook to the sortedFile
            FileOutputStream fos = new FileOutputStream(sortedFile);
            newWorkbook.write(fos);
            fos.close();

            // Close the input file
            fis.close();

            System.out.println("Excel sheet sorted based on column " + columnIndex + " and a new file generated: " + sortedFileName);
        } catch (Exception e) {
            e.printStackTrace();
        }

        return sortedFile;
    }






}
