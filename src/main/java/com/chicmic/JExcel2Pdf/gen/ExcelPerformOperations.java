package com.chicmic.JExcel2Pdf.gen;


import com.chicmic.JExcel2Pdf.gen.Util.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import static com.chicmic.JExcel2Pdf.gen.DateConverter.getTodaysDate;
import static com.chicmic.JExcel2Pdf.gen.FolderCreate.pathBefore;

public class ExcelPerformOperations {
    Integer indexOfRecipientColumnD = 3;
    Integer indexOfSOFTEXNumberColumnF = 5;
    // HashMap for updating document with specific text at location paragaraph index and run index
    HashMap<String, Pair<Integer, Integer>> textParaRunIndexHashMap = new HashMap<>();
    double billAmount = 0.0; // Initialize the billAmount column g
    double chargesAmount = 0.0; // Initialize the charges column h
    double finalBillAmount = 0.0; // Initialize the final bill column i
    String currentDate = getTodaysDate();


    public void excelPerformOperations(File excelFile) throws IOException {
        textParaRunIndexHashMap.put(String.valueOf(billAmount), new Pair<Integer, Integer>(9, 3));
        textParaRunIndexHashMap.put(String.valueOf(chargesAmount), new Pair<Integer, Integer>(10, 3));
        textParaRunIndexHashMap.put(String.valueOf(finalBillAmount), new Pair<Integer, Integer>(11, 3));
        textParaRunIndexHashMap.put(currentDate, new Pair<Integer, Integer>(1, 2));

        String excelFilePath = excelFile.getParent();

        FolderCreate folderCreate = new FolderCreate();
        DocxFileOperations docxFileOperations = new DocxFileOperations();

        String resultantFilePath = folderCreate.createFolder("Annexure 1", excelFilePath);
        if (resultantFilePath == null) {
            System.out.println("Returning");
            return;
        }


        FileInputStream fis = new FileInputStream(excelFile);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        // Convert the sheet's rows to a list for sorting, skipping the first row
        List<Row> rows = new ArrayList<>();
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            rows.add(row);
        }


        // use this to find the document paragraph index and run index by entering path of doc file
        docxFileOperations.getParagraphAndRunIndices(excelFile.getAbsolutePath());

        String prevD = "";
        String prevF = "";
        int heirarchyIndex = 0;
        String currentWorkingDirectory = null;
        for (int i = 0; i < rows.size(); i++) {
            Row sortedRow = rows.get(i);
            Cell currentCellD = sortedRow.getCell(indexOfRecipientColumnD);
            Cell currentCellB = sortedRow.getCell(indexOfRecipientColumnD - 2);
            if (currentCellD != null) {
                String currentD = currentCellD.toString();
                String currentF = sortedRow.getCell(indexOfRecipientColumnD + 2).toString();
                String currentB = currentCellB.toString();
                double cellValBillAmount = sortedRow.getCell(indexOfRecipientColumnD + 3).getNumericCellValue();
                double cellValChargesAmount = sortedRow.getCell(indexOfRecipientColumnD + 4).getNumericCellValue();
                double cellValFinalBillAmount = sortedRow.getCell(indexOfRecipientColumnD + 5).getNumericCellValue();

                if (prevD.equals(currentD)) {
                    if (currentF.equals(prevF)) {
//                            System.out.println("\u001B[34m in prevF if "  + currentF + "\u001B[0m");

                        billAmount += cellValBillAmount;
                        chargesAmount += cellValChargesAmount;
                        finalBillAmount += cellValFinalBillAmount;
                    } else {

                        docxFileOperations.updateTextAtPosition(excelFilePath, currentWorkingDirectory, textParaRunIndexHashMap);
                        heirarchyIndex++;
                        currentWorkingDirectory = folderCreate.createFolder(String.valueOf(heirarchyIndex), pathBefore(currentWorkingDirectory)); // create folder with name = '1'
                    }
//                        System.out.println("currentD = " + currentD+  " currentF = " + currentF.toString());
                } else {
//                        System.out.println("Else currentD = " + currentD+  " currentF = " + currentF.toString());
                    String path = folderCreate.createFolder(currentD, resultantFilePath);
                    heirarchyIndex = 1;
                    currentWorkingDirectory = folderCreate.createFolder(String.valueOf(heirarchyIndex), path); // create folder with name = '1'
                    billAmount = cellValBillAmount;
                    chargesAmount = cellValChargesAmount;
                    finalBillAmount = cellValFinalBillAmount;
                }
                prevD = currentD;
                prevF = currentF;
            }

        }

//            for (int i = 0; i < rows.size(); i++) {
//                Row sortedRow = rows.get(i);
//                Row newRow = newSheet.createRow(rowIndex++);
//
//                for (int j = 0; j < sortedRow.getLastCellNum(); j++) {
//                    Cell cell = newRow.createCell(j);
//                    Cell originalCell = sortedRow.getCell(j);
//
//                    if (originalCell != null) {
//                        cell.setCellValue(originalCell.toString());
//                    }
//                }
//
//                // Check if the values in column D are the same as the next row and update column G
//                if (i < rows.size() - 1) {
//                    Row nextRow = rows.get(i + 1);
//                    Cell currentCellD = sortedRow.getCell(columnIndexToSort);
//                    Cell nextCellD = nextRow.getCell(columnIndexToSort);
//
//                    if (currentCellD != null && nextCellD != null) {
//                        if (currentCellD.toString().equals(nextCellD.toString())) {
//                            // Get the cells in column G and update their values
//                            Cell currentCellG = newRow.createCell(6); // Assuming G is column 7 (0-based index)
//                            Cell nextCellG = nextRow.getCell(6);
//                            double currentCellValueG = currentCellG.getNumericCellValue();
//                            double nextCellValueG = nextCellG.getNumericCellValue();
//                            sum += currentCellValueG;
//
//                        } else {
//                            // Values in column D are different; set the sum and reset it
//                            System.out.println("sum = " + sum);
//                            sum = 0; // Reset the sum
//                        }
//                    }
//                }
//            }

        // Write the new workbook to an output file


        System.out.println("Excel sheet sorted, column G updated, and new file generated based on column D.");
    }
}


