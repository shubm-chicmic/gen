//package com.chicmic.JExcel2Pdf.gen;
//
//public class folderCreate {
//    public void createHeirarchy() {
//        if (i < rows.size() - 1) {
//            Row nextRow = rows.get(i + 1);
//            Cell currentCellD = sortedRow.getCell(columnIndexToSort);
//            Cell nextCellD = nextRow.getCell(columnIndexToSort);
//
//            if (currentCellD != null && nextCellD != null) {
//                if (currentCellD.toString().equals(nextCellD.toString())) {
//                    // Get the cells in column F and update their values
//                    Cell currentCellF = newRow.createCell(5); // Assuming F is column 6 (0-based index)
//                    Cell nextCellF = nextRow.getCell(5);
//                    double currentCellValueF = currentCellF.getNumericCellValue();
//                    double nextCellValueF = nextCellF.getNumericCellValue();
//
//                    currentCellF.setCellValue(currentCellValueF + nextCellValueF);
//                }
//            }
//        }
//    }
//    }
//}
