package com.chicmic.JExcel2Pdf.gen;

import com.chicmic.JExcel2Pdf.gen.Util.Pair;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;

public class DocxFileOperations {
    String updatedDocumentName = "L1 Request letter for Submission of Export doc.docx";

    public  void updateTextAtPosition(String inputFilePath, String outputFilePath, int paragraphIndex, int runIndex, String newText) throws IOException {
        outputFilePath += "/" + updatedDocumentName;
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        // Check if the specified paragraph and run indices are within valid ranges
        if (paragraphIndex >= 0 && paragraphIndex < document.getParagraphs().size()) {
            XWPFParagraph paragraph = document.getParagraphs().get(paragraphIndex);
            if (runIndex >= 0 && runIndex < paragraph.getRuns().size()) {
                XWPFRun run = paragraph.getRuns().get(runIndex);
                run.setText(newText, 0); // Replace the existing text with the new text
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
        document.write(fileOutputStream);
        fileOutputStream.close();
    }
    public void updateTextAtPosition(String inputFilePath, String outputFilePath, HashMap<String, Pair<Integer, Integer>> textParaRunIndexMap) throws IOException {
        outputFilePath += "/" + updatedDocumentName;

        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        for (String text : textParaRunIndexMap.keySet()) {
            Pair<Integer, Integer> paraRunIndices = textParaRunIndexMap.get(text);
            int paragraphIndex = paraRunIndices.first();
            int runIndex = paraRunIndices.second();

            // Check if the specified paragraph index is within a valid range
            if (paragraphIndex >= 0 && paragraphIndex < document.getParagraphs().size()) {
                XWPFParagraph paragraph = document.getParagraphs().get(paragraphIndex);

                // Ensure the run index is within the valid range
                int maxRunIndex = paragraph.getRuns().size() - 1;
                if (runIndex >= 0 && runIndex <= maxRunIndex) {
                    XWPFRun run = paragraph.getRuns().get(runIndex);
                    run.setText(text, 0); // Replace the existing text with the new text
                }
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
        document.write(fileOutputStream);
        fileOutputStream.close();
    }
    public  void updateTextInRange(String inputFilePath, String outputFilePath, int paragraphIndex, int runStartIndex, int runEndIndex, String newText) throws IOException {
        outputFilePath += "/" + updatedDocumentName;

        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        // Check if the specified paragraph index is within a valid range
        if (paragraphIndex >= 0 && paragraphIndex < document.getParagraphs().size()) {
            XWPFParagraph paragraph = document.getParagraphs().get(paragraphIndex);

            // Ensure the run indices are within the valid range
            int maxRunIndex = paragraph.getRuns().size() - 1;
            if (runStartIndex >= 0 && runStartIndex <= maxRunIndex && runEndIndex >= runStartIndex && runEndIndex <= maxRunIndex) {
                // Create a new run with the updated text
                XWPFRun updatedRun = paragraph.insertNewRun(runStartIndex);
                updatedRun.setText(newText);

                // Remove the runs in the specified range
                for (int i = runStartIndex + 1; i <= runEndIndex; i++) {
                    paragraph.removeRun(runStartIndex + 1);
                }
            }
        }

        FileOutputStream fileOutputStream = new FileOutputStream(outputFilePath);
        document.write(fileOutputStream);
        fileOutputStream.close();
    }
    public  void getParagraphAndRunIndices(String inputFilePath) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(inputFilePath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (int paragraphIndex = 0; paragraphIndex < paragraphs.size(); paragraphIndex++) {
            XWPFParagraph paragraph = paragraphs.get(paragraphIndex);
            System.out.println("\u001B[34m Paragraph Index: " + paragraphIndex + "\u001B[0m");

            List<XWPFRun> runs = paragraph.getRuns();
            for (int runIndex = 0; runIndex < runs.size(); runIndex++) {
                XWPFRun run = runs.get(runIndex);
                System.out.println("\u001B[35m"+ "Run Index: " + runIndex);
                System.out.println("Run Text: " + run.getText(0) + "\u001B[0m"); // Print the text of the run
            }
        }
    }
}
