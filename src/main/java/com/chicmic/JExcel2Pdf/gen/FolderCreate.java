package com.chicmic.JExcel2Pdf.gen;

import java.io.File;

public class FolderCreate {
    public String createFolder(String folderName, String path) {
        File folder = new File(path, folderName);

        if (!folder.exists()) {
            boolean created = folder.mkdirs();
            if (created) {
                System.out.println("Folder created: " + folder.getAbsolutePath());
                return folder.getAbsolutePath();
            } else {
                System.err.println("Failed to create folder: " + folder.getAbsolutePath());
            }
        } else {
            System.out.println("Folder already exists: " + folder.getAbsolutePath());
        }

        return null; // Return null in case of failure

    }
    public static String pathBefore(String path) {
        File file = new File(path);

        if (file.exists()) {
            File parent = file.getParentFile();
            if (parent != null) {
                return parent.getAbsolutePath();
            }
        }

        return null; // Return null if the path doesn't exist or there's no parent folder
    }


}
