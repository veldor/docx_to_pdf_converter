package net.veldor.docx_to_pdf.utils;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class Parser {
    public Parser(String file, String destinationDir) {
        // получу файл
        File sourceFile = FilesHandler.getFile(file);
        File destination = FilesHandler.getFile(destinationDir);
        if(sourceFile.isFile() && destination.isDirectory()){
            // обработаю файл
            BufferedInputStream mInputStream = new BufferedInputStream(new FileInputStream(f));
            XWPFDocument mDocument = null;
            try {
                mDocument = new XWPFDocument(mInputStream);
            } catch (IOException e) {
                e.printStackTrace();
            }
            XWPFWordExtractor extractor = new XWPFWordExtractor(mDocument);
            String value = extractor.getText();
        }
        else{
            System.out.println("Неверный путь к файлу или к папке назначения");
        }
    }
}
