package net.veldor.docx_to_pdf.utils;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

public class Parser {

    public static final String DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    public static final String PDF_MIME = "application/pdf";

    public Parser(String file, String destinationDir) {
        // получу файл
        File sourceFile = FilesHandler.getFile(file);
        if(!sourceFile.isFile()){
            System.out.println("Не найден файл");
            return;
        }
        try {
            String contentType = null;
            if(file.endsWith(".docx")){
                contentType = DOCX_MIME;
            }
            else if(file.endsWith(".pdf")){
                contentType = PDF_MIME;
            }
            if (contentType != null && contentType.equals(DOCX_MIME)) {
                File destination = FilesHandler.getFile(destinationDir);
                if (sourceFile.isFile() && destination.isDirectory()) {
                    // обработаю файл
                    BufferedInputStream mInputStream = new BufferedInputStream(new FileInputStream(sourceFile));
                    XWPFDocument mDocument;
                    try {
                        String executionNumber;
                        String executionArea;
                        mDocument = new XWPFDocument(mInputStream);
                        XWPFWordExtractor extractor = new XWPFWordExtractor(mDocument);
                        String value = extractor.getText();
                        // разобью текст по переносам строк
                        executionNumber = StringsHandler.getExecutionNumber(value);
                        // получу номер обследования и область обследования из документа
                        if (executionNumber != null && !executionNumber.isEmpty()) {
                            // если в папке назначения нет файла с подобным именем
                            if (!(new File(destination, executionNumber + ".pdf")).exists()) {
                                PdfOptions options = PdfOptions.create();
                                OutputStream out = new FileOutputStream(new File(destination, executionNumber + ".pdf"));
                                PdfConverter.getInstance().convert(mDocument, out, options);
                                System.out.println("Заключение добавлено");
                                System.out.println(executionNumber + ".pdf");
                            } else {
                                int counter = 0;
                                File existentFile;
                                while (true) {
                                    if (counter == 0) {
                                        existentFile = new File(destination, executionNumber + ".pdf");
                                    } else {
                                        existentFile = new File(destination, executionNumber + "-" + counter + ".pdf");
                                    }
                                    if (existentFile.exists()) {
                                        PDDocument document = PDDocument.load(existentFile);
                                        if (!document.isEncrypted()) {
                                            PDFTextStripper stripper = new PDFTextStripper();
                                            String text = stripper.getText(document);
                                            // найду область обследования
                                            executionArea = StringsHandler.getExecutionArea(value);
                                            String oldExecutionArea = StringsHandler.getExecutionArea(text);
                                            document.close();
                                            if (executionArea != null && executionArea.equals(oldExecutionArea)) {
                                                // перезапишу имеющийся файл
                                                // файл с данной областью обследования не найден, сохраню его отдельно
                                                PdfOptions options = PdfOptions.create();
                                                OutputStream out = new FileOutputStream(existentFile);
                                                PdfConverter.getInstance().convert(mDocument, out, options);
                                                System.out.println("Файл заключения перезаписан");
                                                System.out.println(existentFile.getName());
                                                break;
                                            }
                                        }
                                    } else {
                                        // файл с данной областью обследования не найден, сохраню его отдельно
                                        PdfOptions options = PdfOptions.create();
                                        OutputStream out = new FileOutputStream(existentFile);
                                        PdfConverter.getInstance().convert(mDocument, out, options);
                                        System.out.println("Добавлено дополнительное заключение");
                                        System.out.println(existentFile.getName());
                                        break;
                                    }
                                    counter++;
                                }
                            }
                        } else {
                            System.out.println("В тексте не найден номер обследования");
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                } else {
                    System.out.println("Неверный путь к файлу или к папке назначения");
                }
            }
            else if(contentType != null){
                // обработаю PDF
                File destination = FilesHandler.getFile(destinationDir);
                if(sourceFile.isFile() && destination.isDirectory()){
                    // прочитаю данные из файла
                    PDDocument document = PDDocument.load(sourceFile);
                    if (!document.isEncrypted()) {
                        PDFTextStripper stripper = new PDFTextStripper();
                        String text = stripper.getText(document);
                        document.close();
                        String executionNumber = StringsHandler.getExecutionNumber(text);
                        if(executionNumber != null && !executionNumber.isEmpty()){
                            if (!(new File(destination, executionNumber + ".pdf")).exists()) {
                                System.out.println("Заключение добавлено");
                                // копирую файл в папку
                                boolean result = sourceFile.renameTo(new File(destination, executionNumber + ".pdf"));
                                if(!result){
                                    System.out.println("Не удалось создать файл");
                                }
                                else{
                                    System.out.println(executionNumber + ".pdf");
                                }
                            }
                            else{
                                int counter = 0;
                                File existentFile;
                                // найду область обследования
                                String executionArea = StringsHandler.getExecutionArea(text);
                                while (true) {
                                    if (counter == 0) {
                                        existentFile = new File(destination, executionNumber + ".pdf");
                                    } else {
                                        existentFile = new File(destination, executionNumber + "-" + counter + ".pdf");
                                    }
                                    if (existentFile.exists()) {
                                        document = PDDocument.load(existentFile);
                                        if (!document.isEncrypted()) {
                                            stripper = new PDFTextStripper();
                                            text = stripper.getText(document);
                                            String oldExecutionArea = StringsHandler.getExecutionArea(text);
                                            document.close();
                                            if (executionArea != null && executionArea.equals(oldExecutionArea)) {
                                                boolean deleteResult = existentFile.delete();
                                                if(!deleteResult){
                                                    System.out.println("error delete target file");
                                                }
                                                // перезапишу имеющийся файл
                                                System.out.println("Файл заключения перезаписан");
                                                boolean result = sourceFile.renameTo(existentFile);
                                                if(!result){
                                                    System.out.println("Не смог создать файл");
                                                }
                                                else{
                                                    System.out.println(existentFile.getName());
                                                }
                                                break;
                                            }
                                        }
                                    } else {
                                        document.close();
                                        // файл с данной областью обследования не найден, сохраню его отдельно
                                        System.out.println("Добавлено дополнительное заключение");
                                        boolean result = sourceFile.renameTo(existentFile);
                                        if(!result){
                                            System.out.println("Не смог создать файл");
                                        }
                                        else{
                                            System.out.println(existentFile.getName());
                                        }
                                        break;
                                    }
                                    counter++;
                                }
                            }
                        }
                    }
                }
            }
            else{
                System.out.println("Неверный формат файла");
            }

        } catch (Exception e) {
            System.out.println("Have error: " + e.getMessage());
            StringWriter errors = new StringWriter();
            e.printStackTrace(new PrintWriter(errors));
            System.out.println(errors.toString());
        }
    }
}
