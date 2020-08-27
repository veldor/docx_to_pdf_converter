package net.veldor.docx_to_pdf.utils;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.util.Random;

public class Parser {

    public static final String DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    public static final String DOC_MIME = "application/msword";
    public static final String PDF_MIME = "application/pdf";
    private static File destination;

    public static void parse(String file, String destinationDir) {
        // получу файл
        File sourceFile = FilesHandler.getFile(file);
        if (!sourceFile.isFile()) {
            System.out.println("Не найден файл");
            return;
        }
        destination = FilesHandler.getFile(destinationDir);
        if (!destination.isDirectory()) {
            System.out.println("Не найдена папка назначения");
            return;
        }
        try {
            String contentType = null;
            if (file.endsWith(".docx")) {
                contentType = DOCX_MIME;
                System.out.println("handle docx");
            } else if (file.endsWith(".pdf")) {
                contentType = PDF_MIME;
                System.out.println("handle pdf");
            } else if (file.endsWith(".doc")) {
                System.out.println("handle doc");
                contentType = DOC_MIME;
            }
            if (contentType != null && contentType.equals(DOCX_MIME)) {
                // обработаю .docx
                    // обработаю файл
                    try (BufferedInputStream mInputStream = new BufferedInputStream(new FileInputStream(sourceFile))) {
                        XWPFDocument mDocument;
                        String executionNumber;
                        mDocument = new XWPFDocument(mInputStream);
                        XWPFWordExtractor extractor = new XWPFWordExtractor(mDocument);
                        String value = extractor.getText();
                        // найду область обследования
                        String executionArea = StringsHandler.getExecutionArea(value);
                        if (executionArea != null) {
                            // разобью текст по переносам строк
                            executionNumber = StringsHandler.getExecutionNumber(value);
                            // получу номер обследования и область обследования из документа
                            if (executionNumber != null && !executionNumber.isEmpty()) {
                                // если в папке назначения нет файла с подобным именем
                                if (!(new File(destination, executionNumber + ".pdf")).exists()) {
                                    try {
                                        Converter.convertByMsWord(sourceFile, new File(destination, executionNumber + ".pdf"));
                                        System.out.println("created pdf by 1 method");
                                    } catch (Exception e) {
                                        Converter.convertByApache(mDocument, new File(destination, executionNumber + ".pdf"));
                                        System.out.println("created pdf by 2 method");
                                    }
                                    System.out.println(executionNumber + ".pdf");
                                } else {
                                    System.out.println("file with current name exists");
                                    int counter = 0;
                                    File existentFile;
                                    while (true) {
                                        System.out.println("tick check file");
                                        if (counter == 0) {
                                            existentFile = new File(destination, executionNumber + ".pdf");
                                        } else {
                                            existentFile = new File(destination, executionNumber + "-" + counter + ".pdf");
                                        }
                                        System.out.println("check file " + existentFile);
                                        if (existentFile.exists()) {
                                            System.out.println("file exists");
                                            PDDocument document = PDDocument.load(existentFile);
                                            if (!document.isEncrypted()) {
                                                PDFTextStripper stripper = new PDFTextStripper();
                                                String text = stripper.getText(document);
                                                String oldExecutionArea = StringsHandler.getExecutionArea(text);
                                                System.out.println("old execution area is " + oldExecutionArea);
                                                document.close();
                                                if (oldExecutionArea != null && executionArea.startsWith(oldExecutionArea)) {
                                                    // перезапишу имеющийся файл
                                                    // файл с данной областью обследования не найден, сохраню его отдельно
                                                    // попробую конвертировать файл 1 способом
                                                    try {
                                                        Converter.convertByMsWord(sourceFile, existentFile);
                                                        System.out.println("existent pdf rewrited by 1 method");
                                                    } catch (Exception e) {
                                                        Converter.convertByApache(mDocument, existentFile);
                                                        System.out.println("existent pdf rewrited by 2 method");
                                                    }
                                                    break;
                                                }
                                            }
                                        } else {
                                            // попробую конвертировать файл 1 способом
                                            try {
                                                Converter.convertByMsWord(sourceFile, existentFile);
                                                System.out.println("created additional pdf by 1 method");
                                            } catch (Exception e) {
                                                Converter.convertByApache(mDocument, existentFile);
                                                System.out.println("created additional pdf by 2 method");
                                            }
                                            break;
                                        }
                                        counter++;
                                    }
                                }
                            } else {
                                System.out.println("В тексте не найден номер обследования");
                            }
                        } else {
                            System.out.println("В тексте не найдена зона обследования");
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
            }
            else if(contentType != null && contentType.equals(DOC_MIME)){
                System.out.println("founded doc");
                // конвертирую doc в pdf со временным именем
                try{
                    File outputFile = new File(destination, (new RandomString()).nextString() + ".pdf");
                    Converter.convertDocByMsWord(sourceFile, outputFile);
                    System.out.println("Конвертирован .doc");
                    handlePDF(outputFile);
                }
                catch (Exception e){
                    System.out.println("Не удалось конвертировать .doc");
                }

            }
            else if (contentType != null) {
                handlePDF(sourceFile);
            } else {
                System.out.println("Неверный формат файла");
            }

        } catch (Exception e) {
            System.out.println("Have error: " + e.getMessage());
            StringWriter errors = new StringWriter();
            e.printStackTrace(new PrintWriter(errors));
            System.out.println(errors.toString());
        }
    }

    private static void handlePDF(File sourceFile) throws IOException {
        // обработаю PDF
        if (sourceFile.isFile() && destination.isDirectory()) {
            // прочитаю данные из файла
            PDDocument document = PDDocument.load(sourceFile);
            if (!document.isEncrypted()) {
                PDFTextStripper stripper = new PDFTextStripper();
                String text = stripper.getText(document);
                document.close();
                String executionNumber = StringsHandler.getExecutionNumber(text);
                if (executionNumber != null && !executionNumber.isEmpty()) {
                    if (!(new File(destination, executionNumber + ".pdf")).exists()) {
                        System.out.println("Заключение добавлено");
                        // копирую файл в папку
                        boolean result = sourceFile.renameTo(new File(destination, executionNumber + ".pdf"));
                        if (!result) {
                            System.out.println("Не удалось создать файл");
                        } else {
                            System.out.println(executionNumber + ".pdf");
                        }
                    } else {
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
                                        if (!deleteResult) {
                                            System.out.println("error delete target file");
                                        }
                                        // перезапишу имеющийся файл
                                        System.out.println("Файл заключения перезаписан");
                                        boolean result = sourceFile.renameTo(existentFile);
                                        if (!result) {
                                            System.out.println("Не смог создать файл");
                                        } else {
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
                                if (!result) {
                                    System.out.println("Не смог создать файл");
                                } else {
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
}
