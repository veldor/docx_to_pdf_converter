package net.veldor.docx_to_pdf.utils;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

public class Converter {
    public static void convertByMsWord(File sourceFile, File outputFile) throws IOException {
        try (InputStream docxInputStream = new FileInputStream(sourceFile); OutputStream outputStream = new FileOutputStream(outputFile)) {
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
        }
    }

    public static void convertByApache(XWPFDocument source, File existentFile) throws Exception {
        PdfOptions options = PdfOptions.create();
        try (OutputStream out = new FileOutputStream(existentFile)) {
            PdfConverter.getInstance().convert(source, out, options);
        }
    }

    public static void convertDocByMsWord(File sourceFile, File outputFile) throws Exception {
        try (InputStream docxInputStream = new FileInputStream(sourceFile); OutputStream outputStream = new FileOutputStream(outputFile)) {
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.MS_WORD).to(outputStream).as(DocumentType.PDF).execute();
        }
    }
}
