package net.veldor.docx_to_pdf;

import net.veldor.docx_to_pdf.utils.Parser;

public class Main {
    public static void main(String[] args) {
        if(args.length == 2){
            // агрументы переданы, передам их парсеру
            Parser.parse(args[0], args[1]);
        }
        else{
            System.out.println("Нужно передать имя файла и адрес папки назначения");
        }
    }
}
