package com.xysy;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.InputStream;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) {
        String filePath="/pizhu.docx";
        App instance= new App();
        try {
            InputStream is = instance.getClass().getResourceAsStream(filePath);
            XWPFDocument xwpfDocument = new XWPFDocument(is);
            XWPFWordExtractor xwpfWordExtractor = new XWPFWordExtractor(xwpfDocument);

            System.out.println("debug");
        }catch (Exception e){
            e.printStackTrace();
        }

    }
}
