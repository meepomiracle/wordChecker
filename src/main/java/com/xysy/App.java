package com.xysy;

import com.xysy.util.WordUtil;

import java.io.FileOutputStream;
import java.io.InputStream;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) {
        String filePath="/sample/sample.docx";
        App instance= new App();
        try {
            InputStream is = instance.getClass().getResourceAsStream(filePath);
            XWPFExtendDocument xwpfDocument = new XWPFExtendDocument(is);
            int totalPage=xwpfDocument.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();
            //标题判断
            WordUtil.judgeTitle(xwpfDocument);

            //段落判断
            WordUtil.judgeParagraph(xwpfDocument);
            //输出doc
            xwpfDocument.write(new FileOutputStream("f:/out/wordCheckerOut"+System.currentTimeMillis()+".docx"));
            xwpfDocument.close();
            System.out.println("debug");
        }catch (Exception e){
            e.printStackTrace();
        }

    }






}
