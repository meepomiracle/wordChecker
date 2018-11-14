package com.xysy;

import com.xysy.domain.entity.XWPFExtendDocument;
import com.xysy.util.WordUtil;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) {
        if(args.length<2){
            System.err.println("请输入word文件路径和输出word目录，仅支持.docx格式");
            System.exit(0);
        }

        String inputPath = args[0];
        String outPath=args[1];
        File out =new File(outPath);
        if(!out.exists()){
            out.mkdirs();
        }

        String filePath="/sample/sample.docx";
        App instance= new App();
        try {
//            InputStream is = instance.getClass().getResourceAsStream(filePath);
            InputStream is = new FileInputStream(inputPath);
            XWPFExtendDocument xwpfDocument = new XWPFExtendDocument(is);
            XWPFWordExtractor xwpfWordExtractor = new XWPFWordExtractor(xwpfDocument);
            int totalPages = xwpfWordExtractor.getExtendedProperties().getPages();
            int wordCount=xwpfWordExtractor.getExtendedProperties().getCharacters();

            //标题判断
            WordUtil.judgeTitle(xwpfDocument);
            //段落判断
            WordUtil.judgeParagraph(xwpfDocument);
            //最前面增加文章统计
            XmlCursor xmlCursor=xwpfDocument.getParagraphs().get(0).getCTP().newCursor();
            xmlCursor.toPrevSibling();
            XWPFParagraph p=xwpfDocument.insertNewParagraph(xmlCursor);
            XWPFRun r = p.createRun();
            r.setText(String.format("全文页数：%d 页           ",totalPages));
            r.setColor("3A5FCD");
            r=p.createRun();
            r.setText(String.format("字数统计：%d",wordCount));
            r.setColor("3A5FCD");

            xmlCursor= xwpfDocument.getParagraphs().get(0).getCTP().newCursor();
            xmlCursor.toNextSibling();
            p=xwpfDocument.insertNewParagraph(xmlCursor);
            r = p.createRun();
            r.setText("----------------------------------------------------------------");
            //输出doc
            xwpfDocument.write(new FileOutputStream(outPath+"/out"+System.currentTimeMillis()+".docx"));
            xwpfDocument.close();
        }catch (Exception e){
            e.printStackTrace();
        }

    }






}
