package com.xysy.util;

import com.xysy.XWPFExtendDocument;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTComment;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTComments;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import java.math.BigInteger;
import java.util.Calendar;
import java.util.List;

public class WordUtil {

    public static void judgeParagraph(XWPFExtendDocument xwpfDocument){
        List<XWPFParagraph> paragraphs = xwpfDocument.getParagraphs();
        int requiredCharIndent=2;
        if(CollectionUtils.isNotEmpty(paragraphs)){
            for (int i = 0; i <paragraphs.size() ; i++) {
                XWPFParagraph xwpfParagraph = paragraphs.get(i);
                if(StringUtils.isBlank(xwpfParagraph.getStyle())){//正文样式
                    int firstLineIndent=xwpfParagraph.getFirstLineIndent();//首行缩进
                    int charIndent = firstLineIndent/210;
                    if(charIndent!=requiredCharIndent){//如果不满足首行缩进2字符
                        List<XWPFRun> runs=xwpfParagraph.getRuns();
                        String content=CollectionUtils.isNotEmpty(runs)?StringUtils.join(runs,""):null;
                        if(StringUtils.isNotBlank(content)){
                            System.out.println(content);
                            CTR first = runs.get(0).getCTR();
                            String comment = String.format("首行缩进(要求：%d个字符, 实际：%d字符)",requiredCharIndent,charIndent);
                            WordUtil.addComment(xwpfDocument,first,comment);
                        }else{
                            System.err.println("没有内容,paragraph num:"+i);
                        }

                    }
                }
            }
        }
    }


    public static void addComment(XWPFExtendDocument xwpfExtendDocument,CTR ctr,String comment) {
        //添加批注
        XWPFExtendDocument.MyXWPFCommentsDocument myXWPFCommentsDocument = xwpfExtendDocument.getMyXWPFCommentsDocument();
        CTComments comments = myXWPFCommentsDocument.getComments();
        // 创建绑定ID
        BigInteger cId = xwpfExtendDocument.getCommentId();
        // 创建批注对象
        CTComment ctComment = comments.addNewComment();
        ctComment.setAuthor("自动检查助手");
        ctComment.setInitials("AR");
        ctComment.setDate(Calendar.getInstance());
        // 设置批注内的内容
        ctComment.addNewP().addNewR().addNewT().setStringValue(comment);
        // 将上面创建的绑定ID设置进批注对象
        ctComment.setId(cId);
        ctr.addNewCommentReference().setId(cId);
        xwpfExtendDocument.setCommentId(cId.add(BigInteger.ONE));
    }
}
