package com.xysy.util;

import com.xysy.XWPFExtendDocument;
import com.xysy.domain.constants.Constants;
import com.xysy.domain.enums.AlignmentEnum;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;
import java.util.Calendar;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WordUtil {

    public static void judgeParagraph(XWPFExtendDocument xwpfDocument) {

        List<XWPFParagraph> paragraphs = xwpfDocument.getParagraphs();
        //doc样式
        XWPFStyles xwpfStyles = xwpfDocument.getStyles();
        //默认正文样式id
        String defaultStyleId = "a";
        //默认正文样式
        XWPFStyle defaultMainStyle = xwpfStyles.getStyle(defaultStyleId);
        XWPFStyle xwpfStyle = null;
        //测试 要求缩进
        int requiredCharIndent = 2;
        //要求对齐 左对齐
        AlignmentEnum requiredAlignment = AlignmentEnum.left;
        //要求行距,20磅
        double requiredLineSpacing = 20.0;
        if (CollectionUtils.isNotEmpty(paragraphs)) {
            for (int i = 0; i < paragraphs.size(); i++) {
                StringBuilder sb = new StringBuilder();
                XWPFParagraph xwpfParagraph = paragraphs.get(i);

                String styleId = xwpfParagraph.getStyleID();
                if (StringUtils.isBlank(styleId)) {
                    styleId = defaultStyleId;
                }
                if (StringUtils.isBlank(styleId)) {//如果查不到段落样式，默认使用正文样式
                    xwpfStyle = defaultMainStyle;
                } else {
                    xwpfStyle = xwpfStyles.getStyle(styleId);
                }
                //根据样式获取段落首行缩进
                int firstLineIndent = getFirstLineIndentByStyle(xwpfStyle);
//                    int firstLineIndent = xwpfParagraph.getFirstLineIndent();//首行缩进
                //段落对齐方式
                AlignmentEnum alignment = getParagraphAlignmentByStyle(xwpfStyle);
                //行距
                double lineSpacing = getLineSpaceByStyle(xwpfStyle);
                List<XWPFRun> runs = xwpfParagraph.getRuns();
                String content = CollectionUtils.isNotEmpty(runs) ? StringUtils.join(runs, "") : null;
                if (StringUtils.isNotEmpty(content)) {
                    System.out.println(content);
                    String comment = null;
                    //检查缩进
                    if (styleId.equals(defaultStyleId)) {//仅对正文检查缩进
                        if (firstLineIndent != requiredCharIndent) {//如果不满足首行缩进2字符
                            comment = String.format("首行缩进(要求：%d个字符, 实际：%d字符)", requiredCharIndent, firstLineIndent);
                            sb.append(comment);

                        }
                    }

                    //检查对齐方式
                    if (styleId.equals(defaultStyleId)) {//仅对正文检查对齐
                        if (!alignment.equals(requiredAlignment)) {
                            comment = String.format("对齐(要求：%s, 实际:%s)", requiredAlignment.getDesc(), alignment.getDesc());
                            sb.append(comment);
                        }
                    }

                    //检查行距
                    if (styleId.equals(defaultStyleId)) {
                        if (requiredLineSpacing != lineSpacing) {
                            comment = String.format("行距值(要求：%.1f磅, 实际：%.1f磅)", requiredLineSpacing, lineSpacing);
                            sb.append(comment);
                        }
                    }

                    comment = sb.toString();

                    CTR first = runs.get(0).getCTR();
                    if (StringUtils.isNotBlank(comment)) {
                        WordUtil.addComment(xwpfDocument, first, comment);
                    }

                    //段落内容判断
                    for(XWPFRun run:runs){
                        judgeRun(run,xwpfStyle, xwpfDocument);
                    }

                } else {
                    System.err.println("没有内容,paragraph num:" + i);
                }
            }
        }
    }

    /**
     * 判断区域文本是否和段落一致
     *
     * @param run
     */
    public static void judgeRun(XWPFRun run, XWPFStyle xwpfStyle,XWPFExtendDocument xwpfDocument) {
        //段落字体样式
        String pChineseFontType = getChineseFontType(xwpfStyle);
        String pWesternFontType = getWesternFontType(xwpfStyle);
        int pFontSize = getFontSize(xwpfStyle);

        String rChineseFontType = getChineseFontType(run);
        String rWesternFontType = getWesternFontType(run);
        int rFontSize = getFontSize(run);

        StringBuilder sb = new StringBuilder();
        String comment=null;
        if(!pChineseFontType.equals(rChineseFontType)){
            comment=String.format("(与段落字体样式不一致，段落字体样式:%s,实际样式为:%s)",pChineseFontType,rChineseFontType);
            sb.append(comment);
        }

        if(!pWesternFontType.equals(rWesternFontType)){
            comment=String.format("(与段落字体样式不一致，段落字体样式:%s,实际样式为:%s)",pWesternFontType,rWesternFontType);
            sb.append(comment);
        }

        if(pFontSize!=rFontSize){
            comment=String.format("(与段落字体大小不一致)");
            sb.append(comment);
        }
        comment=sb.toString();
        CTR ctr = run.getCTR();
        if(StringUtils.isNotBlank(comment)){
            addComment(xwpfDocument,ctr,comment);
        }

    }

    public static String getChineseFontType(XWPFRun run) {
        String fontType = Constants.DEFAULT_CHINESE_FONT;
        CTR ctr = run.getCTR();
        if (ctr != null) {
            CTRPr ctrPr = ctr.getRPr();
            if (ctrPr != null) {
                CTFonts ctFonts = ctrPr.getRFonts();
                if (ctFonts != null) {
                    fontType = ctFonts.getEastAsia();
                }
            }
        }
        return fontType;
    }

    public static String getWesternFontType(XWPFRun run) {
        String fontType = Constants.DEFAULT_WEST_FONT;
        CTR ctr = run.getCTR();
        if (ctr != null) {
            CTRPr ctrPr = ctr.getRPr();
            if (ctrPr != null) {
                CTFonts ctFonts = ctrPr.getRFonts();
                if (ctFonts != null) {
                    fontType = ctFonts.getAscii();
                }
            }
        }
        return fontType;
    }

    public static String getChineseFontType(XWPFStyle xwpfStyle) {
        String fontType = Constants.DEFAULT_CHINESE_FONT;
        CTStyle ctStyle = xwpfStyle.getCTStyle();
        if (ctStyle != null) {
            CTRPr ctrPr = ctStyle.getRPr();
            if (ctrPr != null) {
                CTFonts ctFonts = ctrPr.getRFonts();
                if (ctFonts != null) {
                    fontType = ctFonts.getEastAsia();
                }
            }
        }
        return fontType;
    }

    public static String getWesternFontType(XWPFStyle xwpfStyle) {
        String fontType = Constants.DEFAULT_WEST_FONT;
        CTStyle ctStyle = xwpfStyle.getCTStyle();
        if (ctStyle != null) {
            CTRPr ctrPr = ctStyle.getRPr();
            if (ctrPr != null) {
                CTFonts ctFonts = ctrPr.getRFonts();
                if (ctFonts != null) {
                    fontType = ctFonts.getAscii();
                }
            }
        }
        return fontType;
    }


    public static int getFontSize(XWPFRun run) {
        BigInteger fontSize = Constants.DEFAULT_FONT_SIZE;
        CTR ctr = run.getCTR();
        CTRPr ctrPr = ctr.getRPr();
        if (ctrPr != null) {
            CTHpsMeasure ctHpsMeasure = ctrPr.getSz();
            if (ctHpsMeasure != null) {
                fontSize = ctHpsMeasure.getVal();
            }
        }
        return fontSize.intValue();
    }

    public static int getFontSize(XWPFStyle xwpfStyle) {
        BigInteger fontSize = Constants.DEFAULT_FONT_SIZE;
        CTStyle ctStyle = xwpfStyle.getCTStyle();
        if (ctStyle != null) {
            CTRPr ctrPr = ctStyle.getRPr();
            if (ctrPr != null) {
                CTHpsMeasure ctHpsMeasure = ctrPr.getSz();
                if (ctHpsMeasure != null) {
                    fontSize = ctHpsMeasure.getVal();
                }
            }
        }
        return fontSize.intValue();
    }

    public static void judgeTitle(XWPFExtendDocument xwpfDocument) {
        List<XWPFParagraph> paragraphs = xwpfDocument.getParagraphs();
        //doc样式
        XWPFStyles xwpfStyles = xwpfDocument.getStyles();
        XWPFParagraph titleParagraph = null;
        if (CollectionUtils.isNotEmpty(paragraphs)) {
            for (XWPFParagraph p : paragraphs) {
                String title = p.getText();
                if (StringUtils.isNotBlank(title)) {
                    titleParagraph = p;
                }
            }
        }
        String title = titleParagraph.getText();
        String year = extractNumber(title);
        year = year.substring(0,year.length()>=4?4:year.length());
        int currentYear = Calendar.getInstance().get(Calendar.YEAR);
        if (StringUtils.isBlank(year) || (StringUtils.isNotBlank(year) && !year.equals(String.valueOf(currentYear)))) {
            List<XWPFRun> runs = titleParagraph.getRuns();
            if (CollectionUtils.isNotEmpty(runs)) {
                for (XWPFRun run : runs) {
                    String text = run.text();
                    if (isNumber(text)) {
                        CTR ctr = run.getCTR();
                        String comment = String.format("试卷标题时间与当前时间不符,标题时间:%s,当前时间:%d", year, currentYear);
                        addComment(xwpfDocument, ctr, comment);
                        break;
                    }
                }
            }

        }
    }

    /**
     * 判断是否是数字（包含小数）
     *
     * @param target
     * @return
     */
    public static boolean isNumber(String target) {
        String regx = "-?[0-9]+.*[0-9]*";
        return regxMatch(target, regx);
    }

    public static boolean regxMatch(String target, String regx) {
        Pattern pattern = Pattern.compile(regx);
        Matcher matcher = pattern.matcher(target);
        return matcher.matches();
    }

    public static String extractNumber(String text) {
        String rex = "[^0-9]";
        Pattern p = Pattern.compile(rex);
        Matcher m = p.matcher(text);
        String num = m.replaceAll("").trim();
        return num;
    }

    public static int getFirstLineIndentByStyle(XWPFStyle xwpfStyle) {
        try {
            CTStyle ctStyle = Optional.ofNullable(xwpfStyle.getCTStyle()).get();
            CTPPr ctpPr = Optional.ofNullable(ctStyle.getPPr()).get();
            CTInd ctInd = Optional.ofNullable(ctpPr.getInd()).get();
            int firstLineChars = Optional.ofNullable(ctInd.getFirstLineChars()).orElse(BigInteger.ZERO).intValue() / 100;
            return firstLineChars;
        } catch (NoSuchElementException e) {
            return 0;
        }
    }

    public static double getLineSpaceByStyle(XWPFStyle xwpfStyle) {
        CTStyle ctStyle = Optional.ofNullable(xwpfStyle.getCTStyle()).get();
        CTPPr ctpPr = Optional.ofNullable(ctStyle.getPPr()).get();
        CTSpacing ctSpacing = Optional.ofNullable(ctpPr.getSpacing()).get();
        BigInteger line = Optional.ofNullable(ctSpacing.getLine()).orElse(BigInteger.valueOf(180L));
        //单倍行距默认18磅
        double space = line.intValue() / (double) Constants.POUND_UNIT;
        return space;
    }

    /**
     * 根据样式获取段落对齐方式
     * 默认返回两边对齐
     *
     * @param xwpfStyle
     * @return
     */
    public static AlignmentEnum getParagraphAlignmentByStyle(XWPFStyle xwpfStyle) {
        STJc.Enum jc = null;
        try {
            CTStyle ctStyle = Optional.ofNullable(xwpfStyle.getCTStyle()).get();
            CTPPr ctpPr = Optional.ofNullable(ctStyle.getPPr()).get();
            CTJc ctJc = Optional.ofNullable(ctpPr.getJc()).get();
            jc = Optional.ofNullable(ctJc.getVal()).orElse(STJc.BOTH);
        } catch (NoSuchElementException e) {
            jc = STJc.BOTH;
        }
        AlignmentEnum alignmentEnum = AlignmentEnum.convert(jc);
        return alignmentEnum;
    }

    public static void addComment(XWPFExtendDocument xwpfExtendDocument, CTR ctr, String comment) {
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
