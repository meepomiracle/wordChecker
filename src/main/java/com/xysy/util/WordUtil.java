package com.xysy.util;

import com.google.common.collect.Lists;
import com.xysy.domain.constants.Constants;
import com.xysy.domain.entity.RepeatTuple;
import com.xysy.domain.entity.XWPFExtendDocument;
import com.xysy.domain.entity.XWPFExtendParagraph;
import com.xysy.domain.enums.AlignmentEnum;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.math.BigInteger;
import java.util.*;
import java.util.stream.Collectors;

import static com.xysy.util.RegxUtil.extractNumber;
import static com.xysy.util.RegxUtil.isNumber;

public class WordUtil {

    public static void judgeParagraph(XWPFExtendDocument xwpfDocument) {

        List<XWPFExtendParagraph> level1PNos = Lists.newArrayList();
        List<XWPFExtendParagraph> level2PNos = Lists.newArrayList();
        List<XWPFExtendParagraph> choiceParagraphs=Lists.newArrayList();
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
                collectParagraphNo(paragraphs,i,xwpfParagraph, level1PNos, level2PNos,choiceParagraphs);
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
                int firstLineIndent = getFirstLineIndentByStyle(xwpfStyle,xwpfParagraph);
//                    int firstLineIndent = xwpfParagraph.getFirstLineIndent();//首行缩进
                //段落对齐方式
                AlignmentEnum alignment = getParagraphAlignmentByStyle(xwpfStyle,xwpfParagraph);
                //行距
                double lineSpacing = getLineSpaceByStyle(xwpfStyle,xwpfParagraph);
                List<XWPFRun> runs = xwpfParagraph.getRuns();
                String content = CollectionUtils.isNotEmpty(runs) ? StringUtils.join(runs, "") : null;
                if (StringUtils.isNotEmpty(content)) {
//                    System.out.println(content);
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
                        if (!requiredAlignment.equals(alignment)) {
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
                    for (XWPFRun run : runs) {
                        judgeRun(run, xwpfStyle, xwpfDocument);
                    }

                } else {
//                    System.err.println("没有内容,paragraph num:" + i);
                }
            }

            judgeParagraphNo(xwpfDocument, level1PNos, level2PNos);
            judgeChoiceNo(xwpfDocument, choiceParagraphs);
        }
    }

    private static void judgeChoiceNo(XWPFExtendDocument xwpfDocument,List<XWPFExtendParagraph> choiceNos){
        if(CollectionUtils.isNotEmpty(choiceNos)){
            Map<String, List<XWPFExtendParagraph>> map = choiceNos.stream().collect(Collectors.groupingBy(XWPFExtendParagraph::getParentParagraphNo));
            Iterator<Map.Entry<String, List<XWPFExtendParagraph>>> iterator = map.entrySet().iterator();
            while (iterator.hasNext()) {
                Map.Entry<String, List<XWPFExtendParagraph>> entry = iterator.next();
                List<XWPFExtendParagraph> groupChoiceNos = entry.getValue();
                judgeChoiceNoOfGroup(xwpfDocument, groupChoiceNos);
            }
        }
    }

    private static void judgeChoiceNoOfGroup(XWPFExtendDocument xwpfDocument,List<XWPFExtendParagraph> choiceNos){
        List<RepeatTuple> repeatTuples = Lists.newArrayList();

        if (CollectionUtils.isNotEmpty(choiceNos)) {
            for (int i = 0; i < choiceNos.size(); i++) {
                XWPFExtendParagraph p = choiceNos.get(i);
                String pNo = p.getParagraphNo().toLowerCase();
                CTR ctr = p.getXwpfParagraph().getRuns().get(0).getCTR();
                //判断选项编号是否连续
                if (i > 0) {
                    String prePNo = choiceNos.get(i - 1).getParagraphNo().toLowerCase();
                    char choiceNo = pNo.charAt(0);
                    char preChoiceNo=prePNo.charAt(0);
                    if (choiceNo - preChoiceNo != 1) {
                        String comment = String.format("当前选项编号与上一个选项编号不连续或重复，当前选项编号:(%s),上一选项编号:(%s)", pNo, prePNo);
                        addComment(xwpfDocument, ctr, comment);
                    }
                }
                //判断选项列表中是否有重复内容
                RepeatTuple repeatTuple = judgeChoiceDuplicate(p, choiceNos);
                if(repeatTuple!=null){
                    if(!repeatTuples.contains(repeatTuple)){
                        repeatTuples.add(repeatTuple);
                        String comment = String.format("当前选项内容与其他选项内容重复，当前选项编号:(%s),重复选项编号:(%s)",repeatTuple.getNo1(),repeatTuple.getNo2());
                        addComment(xwpfDocument, ctr,comment);
                    }
                }
            }
        }
    }



    private static void judgeParagraphNo(XWPFExtendDocument xwpfDocument, List<XWPFExtendParagraph> level1PNos, List<XWPFExtendParagraph> level2PNos) {
        judgeParagraphNo(xwpfDocument, level1PNos);
        if (CollectionUtils.isNotEmpty(level2PNos)) {
            Map<String, List<XWPFExtendParagraph>> map = level2PNos.stream().collect(Collectors.groupingBy(XWPFExtendParagraph::getParentParagraphNo));
            Iterator<Map.Entry<String, List<XWPFExtendParagraph>>> iterator = map.entrySet().iterator();
            while (iterator.hasNext()) {
                Map.Entry<String, List<XWPFExtendParagraph>> entry = iterator.next();
                List<XWPFExtendParagraph> groupLevel2PNos = entry.getValue();
                judgeParagraphNo(xwpfDocument, groupLevel2PNos);
            }
        }
    }

    private static void judgeParagraphNo(XWPFExtendDocument xwpfDocument, List<XWPFExtendParagraph> pNos) {
        List<RepeatTuple> repeatTuples = Lists.newArrayList();

        if (CollectionUtils.isNotEmpty(pNos)) {
            for (int i = 0; i < pNos.size(); i++) {
                XWPFExtendParagraph p = pNos.get(i);
                String pNo = p.getParagraphNo();
                String pNoN = RegxUtil.extractNumber(pNo);
                CTR ctr = p.getXwpfParagraph().getRuns().get(0).getCTR();
                //判断段落编号是否连续
                if (i > 0) {
                    String prePNo = pNos.get(i - 1).getParagraphNo();
                    String prePNoN = RegxUtil.extractNumber(prePNo);
                    if (Integer.valueOf(pNoN).intValue() - Integer.valueOf(prePNoN).intValue() != 1) {

                        String comment = String.format("当前标题编号与上一个标题编号不连续或重复，当前编号:(%s),上一编号:(%s)", pNo, prePNo);
                        addComment(xwpfDocument, ctr, comment);
                    }
                }
                //判断标题列表中是否有重复内容
                RepeatTuple repeatTuple = judgeDuplicate(p, pNos);
                if(repeatTuple!=null){
                    if(!repeatTuples.contains(repeatTuple)){
                        repeatTuples.add(repeatTuple);
                        String comment = String.format("当前标题内容与其他标题内容重复，当前标题编号:(%s),重复标题编号:(%s)",repeatTuple.getNo1(),repeatTuple.getNo2());
                        addComment(xwpfDocument, ctr,comment);
                    }
                }
            }
        }
    }

    /**
     * 判断选项内容是否有重复
     * @param p
     * @param choiceNos
     * @return
     */
    private static RepeatTuple judgeChoiceDuplicate(XWPFExtendParagraph p, List<XWPFExtendParagraph> choiceNos) {
        String text = p.getXwpfParagraph().getText();
        String regx = "[a-fA-F](\\.)";
        //提取标题外的内容
        String content = RegxUtil.relace(regx,text);

        for(XWPFExtendParagraph tp:choiceNos){
            if(StringUtils.equals(p.getParagraphNo(),tp.getParagraphNo())){//排除自己
                continue;
            }
            String tContent=RegxUtil.relace(regx,tp.getXwpfParagraph().getText());
            if(StringUtils.equals(content,tContent)){
                return new RepeatTuple(p.getParagraphNo(),tp.getParagraphNo());
            }
        }

        return null;
    }
    /**
     * 判断段落内容是否重复
     * @param p
     * @param pNos
     * @return
     */
    private static RepeatTuple judgeDuplicate(XWPFExtendParagraph p,List<XWPFExtendParagraph> pNos){
        String text = p.getXwpfParagraph().getText();
        String level1Regx = "^\\d(\\.)";
        String level2Regx = "^\\d(\\.)\\d";
        int level = p.getLevel();
        String regx = level1Regx;
        if(level==1){
            regx = level1Regx;
        }
        if(level==2){
            regx=level2Regx;
        }
        //提取标题外的内容
        String content = RegxUtil.relace(regx,text);

        for(XWPFExtendParagraph tp:pNos){
            if(StringUtils.equals(p.getParagraphNo(),tp.getParagraphNo())){//排除自己
                continue;
            }
            String tContent=RegxUtil.relace(regx,tp.getXwpfParagraph().getText());
            if(StringUtils.equals(content,tContent)){
                return new RepeatTuple(p.getParagraphNo(),tp.getParagraphNo());
            }
        }

        return null;
    }


    public static void collectParagraphNo(List<XWPFParagraph> paragraphs,int pos,XWPFParagraph paragraph, List<XWPFExtendParagraph> level1PNos, List<XWPFExtendParagraph> level2PNos,List<XWPFExtendParagraph> choiceParagraphs) {
        String defaultLevel = "top";
        String text = paragraph.getText()!=null?CommonUtil.trim(paragraph.getText()):null;
        if (StringUtils.isBlank(text)) {
            return;
        }
        //判断段落是否以编号开头,如1.1,1.2等等
        String level1Regx = "^\\d(\\.)[^0-9]";
        String level2Regx = "^\\d(\\.)\\d";

        //判断段落是否以选择题选项开头,如A,B,C,D;
        String choiceRegx="[a-fA-F](\\.)";

//        if(RegxUtil.regxMatch(text,level1Regx)){
        String paragraphNo = RegxUtil.regxExtract(text, level1Regx);
        if (StringUtils.isNotBlank(paragraphNo)) {
            paragraphNo = RegxUtil.extractNumber(paragraphNo);
            level1PNos.add(new XWPFExtendParagraph(paragraphNo, 1,paragraph, defaultLevel));
            return;
        }

//        }

//        if(RegxUtil.regxMatch(text,level2Regx )){
        paragraphNo = RegxUtil.regxExtract(text, level2Regx);
        if (StringUtils.isNotBlank(paragraphNo)) {
            //二级标题的上一级标题一般来说已经存在,且对应一级标题列表最后一个
            String parentPNo = defaultLevel;
            if (CollectionUtils.isNotEmpty(level1PNos)) {
                parentPNo = level1PNos.get(level1PNos.size() - 1).getParagraphNo();
            }
            level2PNos.add(new XWPFExtendParagraph(paragraphNo,2, paragraph, parentPNo));
        }

//        }

        String choiceNo=RegxUtil.regxExtract(text, choiceRegx);
        if(StringUtils.isNotBlank(choiceNo)){
            choiceNo = RegxUtil.extractCharacter(choiceNo);
            //理论上来说选择题一定有父标题，不会在顶级
            String parentPNo = defaultLevel;
            //寻找该选择题编号的最近一个父标题
            if(pos>0){
                int index = pos-1;
                XWPFParagraph p =paragraphs.get(index);
                String pText = CommonUtil.trim(p.getText());
                while(index-->0){
                    //先寻找二级标题
                    parentPNo = RegxUtil.regxExtract(pText, level2Regx);
                    if(StringUtils.isNotBlank(parentPNo)){
                        choiceParagraphs.add(new XWPFExtendParagraph(choiceNo,3,paragraph, parentPNo));
                        break;
                    }else{//找不到再寻找一级标题
                        parentPNo = RegxUtil.regxExtract(pText,level1Regx);
                        if(StringUtils.isNotBlank(parentPNo)){
                            parentPNo = RegxUtil.extractNumber(parentPNo);
                            choiceParagraphs.add(new XWPFExtendParagraph(choiceNo,2,paragraph, parentPNo));
                            break;
                        }
                    }
                    p=paragraphs.get(index);
                    pText = CommonUtil.trim(p.getText());
                }
            }

        }
    }

    /**
     * 判断区域文本是否和段落一致
     *
     * @param run
     */
    public static void judgeRun(XWPFRun run, XWPFStyle xwpfStyle, XWPFExtendDocument xwpfDocument) {
        //段落字体样式
        String pChineseFontType = getChineseFontType(xwpfStyle);
        String pWesternFontType = getWesternFontType(xwpfStyle);
        BigInteger pFontSize = getFontSize(xwpfStyle);

        String rChineseFontType = getChineseFontType(run);
        String rWesternFontType = getWesternFontType(run);
        BigInteger rFontSize = getFontSize(run);

        StringBuilder sb = new StringBuilder();
        String comment = null;
        if (rChineseFontType != null && !pChineseFontType.equals(rChineseFontType)) {
            comment = String.format("(与段落字体样式不一致，段落字体样式:%s,实际样式为:%s)", pChineseFontType, rChineseFontType);
            sb.append(comment);
        }

        if (rWesternFontType != null && !pWesternFontType.equals(rWesternFontType)) {
            comment = String.format("(与段落字体样式不一致，段落字体样式:%s,实际样式为:%s)", pWesternFontType, rWesternFontType);
            sb.append(comment);
        }

        if (rFontSize != null && pFontSize != null && !pFontSize.equals(rFontSize)) {
            comment = String.format("(与段落字体大小不一致)");
            sb.append(comment);
        }
        comment = sb.toString();
        CTR ctr = run.getCTR();
        if (StringUtils.isNotBlank(comment)) {
            addComment(xwpfDocument, ctr, comment);
        }

    }

    public static String getChineseFontType(XWPFRun run) {
//        String fontType = Constants.DEFAULT_CHINESE_FONT;
        String fontType = null;
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
//        String fontType = Constants.DEFAULT_WEST_FONT;
        String fontType = null;
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
                    fontType = ctFonts.getEastAsia() != null ? ctFonts.getEastAsia() : fontType;
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
                    fontType = ctFonts.getAscii() != null ? ctFonts.getAscii() : fontType;
                }
            }
        }
        return fontType;
    }


    public static BigInteger getFontSize(XWPFRun run) {
//        BigInteger fontSize = Constants.DEFAULT_FONT_SIZE;
        BigInteger fontSize = null;
        CTR ctr = run.getCTR();
        CTRPr ctrPr = ctr.getRPr();
        if (ctrPr != null) {
            CTHpsMeasure ctHpsMeasure = ctrPr.getSz();
            if (ctHpsMeasure != null) {
                fontSize = ctHpsMeasure.getVal();
            }
        }
        return fontSize;
    }

    public static BigInteger getFontSize(XWPFStyle xwpfStyle) {
//        BigInteger fontSize = Constants.DEFAULT_FONT_SIZE;
        BigInteger fontSize = null;
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
        return fontSize;
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
                    break;
                }
            }
        }
        String title = titleParagraph.getText();
        String year = extractNumber(title);
        year = year.substring(0, year.length() >= 4 ? 4 : year.length());
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

        //页眉判断
        List<XWPFHeader> headers=xwpfDocument.getHeaderList();
        if(CollectionUtils.isNotEmpty(headers)){
            for(XWPFHeader xwpfHeader:headers){
                String headContent=xwpfHeader.getText();
                if(StringUtils.isNotBlank(headContent)){
                    title=CommonUtil.trim(title);
                    headContent=CommonUtil.trim(headContent);
                    if(!title.equals(headContent)){
                        //最前面增加提示
                        XmlCursor xmlCursor=xwpfDocument.getParagraphs().get(0).getCTP().newCursor();
                        xmlCursor.toPrevSibling();
                        XWPFParagraph p=xwpfDocument.insertNewParagraph(xmlCursor);
                        XWPFRun r = p.createRun();
                        r.setText(String.format("试卷标题与页眉不一致"));
                        r.setColor("FF0000");
                        xmlCursor= xwpfDocument.getParagraphs().get(0).getCTP().newCursor();
                        xmlCursor.toNextSibling();
                        p=xwpfDocument.insertNewParagraph(xmlCursor);
                        r = p.createRun();
                        r.setText("----------------------------------------------------------------");
                    }
                }
            }
        }
    }


    public static int getFirstLineIndentByStyle(XWPFStyle xwpfStyle, XWPFParagraph xwpfParagraph) {
        CTP ctp=xwpfParagraph.getCTP();
        if(ctp!=null){
            CTPPr ctpPr=ctp.getPPr();
            if(ctpPr!=null){
                CTInd ctInd=ctpPr.getInd();
                if(ctInd!=null){
                    BigInteger firstLineChars=ctInd.getFirstLineChars();
                    if(firstLineChars!=null){
                        return firstLineChars.intValue()/100;
                    }
                }
            }
        }
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

    public static double getLineSpaceByStyle(XWPFStyle xwpfStyle, XWPFParagraph xwpfParagraph) {
        CTP ctp=xwpfParagraph.getCTP();
        if(ctp!=null){
            CTPPr ctpPr=ctp.getPPr();
            if(ctpPr!=null){
                CTSpacing ctsPacing=ctpPr.getSpacing();
                if(ctsPacing!=null){
                    BigInteger ctsPacingLine=ctsPacing.getLine();
                    if(ctsPacingLine!=null){
                        return ctsPacingLine.intValue()/ (double) Constants.POUND_UNIT;
                    }
                }
            }
        }

        try {
            CTStyle ctStyle = Optional.ofNullable(xwpfStyle.getCTStyle()).get();
            CTPPr ctpPr = Optional.ofNullable(ctStyle.getPPr()).get();
            CTSpacing ctSpacing = Optional.ofNullable(ctpPr.getSpacing()).get();
            BigInteger line = Optional.ofNullable(ctSpacing.getLine()).orElse(BigInteger.valueOf(360L));
            //单倍行距默认18磅
            double space = line.intValue() / (double) Constants.POUND_UNIT;
            return space;
        }catch (NoSuchElementException e){
            return BigInteger.valueOf(360L).intValue()/(double) Constants.POUND_UNIT;
        }

    }

    /**
     * 根据样式获取段落对齐方式
     * 默认返回两边对齐
     *
     * @param xwpfStyle
     * @param xwpfParagraph
     * @return
     */
    public static AlignmentEnum getParagraphAlignmentByStyle(XWPFStyle xwpfStyle, XWPFParagraph xwpfParagraph) {
        STJc.Enum jc = null;
        CTP ctp=xwpfParagraph.getCTP();
        if(ctp!=null){
            CTPPr ctpPr=ctp.getPPr();
            if(ctpPr!=null){
                CTJc ctJc=ctpPr.getJc();
                if(ctJc!=null){
                    jc=ctJc.getVal();
                    if(jc!=null){
                        return AlignmentEnum.convert(jc);
                    }
                }
            }
        }

        try {
            CTStyle ctStyle = Optional.ofNullable(xwpfStyle.getCTStyle()).get();
            CTPPr ctpPr = Optional.ofNullable(ctStyle.getPPr()).get();
            CTJc ctJc = Optional.ofNullable(ctpPr.getJc()).get();
            jc = Optional.ofNullable(ctJc.getVal()).orElse(STJc.BOTH);
        } catch (NoSuchElementException e) {
            ParagraphAlignment paragraphAlignment=xwpfParagraph.getAlignment();
            if(paragraphAlignment!=null){
                return AlignmentEnum.convert(paragraphAlignment);
            }else{
                jc = STJc.BOTH;
            }
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
