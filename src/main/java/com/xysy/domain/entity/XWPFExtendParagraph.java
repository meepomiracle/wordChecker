package com.xysy.domain.entity;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class XWPFExtendParagraph {

    private String paragraphNo;

    private int level;

    private XWPFParagraph xwpfParagraph;

    private String parentParagraphNo;
    public XWPFExtendParagraph(String paragraphNo, XWPFParagraph xwpfParagraph,String parentParagraphNo) {
        this.paragraphNo = paragraphNo;
        this.xwpfParagraph = xwpfParagraph;
        this.parentParagraphNo=parentParagraphNo;
    }

    public XWPFExtendParagraph(String paragraphNo, int level, XWPFParagraph xwpfParagraph, String parentParagraphNo) {
        this.paragraphNo = paragraphNo;
        this.level = level;
        this.xwpfParagraph = xwpfParagraph;
        this.parentParagraphNo = parentParagraphNo;
    }

    public int getLevel() {
        return level;
    }

    public void setLevel(int level) {
        this.level = level;
    }

    public String getParagraphNo() {
        return paragraphNo;
    }

    public void setParagraphNo(String paragraphNo) {
        this.paragraphNo = paragraphNo;
    }

    public XWPFParagraph getXwpfParagraph() {
        return xwpfParagraph;
    }

    public void setXwpfParagraph(XWPFParagraph xwpfParagraph) {
        this.xwpfParagraph = xwpfParagraph;
    }

    public String getParentParagraphNo() {
        return parentParagraphNo;
    }

    public void setParentParagraphNo(String parentParagraphNo) {
        this.parentParagraphNo = parentParagraphNo;
    }
}
