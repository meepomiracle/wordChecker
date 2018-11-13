package com.xysy.domain.entity;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class XWPFExtendParagraph {

    private String paragraphNo;

    private XWPFParagraph xwpfParagraph;

    private String parentParagraphNo;
    public XWPFExtendParagraph(String paragraphNo, XWPFParagraph xwpfParagraph,String parentParagraphNo) {
        this.paragraphNo = paragraphNo;
        this.xwpfParagraph = xwpfParagraph;
        this.parentParagraphNo=parentParagraphNo;
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
