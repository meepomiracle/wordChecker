package com.xysy.domain.enums;

import com.xysy.util.CommonUtil;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;

public enum AlignmentEnum {
    left(1,"左对齐"),
    center(2,"居中"),
    right(3,"右对齐"),
    both(4,"两端对齐"),
    mediumKashida(5),
    distribute(6),
    numTab(7),
    highKashida(8),
    lowKashida(9),
    thaiDistribute(10);

    private int code;

    private String desc;

    AlignmentEnum(int code, String desc) {
        this.code = code;
        this.desc = desc;
    }

    AlignmentEnum(int code) {
        this.code = code;
    }

    public int getCode() {
        return code;
    }

    public void setCode(int code) {
        this.code = code;
    }

    public String getDesc() {
        return desc;
    }

    public void setDesc(String desc) {
        this.desc = desc;
    }

    public static AlignmentEnum convert(STJc.Enum anEnum){
        AlignmentEnum alignmentEnum = AlignmentEnum.valueOf(anEnum.toString());
        return alignmentEnum;
    }


    public static AlignmentEnum convert(ParagraphAlignment paragraphAlignment){
        AlignmentEnum alignmentEnum = AlignmentEnum.valueOf(CommonUtil.UnderlineToHump(paragraphAlignment.toString()));
        return alignmentEnum;
    }
    public static void main(String[] args) {
        STJc.Enum anEnum = STJc.BOTH;
        AlignmentEnum alignmentEnum = convert(anEnum);
        System.out.println(alignmentEnum.desc);
    }
}
