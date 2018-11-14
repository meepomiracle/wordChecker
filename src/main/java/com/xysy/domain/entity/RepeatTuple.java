package com.xysy.domain.entity;

import org.apache.commons.lang3.StringUtils;

public class RepeatTuple {

    private String no1;

    private String no2;

    public RepeatTuple(String no1, String no2) {
        this.no1 = no1;
        this.no2 = no2;
    }

    public String getNo1() {
        return no1;
    }

    public void setNo1(String no1) {
        this.no1 = no1;
    }

    public String getNo2() {
        return no2;
    }

    public void setNo2(String no2) {
        this.no2 = no2;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        RepeatTuple that = (RepeatTuple) o;
        return StringUtils.equals(that.getNo1(),no1) && StringUtils.equals(that.getNo2(),no2) || StringUtils.equals(that.getNo1(),no2) && StringUtils.equals(that.getNo2(),no1);

    }

    @Override
    public int hashCode() {
        return  no1.hashCode()+no2.hashCode();
    }
}
