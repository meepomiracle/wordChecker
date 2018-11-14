package com.xysy.util;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class RegxUtil {
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

    public static String regxExtract(String target,String regx){
        Pattern pattern = Pattern.compile(regx);
        Matcher m = pattern.matcher(target);
        while(m.find()){
            return m.group();
        }
        return null;
    }

    public static String extractNumber(String text) {
        String rex = "[^0-9]";
        Pattern p = Pattern.compile(rex);
        Matcher m = p.matcher(text);
        String num = m.replaceAll("").trim();
        return num;
    }

    public static String extractCharacter(String text){
        String rex = "[^a-zA-Z]";
        Pattern p = Pattern.compile(rex);
        Matcher m = p.matcher(text);
        String num = m.replaceAll("").trim();
        return num;
    }

    public static String relace(String regx, String text) {
        Pattern p = Pattern.compile(regx);
        Matcher m = p.matcher(text);
        String num = m.replaceAll("").trim();
        return num;
    }
}
