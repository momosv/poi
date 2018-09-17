package cn.momosv.poi.util;

public enum DatePattern{
    /** 日期格式枚举类，根据需要添加其他格式 **/
    ISO_SECOND("yyyy-MM-dd'T'HH:mm:ss", "^\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}$"),
    ISO_MINUTE("yyyy-MM-dd'T'HH:mm", "^\\d{4}-\\d{2}-\\d{2}T\\d{2}:\\d{2}$"),
    DATE_TIME("yyyy-MM-dd HH:mm:ss", "^\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2}$"),
    DATE_TIME_FULL("yyyy-MM-dd HH:mm:ss,SSS", "^\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2},\\d{3}$"),
    DATE_ONLY("yyyy-MM-dd", "^\\d{4}-\\d{2}-\\d{2}$"),
    DATE_ONLY2("yyyy/MM/dd", "^\\d{4}/\\d{2}/\\d{2}$"),
    YEAR_MONTH("yyyy-MM", "^\\d{4}-\\d{2}$"),
    DATE_TIME_FULL_NUM("yyyyMMddHHmmssSSS", "^\\d{17}$"),;

    DatePattern(String pattern, String regex){
        this.pattern = pattern;
        this.regex = regex;
    }

    public String getPattern() {
        return pattern;
    }
    public String getRegex() {
        return regex;
    }

    private String pattern;
    private String regex;

    /**
     * 根据日期字符串，判断该日期的格式类型。
     *
     * @param dateStr 日期字符串
     * @return 日期的格式类型，比如getPatternByDateStr("2016-04-27 10:15:08")返回："yyyy-MM-dd HH:mm:ss"
     */
    public static String getPatternByDateStr(String dateStr){
        for(DatePattern df : DatePattern.values()){
            if(RegexUtils.matches(dateStr, df.getRegex())){
                return df.getPattern();
            }
        }
        return null;
    }
}

