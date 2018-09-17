package cn.momosv.poi.util;

public enum ExceptionTips {
    /** 日期格式枚举类，根据需要添加其他格式 **/
    name_null("201", "名字不能为空"),
    ;

    ExceptionTips(String code, String desc){
        this.code = code;
        this.desc = desc;
    }

    public String getCode() {
        return code;
    }
    public String getDesc() {
        return desc;
    }

    private String code;
    private String desc;

}
