package cn.momosv.poi.util;


public class Constants {
    public static final String USER_TOKEN = "userToken";
    int MAX_FILE_UPLOAD_SIZE = 5242880;
   //全部-1,0未审批，1通过，2不通过，3已经审批
   public final static Integer USER_ALL_TYPE=-1;
   //0邮箱未验证,1是邮箱认证，2是待审批，3是审批通过，4是不通过
   public final static Integer USER_EMAIL_UN_IDENTIFY=0;
   public final static Integer USER_EMAIL_IDENTIFY=1;
   public final static Integer USER_UN_APPROVED=2;
   public final static Integer USER_PASSED=3;
   public final static Integer USER_UN_PASSED=4;

}
