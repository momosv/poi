package cn.momosv.poi.util;//package cn.momosv.websocket.util;//package cn.momosv.util;
//
//import cn.momosv.hos.model.TbBaseUserPO;
//import org.springframework.beans.factory.annotation.Autowired;
//import org.springframework.stereotype.Component;
//
//import javax.servlet.http.HttpServletRequest;
//import javax.servlet.http.HttpSession;
//import java.util.UUID;
//
//
//@Component
//public class SysUtil {
//	 public static final String USER = "user";
//	 public static final String BASE_PATH = "basePath";
//
//	 @Autowired
//	 private static HttpServletRequest request;
//
//	@Autowired
//	private static HttpSession session;
//	    //① 获取保存在Session中的用户对象
//	  public static TbBaseUserPO getSessionUser() throws Exception {
//		  if(null!=session.getAttribute(USER)) {
//			  TbBaseUserPO userBasePO = (TbBaseUserPO) session.getAttribute(USER);
//			  if (userBasePO == null) {
//				  //throw new Exception("请登录再使用系统");
//			  }
//			  return userBasePO;
//		  }
//		return null;
//	    }
//
//	     //②将用户对象保存到Session中
//	    public static void setSessionUser(TbBaseUserPO TbUserBasePO) {
//			request.getSession().setAttribute(USER,TbUserBasePO);
//	    }
//	  //②将用户对象保存到Session中
//	    public static void removeSessionUser() {
//	       if (request.getSession().getAttribute(USER) != null) {
//			   request.getSession().removeAttribute(USER);// 将用户信息移除session
//			}
//	    }
//
//	    //③ 获取基于应用程序的url绝对路径
//	    public static final String getBasePath(HttpServletRequest request) {
//	        return (String)request.getSession().getAttribute(BASE_PATH);
//	    }
//	    //③ 获取基于应用程序的url绝对路径
//	    public static final String getBasePath(HttpSession s) {
//	        return (String)s.getAttribute(BASE_PATH);
//	    }
//	  //③ 设置基于应用程序的url绝对路径
//	    public static final void setBasePath(HttpServletRequest request, String url) {
//	    	 request.getSession().setAttribute(BASE_PATH,url);
//	    }
//	    //③ 设置基于应用程序的url绝对路径
//	    public static final void setBasePath(HttpSession s, String url) {
//			s.setAttribute(BASE_PATH,url);
//	    }
//
//	public static String UUID36(){
//		return UUID.randomUUID().toString();
//	}
//}
