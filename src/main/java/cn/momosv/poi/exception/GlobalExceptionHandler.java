package cn.momosv.poi.exception;


import cn.momosv.poi.base.baen.Msg;
import org.slf4j.LoggerFactory;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.MissingServletRequestParameterException;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.security.auth.login.LoginException;

@ResponseBody
@SuppressWarnings("all")
@ControllerAdvice
public class GlobalExceptionHandler{

    private static final org.slf4j.Logger LOGGER = LoggerFactory.getLogger(GlobalExceptionHandler.class);

    @ExceptionHandler
    private Msg exceptionHandle(Exception e) {
        e.printStackTrace();
        if(e.getMessage().equals("nullUser")){
            return Msg.fail().setCode(-1).add("msg", e.getMessage());
        }
        if (e instanceof LoginException) {
            return Msg.fail().add("msg", e.getMessage());
        }
        if (e instanceof NullPointerException) {
            return Msg.fail().add("msg", e.getMessage());
        }
        if (e instanceof MissingServletRequestParameterException) {
            return Msg.fail().add("msg", e.getMessage());
        }
        if (e instanceof MyException) {
            return Msg.fail(e.getMessage());
        }
        return  Msg.fail().add("msg",StringUtils.isEmpty(e.getMessage())?"系统出现异常":e.getMessage());
    }
}