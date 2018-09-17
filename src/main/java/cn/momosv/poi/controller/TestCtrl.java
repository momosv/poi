package cn.momosv.poi.controller;

import cn.momosv.poi.base.baen.Msg;
import cn.momosv.poi.exception.MyException;
import cn.momosv.poi.util.ExcelUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author linshengwen
 * @version 1.0
 * @description
 * @date 2018/9/14 15:35
 **/
@RestController
public class TestCtrl {

    @RequestMapping("/")
    public String hello(){
        return "hello";
    }

    @RequestMapping("/testE")
    public Msg testE(MultipartFile file) throws IOException, InvalidFormatException, MyException {
        System.out.println(file.getOriginalFilename());
       String[] arr = file.getOriginalFilename().split("\\.");
        File rf = File.createTempFile(arr[0],"."+arr[1]);
        file.transferTo(rf);
        try{
            List<String> title = new ArrayList<>();
            List<Map>  map = ExcelUtil.getImportData(rf,title);
            System.out.println("");
        }finally {
            rf.deleteOnExit();
        }
       return  Msg.success();
    }
}
