package cn.momosv.poi.util;

import cn.momosv.poi.controller.FileController;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Component;
import org.springframework.web.bind.annotation.PathVariable;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Date;

/**
 * @author linshengwen
 * @version 1.0
 * @description
 * @date 2018/9/18 10:37
 **/
@Component
public class FileUtil {

    private static final Logger log = LoggerFactory.getLogger(FileUtil.class);

    public static final String ROOT = "upload-dir";

    private static  ResourceLoader resourceLoader = null;

    @Autowired
    public FileUtil(ResourceLoader resourceLoader) {
        this.resourceLoader = resourceLoader;
    }

    public static void putWBFile(Workbook wb,String path) throws IOException {
        FileOutputStream output=new FileOutputStream(path);
        wb.write(output);
        output.close();
        output.flush();
        //Files.copy(output,  Paths.get(ROOT, "workbook.xls"));
        //输出Excel文件
    }

    public static File newFile(String type,String name) throws IOException {
        String now = XDateUtils.dateToString(new Date(),"yyyyMMdd-HHmm");
        File file0 =new File(ROOT+"/"+type+"/"+now+"/"+name);
        if(file0 .exists() ){
            file0.delete();
        }
        else if  ( !file0 .isDirectory()) {
            file0.getParentFile().mkdirs();
        }
        return file0;
    }

    public static File createFile(String type,String name) throws IOException {
        File newF = newFile(type,name);
        newF.createNewFile();
        return newF;
    }

    public static File createFile(String path) throws IOException {
        File file0 =new File(path);
        if(file0 .exists() ){
            file0.delete();
        }
        else if  ( !file0 .isDirectory()) {
            file0.getParentFile().mkdirs();
        }
        file0.createNewFile();
        return file0;
    }

    public static   ResponseEntity<?> getResponseFile(String path) {
        try {
            return ResponseEntity.ok(getFile(path));
        } catch (Exception e) {
            return ResponseEntity.notFound().build();
        }
    }

    public static Resource getFile(String path) {
        return resourceLoader.getResource("file:" + path );
    }

}
