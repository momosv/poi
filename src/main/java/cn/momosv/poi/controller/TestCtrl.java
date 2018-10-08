package cn.momosv.poi.controller;

import cn.momosv.poi.base.baen.Msg;
import cn.momosv.poi.exception.MyException;
import cn.momosv.poi.util.ExcelUtil;
import cn.momosv.poi.util.FileUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ResourceLoader;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.nio.file.*;
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











    public static final String ROOT = "upload-dir";

    private final ResourceLoader resourceLoader;

    @Autowired
    public TestCtrl(ResourceLoader resourceLoader) {
        this.resourceLoader = resourceLoader;
    }

    @RequestMapping("/")
    public String hello(){
        return "hello";
    }

    @RequestMapping("/testE")
    public Msg testE(MultipartFile file) throws IOException, InvalidFormatException, MyException {
        File rf = FileUtil.newFile(file.getContentType().split("/")[0],file.getOriginalFilename());
       if( !Files.isDirectory(Paths.get(file.getContentType().split("/")[0]))) {
           Files.createDirectories(Paths.get(file.getContentType().split("/")[0]));
       }
        Path file1 = Files.createFile(Paths.get(file.getContentType().split("/")[0], file.getOriginalFilename()));
        Files.copy(file.getInputStream(),  Paths.get(rf.getParent(), rf.getName()));

        try{
            List<String> title = new ArrayList<>();
           List<Map>  map = ExcelUtil.getImportData(rf,title);
        }finally {
            Files.deleteIfExists(Paths.get(rf.getPath()));
        }

       return  Msg.success();
    }

    @RequestMapping( value = "up")
    private void testExcel(HttpServletResponse response) throws IOException {
        //创建HSSFWorkbook对象(excel的文档对象)
        XSSFWorkbook wb = new XSSFWorkbook();
//建立新的sheet对象（excel的表单）
        XSSFSheet sheet=wb.createSheet("成绩表");
//在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
        XSSFRow row1=sheet.createRow(0);
//创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个
        XSSFCell cell=row1.createCell(0);
        //设置单元格内容
        cell.setCellValue("学员考试成绩一览表");
//合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
        sheet.addMergedRegion(new CellRangeAddress(0,0,0,3));
//在sheet里创建第二行
        XSSFRow row2=sheet.createRow(1);
        //创建单元格并设置单元格内容
        row2.createCell(0).setCellValue("姓名");
        row2.createCell(1).setCellValue("班级");
        row2.createCell(2).setCellValue("笔试成绩");
        row2.createCell(3).setCellValue("机试成绩");
        //在sheet里创建第三行
        XSSFRow row3=sheet.createRow(2);
        row3.createCell(0).setCellValue("李明");
        row3.createCell(1).setCellValue("As178");
        row3.createCell(2).setCellValue(87);
        row3.createCell(3).setCellValue(78);
        //.....省略部分代码
        FileUtil.putWBFile( wb, ROOT+"/workbook.xlsx");
//        FileOutputStream output=new FileOutputStream(ROOT+"/workbook.xlsx");
//        wb.write(output);
//        output.close();
        //Files.copy(output,  Paths.get(ROOT, "workbook.xls"));
        //输出Excel文件
    }
}
