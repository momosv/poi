package cn.momosv.poi.controller;

import cn.momosv.poi.base.baen.Msg;
import cn.momosv.poi.util.DatePattern;
import cn.momosv.poi.util.XDateUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ResourceLoader;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Date;

@CrossOrigin("*")
@RestController
@RequestMapping("upload")
public class FileController {

    private static final Logger log = LoggerFactory.getLogger(FileController.class);

    public static final String ROOT = "upload-dir";

    private final ResourceLoader resourceLoader;

    @Autowired
    public FileController(ResourceLoader resourceLoader) {
        this.resourceLoader = resourceLoader;
    }

    //显示图片的方法关键 匹配路径像 localhost:8080/b7c76eb3-5a67-4d41-ae5c-1642af3f8746.png
    @RequestMapping(value = "/{type}/{upLoadUser}/{filename:.+}")
    public ResponseEntity<?> getFile(@PathVariable String filename,@PathVariable String type,@PathVariable String upLoadUser) {
       String path = Paths.get(ROOT+"/"+type+"/"+upLoadUser, filename).toString();
        try {
            return ResponseEntity.ok(resourceLoader.getResource("file:" + path ));
        } catch (Exception e) {
            return ResponseEntity.notFound().build();
        }

    }
    @RequestMapping( value = "up/{upLoadUser}")
    public Object handleFileUpload(@RequestParam("file") MultipartFile file,@PathVariable @RequestParam(defaultValue = "default") String type,@PathVariable @RequestParam(defaultValue = "default")String upLoadUser,
                                   RedirectAttributes redirectAttributes, HttpServletRequest request) throws IOException {
      //  Files.delete(Paths.get(ROOT, "momo.jpg"));
        String fileName="";
        testExcel(null);

        final File htmlFile = File.createTempFile("temp", ".html");//创建临时文件
        log.info("临时文件所在的本地路径：" + htmlFile.getCanonicalPath());
        FileOutputStream fos = new FileOutputStream(htmlFile);
        try {
            //这里处理业务逻辑
        } finally {
            //关闭临时文件
            fos.flush();
            fos.close();

            htmlFile.delete();//程序退出时删除临时文件
        }

        if (!file.isEmpty()) {
            type = file.getContentType().split("/")[0];
            fileName= XDateUtils.dateToString(new Date(), DatePattern.DATE_TIME_FULL_NUM.getPattern())+file.getOriginalFilename();
            File file0 =new File(ROOT+"/"+type+"/"+upLoadUser);
            if  (!file0 .exists()  && !file0 .isDirectory()){
               log.info("路径:"+ROOT+"/"+type+"/"+upLoadUser+" 不存在");
                file0 .mkdirs();
            }
            Files.copy(file.getInputStream(),  Paths.get(ROOT+"/"+type+"/"+upLoadUser, fileName));
            return Msg.success().add("url","/upload/"+type+"/"+upLoadUser+"/"+fileName);
        }
        return Msg.fail("上传的文件为空");
    }

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

        FileOutputStream output=new FileOutputStream(ROOT+"/workbook.xlsx");
        wb.write(output);
        output.close();
        //Files.copy(output,  Paths.get(ROOT, "workbook.xls"));
        //输出Excel文件
    }


}
