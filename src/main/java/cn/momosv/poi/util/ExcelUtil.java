package cn.momosv.poi.util;

import cn.momosv.poi.exception.MyException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author linshengwen
 * @version 1.0
 * @description
 * @date 2018/9/14 15:52
 **/
public class ExcelUtil {


    public static List<Map> getImportData(File file, List<String> title) throws IOException, InvalidFormatException, MyException {
        if (file.getPath().endsWith("xls")) {
            return dealExcel2003(file, title);
        } else if (file.getPath().endsWith("xlsx")) {
            return dealExcel2007(file, title);
        }
        return null;
    }

    private static List<Map> dealExcel2007(File file, List<String> title) throws IOException, MyException, InvalidFormatException {
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file)); //把一张xlsx的数据表读到wb里
        XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row;
        int totalRow = sheet.getLastRowNum() + 1;
        if (totalRow < 2) {
            throw new MyException("文件内容不能为空");
        }
        //处理title
        row = sheet.getRow(0);
        for (int i = 0; i < row.getLastCellNum(); i++) {
            XSSFCell cell = row.getCell(i);
            switch (cell.getCellType()) {
                case STRING:
                    title.add(cell.getStringCellValue());
                    break;
                default:
                    throw new MyException("标题栏请使用文本格式内容");
            }
        }
        List<Map> list = new ArrayList<>();
        Map<String, Object> map;
        for (int j = 1; j < totalRow; j++) {
            row = sheet.getRow(j);
            map = new HashMap<>();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                XSSFCell cell = row.getCell(i);
                Object obj = null;
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING:
                            obj = cell.getStringCellValue();
                            break;
                        case NUMERIC:
                            try {
                                //判断是否为日期类型
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    //用于转化为日期格式
                                    obj = cell.getDateCellValue();
                                } else {
                                    cell.setCellType(CellType.STRING);
                                    obj = cell.getStringCellValue();
                                    if (obj.toString().contains("."))
                                        obj = Double.valueOf(obj.toString());
                                    else {
                                        obj = Long.valueOf(obj.toString());
                                    }
                                }
                            } catch (Exception e) {
                                obj = "";
                            }
                            break;
                        case FORMULA:
                            try {
                                obj = cell.getStringCellValue();
                            } catch (IllegalStateException e) {
                                //判断是否为日期类型
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    //用于转化为日期格式
                                    obj = cell.getDateCellValue();
                                } else {
                                    cell.setCellType(CellType.STRING);
                                    obj = cell.getStringCellValue();
                                    if (obj.toString().contains("."))
                                        obj = Double.valueOf(obj.toString());
                                    else {
                                        obj = Long.valueOf(obj.toString());
                                    }
                                }
                            }
                            break;
                        case _NONE:
                            obj = "";
                            break;
                        case BLANK:
                            obj = "";
                            break;
                        case BOOLEAN:
                            obj = cell.getBooleanCellValue();
                            break;
                        case ERROR:
                            obj = "";
                        default:
                            obj = "";
                    }
                }
                map.put(title.get(i), obj);
            }
            list.add(map);
        }
 //       wb.close();
   //     FileUtil.putWBFile(wb,FileUtil.ROOT+"/momo.xlsx");
        return list;
    }

    private static List<Map> dealExcel2003(File file, List<String> title) throws IOException, MyException {
        HSSFWorkbook wb = new HSSFWorkbook( new POIFSFileSystem(new FileInputStream(file))); //把一张xls的数据表读到wb里
        HSSFSheet sheet = wb.getSheetAt(0);
        HSSFRow row;
        int totalRow = sheet.getLastRowNum() + 1;
        if (totalRow < 2) {
            throw new MyException("文件内容不能为空");
        }
        //处理title
        row = sheet.getRow(0);
        for (int i = 0; i < row.getLastCellNum(); i++) {
            HSSFCell cell = row.getCell(i);
            switch (cell.getCellType()) {
                case STRING:
                    title.add(cell.getStringCellValue());
                    break;
                default:
                    throw new MyException("标题栏请使用文本格式内容");
            }
        }
        List<Map> list = new ArrayList<>();
        Map<String, Object> map;
        for (int j = 1; j < totalRow; j++) {
            row = sheet.getRow(j);
            map = new HashMap<>();
            for (int i = 0; i < row.getLastCellNum(); i++) {
                HSSFCell cell = row.getCell(i);
                Object obj = null;
                if (cell != null) {
                    switch (cell.getCellType()) {
                        case STRING:
                            obj = cell.getStringCellValue();
                            break;
                        case NUMERIC:
                            try {
                                //判断是否为日期类型
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    //用于转化为日期格式
                                    obj = cell.getDateCellValue();
                                } else {
                                    cell.setCellType(CellType.STRING);
                                    obj = cell.getStringCellValue();
                                    if (obj.toString().contains("."))
                                        obj = Double.valueOf(obj.toString());
                                    else {
                                        obj = Long.valueOf(obj.toString());
                                    }
                                }
                            } catch (Exception e) {
                                obj = "";
                            }
                            break;
                        case FORMULA:
                            try {
                                obj = cell.getStringCellValue();
                            } catch (IllegalStateException e) {
                                //判断是否为日期类型
                                if (DateUtil.isCellDateFormatted(cell)) {
                                    //用于转化为日期格式
                                    obj = cell.getDateCellValue();
                                } else {
                                    cell.setCellType(CellType.STRING);
                                    obj = cell.getStringCellValue();
                                    if (obj.toString().contains("."))
                                        obj = Double.valueOf(obj.toString());
                                    else {
                                        obj = Long.valueOf(obj.toString());
                                    }
                                }
                            }
                            break;
                        case _NONE:
                            obj = "";
                            break;
                        case BLANK:
                            obj = "";
                            break;
                        case BOOLEAN:
                            obj = cell.getBooleanCellValue();
                            break;
                        case ERROR:
                            obj = "";
                        default:
                            obj = "";
                    }
                }
                map.put(title.get(i), obj);

            }
            list.add(map);
        }
        return list;
    }

}
