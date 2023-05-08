package com.example.poidemo.service;

import com.example.poidemo.util.ExcelUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.ResourceUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.List;
import java.util.Map;

@Service
public class IEmployeeServiceImpl implements IEmployeeService{

    @Override
    public String uploadExcel(MultipartFile file) {
        //定义一个数据格式化对象
        XSSFWorkbook wb = null;
        try {
            //excel模板路径
            File cfgFile = ResourceUtils.getFile(ResourceUtils.CLASSPATH_URL_PREFIX + "static/ExcelTemplate/temptest.xlsx");
            InputStream in = new FileInputStream(cfgFile);
            //读取excel模板
            wb = new XSSFWorkbook(in);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        //获取sheet表格，及读取单元格内容
        XSSFSheet sheet = null;
        try{
            sheet = wb.getSheetAt(0);
            //先将获取的单元格设置为String类型，下面使用getStringCellValue获取单元格内容
            //如果不设置为String类型，如果单元格是数字，则报如下异常
            //java.lang.IllegalStateException: Cannot get a STRING value from a NUMERIC cell
            sheet.getRow(2).getCell(2).setCellType(CellType.STRING);
            //读取单元格内容
            String cellValue = sheet.getRow(2).getCell(2).getStringCellValue();

            //添加一行
            XSSFRow row = sheet.createRow(1); //第2行开始写数据
            row.setHeight((short)400); //设置行高
            //向单元格写数据
            row.createCell(1).setCellValue("名称");
        }
        catch (Exception e){
            e.printStackTrace();
        }
        return null;
    }
}
