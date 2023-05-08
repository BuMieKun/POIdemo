package com.example.poidemo.controller;

import com.example.poidemo.service.IEmployeeService;
import com.example.poidemo.util.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletRequest;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

@RestController
@RequestMapping("exe/employee")
public class EmployeeController {

    @Autowired
    private IEmployeeService iEmployeeService;

    @PostMapping("/excel")
    public String parseAndSaveExcel(@RequestPart(value = "attachFile", required = false) MultipartFile file) throws Exception {
        Workbook workbook = null;
        String filetype = file.getOriginalFilename().substring(file.getOriginalFilename().lastIndexOf("."));
        if(".xls".equals(filetype)) {
            workbook = new HSSFWorkbook(file.getInputStream());
        } else if (".xlsx".equals(filetype)) {
            workbook = new XSSFWorkbook(file.getInputStream());
        } else {
            throw new Exception("请上传excel文件！");
        }
        //获取一共有多少sheet，然后遍历
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            //获取sheet中一共有多少行，遍历行（注意第一行是标题）
            int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
            for (int j = 0; j < physicalNumberOfRows; j++) {
                //获取每一行有多少单元格，遍历单元格
                Row row = sheet.getRow(j);
                int physicalNumberOfCells = row.getPhysicalNumberOfCells();
                for (int k = 0; k < physicalNumberOfCells; k++) {
                    Cell cell = row.getCell(k);
                    cell.setCellType(CellType.STRING);
                    String cellValue = cell.getStringCellValue();
                }
            }
        }
        workbook.close();
        return "success";
    }

}
