package com.example.poidemo.util;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.platform.commons.logging.Logger;
import org.junit.platform.commons.logging.LoggerFactory;
import org.junit.platform.commons.util.StringUtils;

import java.io.InputStream;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelUtils {
    private static Logger logger = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * 解析一个workbook.
     *
     * @param workbook
     * @param validateList
     */
    public static Map<String, List<Map<String, String>>> parseExcel(Workbook workbook, String[] validateList) {
        Map<String, List<Map<String, String>>> rowMap = new HashMap<>();
        // 遍历excel表读取每一个sheet.
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            // 验证表头行是否合法
            if (!validateHeader(sheet.getRow(0), validateList)) {
                return rowMap;
            }
            // 数据行
            List<Map<String, String>> rowList = excel2List(sheet, 1);
            if (CollectionUtils.isEmpty(rowList)) {
                System.out.println("\n>>>>>工作表数据为空。");
                return rowMap;
            }
            rowMap.put(String.valueOf(i), rowList);
        }
        return rowMap;
    }

    /**
     * 将excel的行转为list集合.
     *
     * @param sheet
     * @param startRow
     * @return
     */
    public static List<Map<String, String>> excel2List(final Sheet sheet, final int startRow) {
        final List<Map<String, String>> rowList = new ArrayList<>();
        Row row = null;
        for (int i = startRow; i < sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            // 如果一行为空，或者一行没有值，就结束。
            if (null == row || 0 == row.getPhysicalNumberOfCells()) {
                break;
            }
            rowList.add(row2Map(row));
        }
        return rowList;
    }

    /**
     * 行数据转为map
     *
     * @param row 一行数据
     * @return
     */
    private static Map<String, String> row2Map(final Row row) {
        final Map<String, String> map = new HashMap<>();
        for (int j = 0; j < row.getLastCellNum(); ++j) {
            final String value = getCellValue(row.getCell(j));
            map.put(String.valueOf(j), value);
        }
        return map;
    }

    /**
     * 获取单元格值
     *
     * @param cell
     * @return
     */
    private static String getCellValue(final Cell cell) {
        DateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
        DecimalFormat df = new DecimalFormat("0.0000");
        String cellValue = "";
        if (null != cell) {
            switch (cell.getCellType()) {
                // 文本类型
                case STRING:
                    cellValue = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    // 如果是日期
                    if (DateUtil.isCellDateFormatted(cell)) {
                        cellValue = fmt.format(cell.getDateCellValue());
                    } else {
                        // 数字型
                        cellValue = df.format(cell.getNumericCellValue());
                        // 去掉多余的0，如最后一位是.则去掉
                        cellValue = cellValue.replaceAll("0+?$", "").replaceAll("[.]$", "");
                    }
                    break;
                case BOOLEAN: // 布尔型
                    cellValue = String.valueOf(cell.getBooleanCellValue());
                    break;
                case BLANK: // 空白
                    cellValue = cell.getStringCellValue();
                    break;
                case ERROR: // 错误
                    cellValue = "";
                    break;
                case FORMULA: // 公式
                    try {
                        cellValue = String.valueOf(cell.getStringCellValue());
                    } catch (IllegalStateException e) {
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                default:
                    cellValue = cell.getRichStringCellValue() == null ? cellValue
                            : cell.getRichStringCellValue().toString();
            }
        }
        return cellValue;
    }


    /**
     * 数据验证.
     *
     * @param row
     * @param validateList
     */
    public static boolean validateHeader(Row row, String[] validateList) {
        short minColIx = row.getFirstCellNum();
        short maxColIx = row.getLastCellNum();
        for (short colIx = minColIx; colIx < maxColIx; colIx++) {
            String cellVal = row.getCell(colIx).getStringCellValue();
            if (StringUtils.isBlank(cellVal) || !validateList[colIx].equals(cellVal)) {
                System.out.println("员工表头行格式不正确");
                return false;
            }
        }
        return true;
    }

    /**
     * 课程excel
     * @param in
     * @param fileName
     * @return
     * @throws Exception
     */
    public static List getCourseListByExcel(InputStream in, String fileName) throws Exception {
        List list = new ArrayList<>();
        // 创建excel工作簿
        Workbook work = getWorkbook(in, fileName);
        if (null == work) {
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;
        Cell cell = null;
        for (int i = 0; i < work.getNumberOfSheets(); i++) {
            sheet = work.getSheetAt(i);
            if(sheet == null) {
                continue;
            }
            // 滤过第一行标题
            for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
                row = sheet.getRow(j);
                if (row == null || row.getFirstCellNum() == j) {
                    continue;
                }
                List<Object> li = new ArrayList<>();
                for (int y = row.getFirstCellNum(); y < row.getLastCellNum(); y++) {
                    cell = row.getCell(y);
                    // 日期类型转换
                    if(y == 3) {
                        //cell.setCellType(CellType.STRING);
                        double s1 = cell.getNumericCellValue();
                        Date date = HSSFDateUtil.getJavaDate(s1);
                        li.add(date);
                        continue;
                    }
                    li.add(cell);
                }
                list.add(li);
            }
        }
        work.close();
        return list;
    }
    /**
     * 判断文件格式
     * @param in
     * @param fileName
     * @return
     */
    public static Workbook getWorkbook(InputStream in, String fileName) throws Exception {
        Workbook book = null;
        String filetype = fileName.substring(fileName.lastIndexOf("."));
        if(".xls".equals(filetype)) {
            book = new HSSFWorkbook(in);
        } else if (".xlsx".equals(filetype)) {
            book = new XSSFWorkbook(in);
        } else {
            throw new Exception("请上传excel文件！");
        }
        return book;
    }
}
