package oger.util;

import org.apache.commons.codec.binary.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Auther: Oger
 * @Date: 2020-07-21
 * @Description: 导出Excel
 * 设计思想：
 * 1. 样式最简化原则
 * 2. 约定大于规定原则
 * <p>
 * 功能介绍：
 * 1. Controller 每一个接口对应一个demo
 * 2. 可任意创建多个sheet
 * 3. 可在sheet中任意行开始创建table
 * 4. 可按指定顺序导出实体类任意字段
 * - 同一字段在不同table中名称可不同
 * - 可通过Map<String,String> 的方式传入你想导出的字段
 * - 可通过String[] headNames 和 String[] fieldNames 搭配的方式传入你想导出的字段
 * 5. 单sheet单表模式快捷导出Excel
 * 6. 二级树形表头合并的方式创建表
 * 7. 多级表头合并的方式创建表(兼容二级表头合并的方式)
 * 8. 自动设置列宽
 * 9. 无集合属性字段的简单对象表格导出
 * 10. 有集合属性字段的复杂对象表格导出
 * 11. 表格加边框
 */
public class ExcelExportUtil {

    private static Logger logger = LoggerFactory.getLogger(ExcelExportUtil.class);
    private static int DEFAULT_COL_WIDTH = 10;   // 默认列宽
    public static SimpleDateFormat DEFAULT_FORMAT = new SimpleDateFormat("yyyy年MM月dd日");

    /**
     * 自定义模式导出excel： 需自己创建workbook
     *
     * @param fileName
     * @param workbook
     * @param response
     */
    public static void exportExcel(String fileName, HSSFWorkbook workbook, HttpServletResponse response) {
        OutputStream out = null;
        try {
            response.setHeader("Access-Control-Expose-Headers", "Content-Disposition");
            response.setContentType("application/vnd.ms-excel;charset=utf-8");
            String name = URLEncoder.encode(fileName + ".xls", "UTF-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + name + ";filename*=UTF-8''" + name);
            out = response.getOutputStream();
            workbook.write(out);
            out.flush();
        } catch (Exception e) {
            logger.error("导出文件失败", e);
            //  可替换成自己项目中包装的异常类
            throw new RuntimeException("导出文件失败");
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
            } catch (IOException e) {
                logger.error("关闭表格流异常", e);
            }
            try {
                if (out != null) {
                    out.close();
                }
            } catch (IOException e) {
                logger.error("关闭输出流异常", e);
            }
        }

    }

    /**
     * 快捷导出excel: 有sheet标题 有表标题
     *
     * @param fileName
     * @param sheetName
     * @param tableName
     * @param headMap
     * @param dataset
     * @param response
     */
    public static void exportExcel(String fileName, String sheetName, String tableName, Map<String, String> headMap, Collection dataset, HttpServletResponse response) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(sheetName);
        int line = createSheetTitle(headMap.size(), sheetName, sheet, workbook);
        createTable(line, tableName, headMap, dataset, sheet, workbook);
        exportExcel(fileName, workbook, response);
    }

    /**
     * 快捷导出excel: 无sheet标题 无表标题
     *
     * @param fileName
     * @param headMap
     * @param dataset
     * @param response
     */
    public static void exportExcel(String fileName, Map<String, String> headMap, Collection dataset, HttpServletResponse response) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(fileName);
        createTable(0, headMap, dataset, sheet, workbook);
        exportExcel(fileName, workbook, response);
    }

    /**
     * 快捷导出excel: 二级合并表头 无sheet标题 无表标题
     *
     * @param fileName
     * @param dataset
     * @param response
     */
    public static void exportMergeHeadExcel(String fileName, Map<String, Map<String, String>> mergeHeadMap, Collection dataset, HttpServletResponse response) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(fileName);
        createMergeHeadTable(0, mergeHeadMap, dataset, sheet, workbook);
        exportExcel(fileName, workbook, response);
    }

    /**
     * 快捷导出excel: 多级合并表头 无sheet标题 无表标题
     *
     * @param fileName
     * @param dataset
     * @param response
     */
    public static void exportMergeHeadExcel(String fileName, List<Map<String, Object>> mergeHeads, Collection dataset, HttpServletResponse response) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(fileName);
        createMergeHeadTable(0, mergeHeads, dataset, sheet, workbook);
        exportExcel(fileName, workbook, response);
    }

    /**
     * 快捷导出excel:无集合属性字段的简单对象快捷导出excel
     *
     * @param fileName
     * @param names
     * @param t
     * @param response
     */
    public static <T> void exportExcel(String fileName, List<Map<String, Integer>> names, T t, HttpServletResponse response) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(fileName);
        createTable(0, names, t, sheet, workbook);
        exportExcel(fileName, workbook, response);
    }

    /**
     * 创建表格： 有集合属性字段的复杂对象
     *
     * @param line
     * @param names
     * @param t
     * @param sheet
     * @param workbook
     * @param <T>
     * @return
     */
    public static <T> int createTable4Object(int line, List<Map<String, Object>> names, T t, Sheet sheet, HSSFWorkbook workbook) {
        int rows = names.size();
        //计算最大列数
        Integer cols = names.stream().map(map -> {
            int sum = 0;
            for (Map.Entry<String, Object> entry : map.entrySet()) {
                Object value = entry.getValue();
                if (value instanceof Integer) {
                    sum += (Integer) value;
                } else if (value instanceof List) {
                    List<String> fieldNames = (List<String>) value;
                    sum = fieldNames.size();
                }
            }
            return sum;
        }).max(Integer::compareTo).get();
        //创建二维数组
        String[][] cells = new String[rows][cols];
        for (int i = 0; i < rows; i++) {
            cells[i] = new String[cols];
            Map<String, Object> nameMap = names.get(i);
            int index = 0;
            for (Map.Entry<String, Object> entry : nameMap.entrySet()) {
                if (entry.getValue() instanceof Integer) {
                    Integer value = (Integer) entry.getValue();
                    while (value > 0) {
                        cells[i][index] = entry.getKey();
                        value--;
                        index++;
                    }
                }
            }
        }
        CellStyle tableBodyRangeCellStyle = getTableBodyRangeCellStyle(workbook);
        CellStyle tableBodyCellStyle = getTableBodyCellStyle(workbook);
        Row row;
        //创建表
        for (int i = 0; i < rows; i++) {
            row = sheet.createRow(line++);
            Map<String, Object> nameMap = names.get(i);
            int index = 0;
            for (Map.Entry<String, Object> entry : nameMap.entrySet()) {
                String key = entry.getKey();
                if (entry.getValue() instanceof Integer) {  //非集合
                    //创建单元格
                    Integer value = (Integer) entry.getValue();
                    int v = 0;
                    while (v < value) {
                        row.createCell(index + v).setCellStyle(tableBodyCellStyle);
                        v++;
                    }
                    //赋值
                    if (i > 0 && StringUtils.equals(cells[i - 1][index], cells[i][index])) {
                        index += value;
                        continue;
                    }
                    setCellValue(key, t, row.getCell(index), tableBodyCellStyle);
                    //合并单元格
                    int lastRow = i;
                    while (lastRow < rows - 1 && StringUtils.equals(cells[lastRow][index], cells[lastRow + 1][index])) {
                        lastRow++;
                    }
                    if (lastRow > i || value > 1) {
                        sheet.addMergedRegion(new CellRangeAddress(line - 1, line + lastRow - i - 1, index, index + value - 1));//起始行号，终止行号， 起始列号，终止列号
                        row.getCell(index).setCellStyle(tableBodyRangeCellStyle);
                    }
                    index += value;
                } else if (entry.getValue() instanceof List) {  //集合
                    List<String> fieldNames = (List<String>) entry.getValue();
                    try {
                        String getMethodName = "get" + key.substring(0, 1).toUpperCase() + key.substring(1);
                        List dataset = (List) t.getClass().getMethod(getMethodName).invoke(t);
                        for (int n = 0; n < dataset.size(); n++) {
                            Object rowData = dataset.get(n);
                            if (n > 0) {
                                row = sheet.createRow(line++);
                            }
                            for (int m = 0; m < fieldNames.size(); m++) {
                                setCellValue(fieldNames.get(m), rowData, row.createCell(m), tableBodyCellStyle);
                            }
                        }
                    } catch (Exception e) {
                        logger.error("导出文件数据失败", e);
                        //  可替换成自己项目中包装的异常类
                        throw new RuntimeException("导出文件失败");
                    }
                }
            }
        }
        return ++line;
    }

    /**
     * 创建表格： 无集合属性字段的简单对象
     *
     * @param line
     * @param names
     * @param t
     * @param sheet
     * @param <T>
     * @return
     */
    public static <T> int createTable(int line, List<Map<String, Integer>> names, T t, Sheet sheet, HSSFWorkbook workbook) {
        int rows = names.size();
        //计算最大列数
        Integer cols = names.stream().map(map -> {
            int sum = 0;
            for (Map.Entry<String, Integer> entry : map.entrySet()) {
                sum += entry.getValue();
            }
            return sum;
        }).max(Integer::compareTo).get();
        //创建二维数组
        String[][] cells = new String[rows][cols];
        for (int i = 0; i < rows; i++) {
            cells[i] = new String[cols];
            Map<String, Integer> nameMap = names.get(i);
            int index = 0;
            for (Map.Entry<String, Integer> entry : nameMap.entrySet()) {
                Integer value = entry.getValue();
                while (value > 0) {
                    cells[i][index] = entry.getKey();
                    value--;
                    index++;
                }
            }
        }
        CellStyle tableBodyRangeCellStyle = getTableBodyRangeCellStyle(workbook);
        CellStyle tableBodyCellStyle = getTableBodyCellStyle(workbook);
        //创建表
        for (int i = 0; i < rows; i++) {
            Row row = sheet.createRow(line++);
            Map<String, Integer> nameMap = names.get(i);
            int index = 0;
            for (Map.Entry<String, Integer> entry : nameMap.entrySet()) {
                //创建单元格
                Integer value = entry.getValue();
                int v = 0;
                while (v < value) {
                    row.createCell(index + v).setCellStyle(tableBodyCellStyle);
                    v++;
                }
                //赋值
                if (i > 0 && StringUtils.equals(cells[i - 1][index], cells[i][index])) {
                    index += value;
                    continue;
                }
                setCellValue(entry.getKey(), t, row.getCell(index), tableBodyCellStyle);
                //合并单元格
                int lastRow = i;
                while (lastRow < rows - 1 && StringUtils.equals(cells[lastRow][index], cells[lastRow + 1][index])) {
                    lastRow++;
                }
                if (lastRow > i || value > 1) {
                    sheet.addMergedRegion(new CellRangeAddress(line - 1, line + lastRow - i - 1, index, index + value - 1));//起始行号，终止行号， 起始列号，终止列号
                    row.getCell(index).setCellStyle(tableBodyRangeCellStyle);
                }
                index += value;
            }
        }
        return ++line;
    }

    /**
     * 创建表： 多级表头合并 无表标题  从指定行开始  兼容二级表头合并
     *
     * @param line
     * @param mergeHeads： 表头每一行对应一个Map<String,Object>；非底行的key为名称，value为合并单元格数量；底行的key为字段名，value为名称；竖向单元格合并的应在每行Map中都有；Map采用LinkedHashMap
     * @param dataset
     * @param sheet
     * @param workbook
     * @return
     */
    public static int createMergeHeadTable(int line, List<Map<String, Object>> mergeHeads, Collection dataset, Sheet sheet, HSSFWorkbook workbook) {
        CellStyle tableHeadCellStyle = getTableHeadRangeCellStyle(workbook);
        int rows = mergeHeads.size();
        int cols = mergeHeads.get(rows - 1).size();
        String[] fieldNames = new String[cols];
        //创建表头二维数组
        String[][] cells = new String[rows][cols];
        for (int i = 0; i < rows; i++) {
            cells[i] = new String[cols];
            Map<String, Object> mergeHeadMap = mergeHeads.get(i);
            int index = 0;
            for (Map.Entry<String, Object> entry : mergeHeadMap.entrySet()) {
                if (i < rows - 1) {
                    Integer value = Integer.valueOf(entry.getValue().toString());
                    while (value > 0) {
                        cells[i][index] = entry.getKey();
                        value--;
                        index++;
                    }
                } else {
                    cells[i][index] = entry.getValue().toString();
                    fieldNames[index] = entry.getKey();
                    index++;
                }
            }
        }
        //创建表头
        for (int i = 0; i < rows; i++) {
            Row row = sheet.createRow(line++);
            Map<String, Object> mergeHeadMap = mergeHeads.get(i);
            int index = 0;
            for (Map.Entry<String, Object> entry : mergeHeadMap.entrySet()) {
                //创建单元格
                if (i < rows - 1) {
                    Integer value = (Integer) entry.getValue();
                    int v = 0;
                    while (v < value) {
                        row.createCell(index + v).setCellStyle(tableHeadCellStyle);
                        v++;
                    }
                } else {
                    row.createCell(index).setCellStyle(tableHeadCellStyle);
                }
                //赋值并合并单元格
                if (i > 0 && StringUtils.equals(cells[i - 1][index], cells[i][index])) {
                    index++;
                    continue;
                }
                if (i < rows - 1) {
                    Integer value = (Integer) entry.getValue();
                    row.getCell(index).setCellValue(entry.getKey());
                    int lastRow = i;
                    while (lastRow < rows - 1 && StringUtils.equals(cells[lastRow][index], cells[lastRow + 1][index])) {
                        lastRow++;
                    }
                    if (lastRow > i || value > 1) {
                        sheet.addMergedRegion(new CellRangeAddress(line - 1, line + lastRow - i - 1, index, index + value - 1));//起始行号，终止行号， 起始列号，终止列号
                    }
                    index += value;
                } else {
                    row.getCell(index).setCellValue(entry.getValue().toString());
                    index++;
                }
            }
        }
        //创建表体
        return createTableBody(line, fieldNames, dataset, sheet, workbook);
    }

    /**
     * 创建表： 二级树形表头合并 无表标题  从指定行开始
     *
     * @param line
     * @param mergeHeadMap： 外层Map的key为第一行名称，value为子表头Map；里层Map的key为字段名，value为名称；Map采用LinkedHashMap
     * @param dataset
     * @param sheet
     * @param workbook
     * @return
     */
    public static int createMergeHeadTable(int line, Map<String, Map<String, String>> mergeHeadMap, Collection dataset, Sheet sheet, HSSFWorkbook workbook) {
        Row row1 = sheet.createRow(line);
        Row row2 = sheet.createRow(line + 1);
        CellStyle tableHeadCellStyle = getTableHeadRangeCellStyle(workbook);
        int index = 0;
        List<String> fieldNames = new ArrayList<>();
        //创建表头
        for (Map.Entry<String, Map<String, String>> entry : mergeHeadMap.entrySet()) {
            //创建单元格
            String key = entry.getKey();
            Map<String, String> value = entry.getValue();
            int v = 0;
            while (v < value.size()) {
                row1.createCell(index + v).setCellStyle(tableHeadCellStyle);
                row2.createCell(index + v).setCellStyle(tableHeadCellStyle);
                v++;
            }
            //赋值并合并单元格
            if (value.size() < 1) {
                continue;
            }
            row1.getCell(index).setCellValue(key);
            if (value.size() == 1) {
                sheet.addMergedRegion(new CellRangeAddress(line, line + 1, index, index));//起始行号，终止行号， 起始列号，终止列号
                fieldNames.addAll(value.keySet());
                index++;
            } else {
                sheet.addMergedRegion(new CellRangeAddress(line, line, index, index + value.size() - 1));//起始行号，终止行号， 起始列号，终止列号
                for (Map.Entry<String, String> child : value.entrySet()) {
                    row2.getCell(index).setCellValue(child.getValue());
                    fieldNames.add(child.getKey());
                    index++;
                }
            }
        }
        //创建表体
        return createTableBody(line + 2, fieldNames.stream().toArray(String[]::new), dataset, sheet, workbook);
    }

    /**
     * 创建表： 有表标题 从指定行开始
     *
     * @param line      起始行
     * @param tableName
     * @param headMap   要求是LinkedHashMap类型
     * @param dataset
     * @param sheet
     * @param workbook
     * @return 下一行
     */
    public static int createTable(int line, String tableName, Map<String, String> headMap, Collection dataset, Sheet sheet, HSSFWorkbook workbook) {
        String[] fieldNames = new String[headMap.size()];
        String[] headNames = new String[headMap.size()];
        int i = 0;
        for (Map.Entry<String, String> entry : headMap.entrySet()) {
            fieldNames[i] = entry.getKey();
            headNames[i] = entry.getValue();
            i++;
        }
        line = createTableHead(line, tableName, headNames, sheet, workbook);
        return createTableBody(line, fieldNames, dataset, sheet, workbook);
    }

    /**
     * 创建表： 无表标题 从指定行开始
     *
     * @param line     起始行
     * @param headMap  要求是LinkedHashMap类型
     * @param dataset
     * @param sheet
     * @param workbook
     * @return 下一行
     */
    public static int createTable(int line, Map<String, String> headMap, Collection dataset, Sheet sheet, HSSFWorkbook workbook) {
        String[] fieldNames = new String[headMap.size()];
        String[] headNames = new String[headMap.size()];
        int i = 0;
        for (Map.Entry<String, String> entry : headMap.entrySet()) {
            fieldNames[i] = entry.getKey();
            headNames[i] = entry.getValue();
            i++;
        }
        line = createTableHead(line, headNames, sheet, workbook);
        return createTableBody(line, fieldNames, dataset, sheet, workbook);
    }

    /**
     * 创建表头： 有表标题
     *
     * @param line
     * @param tableName
     * @param headNames
     * @param sheet
     * @param workbook
     * @return
     */
    public static int createTableHead(int line, String tableName, String[] headNames, Sheet sheet, HSSFWorkbook workbook) {
        line = createTableTitle(line, tableName, headNames.length, sheet, workbook);
        return createTableHead(line, headNames, sheet, workbook);
    }

    /**
     * 创建表头： 无表标题
     *
     * @param line
     * @param headNames
     * @param sheet
     * @param workbook
     * @return
     */
    public static int createTableHead(int line, String[] headNames, Sheet sheet, HSSFWorkbook workbook) {
        Row head = sheet.createRow(line);
        CellStyle tableHeaderCellStyle = getTableHeadCellStyle(workbook);
        for (int i = 0; i < headNames.length; i++) {
            head.createCell(i).setCellValue(headNames[i]);
            head.getCell(i).setCellStyle(tableHeaderCellStyle);
        }
        return ++line;
    }

    /**
     * 创建表标题
     *
     * @param line
     * @param tableName
     * @param headLength
     * @param sheet
     * @param workbook
     * @return
     */
    public static int createTableTitle(int line, String tableName, int headLength, Sheet sheet, HSSFWorkbook workbook) {
        Row row = sheet.createRow(line);
        Cell cell = row.createCell(0);
        cell.setCellValue(tableName);
        CellStyle tableTitleCellStyle = getTableTitleCellStyle(workbook);
        cell.setCellStyle(tableTitleCellStyle);
        sheet.addMergedRegion(new CellRangeAddress(line, line, 0, headLength - 1));//起始行号，终止行号， 起始列号，终止列号
        return ++line;
    }

    /**
     * 创建表体： fieldNames 必须与 headNames 一一对应
     *
     * @param line       起始行
     * @param sheet
     * @param fieldNames 导出字段名
     * @param dataset
     * @return 下一行
     */
    public static int createTableBody(int line, String[] fieldNames, Collection dataset, Sheet sheet, HSSFWorkbook workbook) {
        Iterator it = dataset.iterator();
        Object rowData;
        Row row;
        int length;
        int[] colLens = new int[fieldNames.length];
//        CellStyle tableBodyCellStyle = getTableBodyCellStyle(workbook);   //表体需要设置边框时可传入setCellValue方法
        while (it.hasNext()) {
            row = sheet.createRow(line++);
            rowData = it.next();
            for (int i = 0; i < fieldNames.length; i++) {
                setCellValue(fieldNames[i], rowData, row.createCell(i), null);
                length = row.getCell(i).getStringCellValue().getBytes().length;
                colLens[i] = length > colLens[i] ? length : colLens[i];
            }
        }
        // 根据数据自动设置列宽
        for (int i = 0, len = colLens.length; i < len; i++) {
            length = colLens[i];
            length = length < DEFAULT_COL_WIDTH ? DEFAULT_COL_WIDTH : length;
            sheet.setColumnWidth(i, length * 256);
        }
        return line + 2;
    }

    private static void setCellValue(String fieldName, Object obj, Cell cell, CellStyle cellStyle) {
        String getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
        try {
            Object value = obj.getClass().getMethod(getMethodName).invoke(obj);
            cell.setCellStyle(cellStyle);
            setCellValue(value, cell);
        } catch (NoSuchMethodException e) {
            cell.setCellValue(fieldName);
        } catch (Exception e) {
            logger.error("导出文件数据失败", e);
            //  可替换成自己项目中包装的异常类
            throw new RuntimeException("导出文件失败");
        }
    }

    private static void setCellValue(Object value, Cell cell) {
        if (value == null) {
            cell.setCellValue("");
        } else if (value instanceof Date) {
            cell.setCellValue(DEFAULT_FORMAT.format((Date) value));
        } else {
            //  需要别的类型可自行扩展
            // 能用toString()直接转string类型的都直接转成string类型
            cell.setCellValue(value.toString());
        }
    }

    /**
     * 根据表头自动设置列宽
     *
     * @param headNames
     * @param sheet
     */
    public static void setColWidth(String[] headNames, Sheet sheet) {
        int length;
        for (int i = 0, len = headNames.length; i < len; i++) {
            length = headNames[i].getBytes().length;
            length = length < DEFAULT_COL_WIDTH ? DEFAULT_COL_WIDTH : length;
            sheet.setColumnWidth(i, length * 256);
        }
    }

    /**
     * 创建sheet标题: 从指定行开始
     *
     * @param length
     * @param sheetName
     * @param sheet
     * @param workbook
     * @return
     */
    public static int createSheetTitle(int length, String sheetName, Sheet sheet, HSSFWorkbook workbook) {
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(sheetName);
        CellStyle sheetHeadCellStyle = getSheetTitleCellStyle(workbook);
        cell.setCellStyle(sheetHeadCellStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, length - 1));//起始行号，终止行号， 起始列号，终止列号
        return 2;
    }

    /**
     * 获取sheet标题单元格样式
     *
     * @param workbook
     * @return
     */
    public static CellStyle getSheetTitleCellStyle(HSSFWorkbook workbook) {
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * 获取表标题单元格样式
     *
     * @param workbook
     * @return
     */
    public static CellStyle getTableTitleCellStyle(HSSFWorkbook workbook) {
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * 获取表头合并单元格样式
     *
     * @param workbook
     * @return
     */
    public static CellStyle getTableHeadRangeCellStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        style.setWrapText(true);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * 获取表头单元格样式
     *
     * @param workbook
     * @return
     */
    public static CellStyle getTableHeadCellStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        style.setWrapText(true);
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * 获取表体合并单元格样式
     *
     * @param workbook
     * @return
     */
    public static CellStyle getTableBodyRangeCellStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        return style;
    }

    /**
     * 获取表体单元格样式
     *
     * @param workbook
     * @return
     */
    public static CellStyle getTableBodyCellStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        return style;
    }
}
