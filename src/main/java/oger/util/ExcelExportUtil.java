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
 * 优点：
 * 1. Controller 每一个接口对应一个demo
 * 2. 可任意创建多个sheet
 * 3. 可在任意位置创建table
 * 4. 可按指定顺序导出实体类任意字段
 * - 同一字段在不同table中名称可不同
 * - 可通过Map<String,String> 的方式传入你想导出的字段
 * - 可通过String[] headNames 和 String[] fieldNames 搭配的方式传入你想导出的字段
 * 5. 单sheet单表模式快捷导出Excel
 * 6. 二级树形表头合并的方式创建表
 * 7. 多级表头合并的方式创建表(兼容二级表头合并的方式)
 * 8. 自动设置列宽
 * 9. 简单对象表格导出
 */
public class ExcelExportUtil {

    private static Logger logger = LoggerFactory.getLogger(ExcelExportUtil.class);
    private static int DEFAULT_COL_WIDTH = 10;   // 默认列宽
    public static String DEFAULT_DATE_PATTERN = "yyyy年MM月dd日";//默认日期格式

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
            // TODO 可替换成自己项目中包装的异常类
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
     * 单sheet单表模式快捷导出excel: 有sheet标题 有表标题
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
     * 单sheet单表模式快捷导出excel: 无sheet标题 无表标题
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
     * 单sheet单表模式快捷导出excel: 二级合并表头 无sheet标题 无表标题
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
     * 单sheet单表模式快捷导出excel: 多级合并表头 无sheet标题 无表标题
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
     * 导出简单对象表格： 对象无集合属性字段
     *
     * @param line
     * @param t
     * @param datas
     * @param sheet
     * @param <T>
     * @return
     */
    public static <T> int createSimpleObjectTable(int line, T t, List<Map<String, Integer>> datas, Sheet sheet, HSSFWorkbook workbook) {
        int rows = datas.size();
        Integer cols = datas.stream().map(map -> {
            int sum = 0;
            for (Map.Entry<String, Integer> entry : map.entrySet()) {
                sum += entry.getValue();
            }
            return sum;
        }).max(Integer::compareTo).get();
        //创建表头二维数组
        String[][] cells = new String[rows][cols];
        for (int i = 0; i < rows; i++) {
            cells[i] = new String[cols];
            Map<String, Integer> dataMap = datas.get(i);
            int index = 0;
            for (Map.Entry<String, Integer> entry : dataMap.entrySet()) {
                Integer value = Integer.valueOf(entry.getValue().toString());
                while (value > 0) {
                    cells[i][index] = entry.getKey();
                    value--;
                    index++;
                }
            }
        }
        SimpleDateFormat sdf = new SimpleDateFormat(DEFAULT_DATE_PATTERN);
        CellStyle tableBodyCellStyle = getTableBodyCellStyle(workbook);
        String getMethodName;
        Object fieldValue;
        //合并单元格
        for (int i = 0; i < rows; i++) {
            Row row = sheet.createRow(line + i);
            Map<String, Integer> dataMap = datas.get(i);
            int index = 0;
            for (Map.Entry<String, Integer> entry : dataMap.entrySet()) {
                if (i > 0 && StringUtils.equals(cells[i - 1][index], cells[i][index])) {
                    index++;
                    continue;
                }
                Integer value = entry.getValue();
                String key = entry.getKey();
                getMethodName = "get" + key.substring(0, 1).toUpperCase() + key.substring(1);
                try {
                    fieldValue = t.getClass().getMethod(getMethodName).invoke(t);
                    if (fieldValue == null) {
                        fieldValue = "";
                    } else if (fieldValue instanceof Date) {
                        fieldValue = sdf.format((Date) fieldValue);
                    }
                    row.createCell(index).setCellValue(fieldValue.toString());
                } catch (NoSuchMethodException e) {
                    row.createCell(index).setCellValue(key);
                } catch (Exception e) {
                    logger.error("导出文件失败", e);
                    // TODO 可替换成自己项目中包装的异常类
                    throw new RuntimeException("导出文件失败");
                }
                int lastRow = i;
                while (lastRow < rows - 1 && StringUtils.equals(cells[lastRow][index], cells[lastRow + 1][index])) {
                    lastRow++;
                }
                if (lastRow > i || value > 1) {
                    sheet.addMergedRegion(new CellRangeAddress(line + i, line + lastRow, index, index + value - 1));//起始行号，终止行号， 起始列号，终止列号
                    row.getCell(index).setCellStyle(tableBodyCellStyle);
                }
                index += value;
            }
        }
        return line + rows;
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
        CellStyle tableHeadCellStyle = getTableHeadCellStyle(workbook);
        int rows = mergeHeads.size();
        int cols = mergeHeads.get(rows - 1).size();
        String[] fieldNames = new String[cols];
        //创建表头二维数组
        String[][] heads = new String[rows][cols];
        for (int i = 0; i < rows; i++) {
            heads[i] = new String[cols];
            Map<String, Object> mergeHeadMap = mergeHeads.get(i);
            int index = 0;
            for (Map.Entry<String, Object> entry : mergeHeadMap.entrySet()) {
                if (i < rows - 1) {
                    Integer value = Integer.valueOf(entry.getValue().toString());
                    while (value > 0) {
                        heads[i][index] = entry.getKey();
                        value--;
                        index++;
                    }
                } else {
                    heads[i][index] = entry.getValue().toString();
                    fieldNames[index] = entry.getKey();
                    index++;
                }
            }
        }
        //创建表头
        for (int i = 0; i < rows; i++) {
            Row row = sheet.createRow(line + i);
            Map<String, Object> mergeHeadMap = mergeHeads.get(i);
            int index = 0;
            for (Map.Entry<String, Object> entry : mergeHeadMap.entrySet()) {
                if (i > 0 && StringUtils.equals(heads[i - 1][index], heads[i][index])) {
                    index++;
                    continue;
                }
                if (i < rows - 1) {
                    Integer value = Integer.valueOf(entry.getValue().toString());
                    row.createCell(index).setCellValue(entry.getKey());
                    row.getCell(index).setCellStyle(tableHeadCellStyle);
                    int lastRow = i;
                    while (lastRow < rows - 1 && StringUtils.equals(heads[lastRow][index], heads[lastRow + 1][index])) {
                        lastRow++;
                    }
                    if (lastRow > i || value > 1) {
                        sheet.addMergedRegion(new CellRangeAddress(line + i, line + lastRow, index, index + value - 1));//起始行号，终止行号， 起始列号，终止列号
                    }
                    index += value;
                } else {
                    row.createCell(index).setCellValue(entry.getValue().toString());
                    row.getCell(index).setCellStyle(tableHeadCellStyle);
                    index++;
                }
            }
        }
        //创建表体
        return createTableBody(line + rows, sheet, fieldNames, dataset);
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
        CellStyle tableHeaderCellStyle = getTableHeadCellStyle(workbook);
        int index = 0;
        List<String> fieldNames = new ArrayList<>();
        //创建表头
        for (Map.Entry<String, Map<String, String>> entry : mergeHeadMap.entrySet()) {
            String key = entry.getKey();
            Map<String, String> value = entry.getValue();
            if (value.size() < 1) {
                continue;
            }
            row1.createCell(index).setCellValue(key);
            row1.getCell(index).setCellStyle(tableHeaderCellStyle);
            if (value.size() == 1) {
                sheet.addMergedRegion(new CellRangeAddress(line, line + 1, index, index));//起始行号，终止行号， 起始列号，终止列号
                fieldNames.addAll(value.keySet());
                index++;
            } else {
                sheet.addMergedRegion(new CellRangeAddress(line, line, index, index + value.size() - 1));//起始行号，终止行号， 起始列号，终止列号
                for (Map.Entry<String, String> child : value.entrySet()) {
                    row2.createCell(index).setCellValue(child.getValue());
                    row2.getCell(index).setCellStyle(tableHeaderCellStyle);
                    fieldNames.add(child.getKey());
                    index++;
                }
            }
        }
        //创建表体
        return createTableBody(line + 2, sheet, fieldNames.stream().toArray(String[]::new), dataset);
    }

    /**
     * 创建表： 有表标题
     *
     * @param tableName
     * @param headMap
     * @param dataset
     * @param sheet
     * @param workbook
     * @return
     */
    public static int createTable(String tableName, Map<String, String> headMap, Collection dataset, Sheet sheet, HSSFWorkbook workbook) {
        return createTable(0, tableName, headMap, dataset, sheet, workbook);
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
        return createTableBody(line, sheet, fieldNames, dataset);
    }

    /**
     * 创建表： 无表标题
     *
     * @param headMap
     * @param dataset
     * @param sheet
     * @param workbook
     * @return
     */
    public static int createTable(Map<String, String> headMap, Collection dataset, Sheet sheet, HSSFWorkbook workbook) {
        return createTable(0, headMap, dataset, sheet, workbook);
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
        return createTableBody(line, sheet, fieldNames, dataset);
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
    public static int createTableBody(int line, Sheet sheet, String[] fieldNames, Collection dataset) {
        Iterator it = dataset.iterator();
        Object rowData;
        String fieldName;
        String getMethodName;
        Row row;
        Cell cell;
        Object value;
        int length;
        int[] headLens = new int[fieldNames.length];
        SimpleDateFormat sdf = new SimpleDateFormat(DEFAULT_DATE_PATTERN);
        while (it.hasNext()) {
            row = sheet.createRow(line);
            line++;
            rowData = it.next();
            for (int i = 0; i < fieldNames.length; i++) {
                fieldName = fieldNames[i];
                getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                try {
                    value = rowData.getClass().getMethod(getMethodName).invoke(rowData);
                    cell = row.createCell(i);
                    if (value == null) {
                        cell.setCellValue("");
                    } else if (value instanceof Date) {
                        cell.setCellValue(sdf.format((Date) value));
                    } else {
                        // TODO 需要别的类型可自行扩展
                        // 能用toString()直接转string类型的都直接转成string类型
                        cell.setCellValue(value.toString());
                    }
                    length = cell.getStringCellValue().getBytes().length;
                    headLens[i] = length > headLens[i] ? length : headLens[i];
                } catch (Exception e) {
                    logger.error("导出文件数据失败", e);
                    // TODO 可替换成自己项目中包装的异常类
                    throw new RuntimeException("导出文件失败");
                }
            }
        }
        // 根据数据自动设置列宽
        for (int i = 0, len = headLens.length; i < len; i++) {
            length = headLens[i];
            length = length < DEFAULT_COL_WIDTH ? DEFAULT_COL_WIDTH : length;
            sheet.setColumnWidth(i, length * 256);
        }
        return line + 2;
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
     * 获取表体单元格样式
     *
     * @param workbook
     * @return
     */
    public static CellStyle getTableBodyCellStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        return style;
    }
}
