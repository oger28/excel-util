package oger.util;

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
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;

/**
 * @Auther: Oger
 * @Date: 2020-07-21
 * @Description: 导出Excel
 * 设计思想：
 * 1. 样式最简化原则
 * 2. 约定大于规定原则
 * <p>
 * 优点：
 * 1. 可任意创建sheet
 * 2. 可在任意位置创建table
 * 3. 可按顺序导出实体类任意字段:
 *      可通过Map<String,String> 的方式传入你想导出的字段
 *      亦可通过String[] headNames 和 String[] fieldNames 搭配的方式传入你想导出的字段
 * 4. 亦可一次性导出单sheet单表模式的Excel
 */
public class ExcelExportUtil {

    private static Logger logger = LoggerFactory.getLogger(ExcelExportUtil.class);

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
            //可替换成自己项目中包装的异常类
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

    public static void exportExcel(String fileName, Map<String, String> headMap, Collection dataset, HttpServletResponse response) {
        exportExcel(fileName, fileName, fileName, headMap, dataset, response);
    }

    public static void exportExcel(String fileName, String sheetName, String tableName, Map<String, String> headMap, Collection dataset, HttpServletResponse response) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet(sheetName);
        ExcelExportUtil.createTable(0, tableName, headMap, dataset, sheet, workbook);
        ExcelExportUtil.exportExcel(fileName, workbook, response);
    }

    public static int createTable(String tableName, Map<String, String> headMap, Collection dataset, Sheet sheet, HSSFWorkbook workbook) {
        return createTable(0, tableName, headMap, dataset, sheet, workbook);
    }

    /**
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

    public static int createTableHead(int line, String tableName, String[] headNames, Sheet sheet, HSSFWorkbook workbook) {
        Row row = sheet.createRow(line);
        Cell cell = row.createCell(0);
        cell.setCellValue(tableName);
        CellStyle tableTitleCellStyle = ExcelExportUtil.getTableTitleCellStyle(workbook);
        cell.setCellStyle(tableTitleCellStyle);
        sheet.addMergedRegion(new CellRangeAddress(line, line, 0, headNames.length - 1));//起始行号，终止行号， 起始列号，终止列号
        line++;
        Row head1 = sheet.createRow(line);
        CellStyle tableHeaderCellStyle = ExcelExportUtil.getTableHeadCellStyle(workbook);
        for (int i = 0; i < headNames.length; i++) {
            head1.createCell(i).setCellValue(headNames[i]);
            head1.getCell(i).setCellStyle(tableHeaderCellStyle);
        }
        return ++line;
    }

    /**
     * fieldNames 必须与 headNames 一一对应
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
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
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
                    } else if (value instanceof Float) {
                        cell.setCellValue(String.valueOf(value));
                    } else if (value instanceof BigDecimal) {
                        cell.setCellValue(((BigDecimal) value).doubleValue());
                    } else if (value instanceof Date) {
                        cell.setCellValue(sdf.format((Date) value));
                    } else {
                        //需要别的类型用的时候自己扩展
                        cell.setCellValue(value.toString());
                    }
                } catch (Exception e) {
                    logger.error("导出文件数据失败", e);
                    //可替换成自己项目中包装的异常类
                    throw new RuntimeException("导出文件失败");
                }
            }
        }
        return ++line;
    }

    public static int createSheetTitle(String fileName, Sheet sheet, HSSFWorkbook workbook) {
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(fileName);
        CellStyle sheetHeadCellStyle = ExcelExportUtil.getSheetHeadCellStyle(workbook);
        cell.setCellStyle(sheetHeadCellStyle);
        //标题 2行 22列 可视具体情况修改，亦可重构方法传参动态设置列数
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 22));//起始行号，终止行号， 起始列号，终止列号
        return 2;
    }

    public static int createSheetTitle(int line, String fileName, Sheet sheet, HSSFWorkbook workbook) {
        Row row = sheet.createRow(line);
        Cell cell = row.createCell(0);
        cell.setCellValue(fileName);
        CellStyle sheetHeadCellStyle = ExcelExportUtil.getSheetHeadCellStyle(workbook);
        cell.setCellStyle(sheetHeadCellStyle);
        //标题 2行 22列 可视具体情况修改，亦可重构方法传参动态设置列数
        sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 22));//起始行号，终止行号， 起始列号，终止列号
        return line + 2;
    }

    /*  请不要修改样式，尽量使用默认样式  */

    private static CellStyle getSheetHeadCellStyle(HSSFWorkbook workbook) {
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    private static CellStyle getTableTitleCellStyle(HSSFWorkbook workbook) {
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

    private static CellStyle getTableHeadCellStyle(HSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        style.setWrapText(true);
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
    }
}
