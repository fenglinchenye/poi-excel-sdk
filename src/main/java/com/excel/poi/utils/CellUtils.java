package com.excel.poi.utils;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;

import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 单元格工具类
 */
public class CellUtils {

    /**
     * 设置日期格式的单元数据
     * @param workbook
     * @param date
     * @param index
     * @param row
     */
    public static void setCellDateValue(Workbook workbook, Date date, Integer index, Row row){
        //处理日期格式
        DataFormat format = workbook.createDataFormat();
        short s = format.getFormat("yyyy年MM月dd日 HH时mm分ss秒");
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("华文楷体");
        font.setItalic(true);
        font.setColor(HSSFFont.COLOR_RED);
        style.setFont(font);
        style.setDataFormat(s);

        Cell cell = row.createCell(index);
        cell.setCellStyle(style);
        cell.setCellValue(date);
    }

    /**
     * 描述：对表格中数值进行格式化
     * @param cell
     * @return
     */
    public static Object getCellValue(Cell cell){
        Object value = null;

        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");  //日期格式化
        DecimalFormat df2 = new DecimalFormat("0.00");  //格式化数字

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                value = cell.getRichStringCellValue().getString();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                if("General".equals(cell.getCellStyle().getDataFormatString())){
                    value = cell.getNumericCellValue();
                }else if("m/d/yy".equals(cell.getCellStyle().getDataFormatString())){
                    value = sdf.format(cell.getDateCellValue());
                }else{
                    value = df2.format(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_BLANK:
                value = "";
                break;
            default:
                break;
        }
        return value;
    }
}
