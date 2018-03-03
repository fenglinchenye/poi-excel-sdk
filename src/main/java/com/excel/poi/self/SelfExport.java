package com.excel.poi.self;

import com.excel.poi.enums.EnumDataStatusModel;
import com.excel.poi.utils.EnumConstantsUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static com.excel.poi.utils.GenernalFieldValueByFields.getFieldValue;


/**
 * 自定义导出
 */
public class SelfExport {

    /**
     * 映射 导出的列名与实体数据中的属性名之间的对应关系
     * k:列名   v: 属性名
     */
    public static Map<String,String> map;

    public static Map<String, String> getMap() {
        return map;
    }

    /**
     *  导出操作
     * @param sourceList  源数据
     * @param fields 导出的列名
     * @param response 响应流
     * @param sheetName 工作页名
     * @param modelEnumClass  必须实现接口EnumDataModel的枚举类(枚举类中有对应的属性)
     * @param <T>
     */
    public static <T> void export(List<T> sourceList, String[] fields, HttpServletResponse response, String sheetName, Class<? extends EnumDataStatusModel> modelEnumClass){

        try{
            if(fields!=null){
                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = workbook.createSheet(sheetName);
                XSSFRow row = sheet.createRow(0);
                //创建标题行
                for (int i = 0; i < fields.length; i++) {
                    sheet.setColumnWidth(i, 5000);
                    row.createCell(i).setCellValue(fields[i]);
                }

                //处理数据行
                for (int i = 0; i < sourceList.size(); i++) {
                    T t = sourceList.get(i);
                    row = sheet.createRow(i+1);
                    for (int j = 0; j < fields.length; j++) {

                        Object obj = getFieldValue(map.get(fields[j]), t);

                        //对日期类型进行筛选。做日期类型的处理
                        if(obj instanceof Date){

                            setCellDateValue(workbook,(Date)obj,j,row);

                        }else{
                            XSSFCell cell = row.createCell(j);
                            if(obj == null){
                                cell.setCellValue("");
                            }else if("status".equals(map.get(fields[j]))){

                                int status = (Integer) obj;
                                if (modelEnumClass==null){
                                    cell.setCellValue(status);
                                }else{
                                    String str_status = EnumConstantsUtil.valueBy(modelEnumClass, status);
                                    cell.setCellValue(str_status);
                                }
                            }else{
                                cell.setCellValue(obj.toString());
                            }
                        }
                    }
                }
                //设置响应输出
                response.setContentType("application/vnd.ms-excel");
                response.setHeader("content-disposition", "attachment;filename="+URLEncoder.encode(sheetName+"记录表.xlsx","utf-8"));
                ServletOutputStream out = response.getOutputStream();
                workbook.write(out);
                out.close();
            }
        }catch(Exception e){
            e.printStackTrace();
        }
    }

    /**
     * 设置日期格式的单元数据
     * @param workbook
     * @param date
     * @param index
     * @param row
     */
    private static void setCellDateValue(Workbook workbook, Date date, Integer index, Row row){

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

}
