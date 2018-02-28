package com.excel.poi.self.export;

import com.excel.poi.enums.EnumDataModel;
import com.excel.poi.utils.EnumConstantsUtil;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.xssf.usermodel.*;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.net.URLEncoder;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static com.excel.poi.utils.GetGenernalFieldValueByFields.invoke;

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
    public static <T> void export(List<T> sourceList, String[] fields, HttpServletResponse response, String sheetName, Class<? extends EnumDataModel> modelEnumClass){

       /* try{
            if(fields!=null){
                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet sheet = workbook.createSheet(sheetName);
                XSSFRow row = sheet.createRow(0);
                //创建标题行
                for (int i = 0; i < fields.length; i++) {
                    sheet.setColumnWidth(i, 5000);
                    row.createCell(i).setCellValue(fields[i]);
                }

                //处理日期格式
                XSSFDataFormat format = workbook.createDataFormat();
                short s = format.getFormat("yyyy年MM月dd日 HH时mm分ss秒");
                XSSFCellStyle style = workbook.createCellStyle();
                XSSFFont font = workbook.createFont();
                font.setFontName("华文楷体");
                font.setItalic(true);
                font.setBold(true);
                font.setColor(HSSFFont.COLOR_RED);
                style.setFont(font);
                style.setDataFormat(s);

                //处理数据行
                for (int i = 0; i < sourceList.size(); i++) {
                    T t = sourceList.get(i);
                    row = sheet.createRow(i+1);
                    for (int j = 0; j < fields.length; j++) {

                        Object obj = invoke(map.get(fields[j]), t);

                        //对日期类型进行筛选。做日期类型的处理
                        if(obj instanceof Date){
                            sheet.setColumnWidth(j, 10000);
                            XSSFCell cell = row.createCell(j);
                            cell.setCellStyle(style);
                            cell.setCellValue((Date)obj);
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
        }*/

        String str_status = EnumConstantsUtil.valueBy(modelEnumClass, 1);

        System.out.println(str_status);
    }
}
