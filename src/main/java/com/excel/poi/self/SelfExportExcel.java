package com.excel.poi.self;

import com.excel.poi.enums.EnumDataStatusModel;
import com.excel.poi.utils.CellUtils;
import com.excel.poi.utils.EnumConstantsUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.Date;
import java.util.List;
import java.util.Map;

import static com.excel.poi.utils.GeneralFieldValueByFieldsUtils.getFieldValue;


/**
 * 自定义普通导出
 * 不能进行合并单元格数据
 */
@Slf4j
public class SelfExportExcel<T> {

    /**
     * 映射 导出的列名与实体数据中的属性名之间的对应关系
     * k:列名   v: 属性名
     */
    private Map<String,String> map;

    public  Map<String, String> getMap() {
        return map;
    }

    public  void setMap(Map<String, String> map) {
        this.map = map;
    }

    public SelfExportExcel() {
    }

    /**
     * 构造方法
     * @param map
     */
    public SelfExportExcel(Map<String, String> map) {
        this.map = map;
    }

    /**
     *  导出操作
     * @param sourceList  源数据
     * @param fields 导出的列名
     * @param response 响应流
     * @param sheetName 工作页名
     * @param modelEnumClass  必须实现接口EnumDataModel的枚举类(枚举类中有对应的属性)
     */
    public void export(List<T> sourceList, String[] fields, HttpServletResponse response, String sheetName, Class<? extends EnumDataStatusModel> modelEnumClass){

        OutputStream out = null;
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

                        // 获取单元格的值
                        Object obj = getFieldValue(map.get(fields[j]), t);

                        //对日期类型进行筛选。做日期类型的处理
                        if(obj instanceof Date){
                            CellUtils.setCellDateValue(workbook,(Date)obj,j,row);
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
                out = response.getOutputStream();
                workbook.write(out);
            }
        }catch(Exception e){
            log.error("【自定义导出excel】导出异常");
        }finally {
            if (out!=null){
                try {
                    out.close();
                } catch (IOException e) {
                    log.error("【自定义导出excel】关闭流异常,exception={}",e);
                }
            }
        }
    }
}
