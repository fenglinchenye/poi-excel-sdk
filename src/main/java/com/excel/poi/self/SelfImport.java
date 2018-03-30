package com.excel.poi.self;

import com.excel.poi.utils.GeneralFieldValueByFields;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

/**
 *
 * 自定义导入 (目前只支持2007 版本  不支持2003 版本)
 *
 */
public class SelfImport {

    private final static String excel2003L =".xls";    //2003- 版本的excel
    private final static String excel2007U =".xlsx";   //2007+ 版本的excel

    private static List<String> columnNameList;

    public static void setColumnNameList(List<String> columnNameList) {
        SelfImport.columnNameList = columnNameList;
    }

    /**
     * 获取IO流中的数据，组装成List<T>对象
     * @param in
     * @param fileName
     * @param targetClazz
     * @param <T>
     * @return
     * @throws Exception
     */
    public <T> List<T> getBankListByExcel(InputStream in, String fileName, Class<T> targetClazz) throws Exception{

        //创建Excel工作薄
        Workbook work = this.getWorkbook(in,fileName);
        if(null == work){
            throw new Exception("创建Excel工作薄为空！");
        }
        Sheet sheet = null;
        Row row = null;

        sheet = work.getSheetAt(0);
        if(sheet==null){return null;}
        // 目标对象集合
        List<T> targetList = new ArrayList<T>();

        //遍历当前sheet中的所有行
        for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
            row = sheet.getRow(j);
            if(row==null||row.getFirstCellNum()==j){continue;}
            List<String> valueList = new ArrayList<String>();

            for (short i = row.getFirstCellNum(); i < row.getLastCellNum() ; i++) {
                Cell cell = row.getCell(i);
                valueList.add(this.getCellValue(cell).toString());
            }

            T target = GeneralFieldValueByFields.createObjectByFields(targetClazz, columnNameList, valueList);

            targetList.add(target);

        }
        in.close();
        return targetList;
    }

    /**
     * 描述：根据文件后缀，自适应上传文件的版本
     * @param inStr,fileName
     * @return
     * @throws Exception
     */
    public  Workbook getWorkbook(InputStream inStr, String fileName) throws Exception{
        Workbook wb = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if(excel2003L.equals(fileType)){
            wb = new HSSFWorkbook(inStr);  //2003-
        }else if(excel2007U.equals(fileType)){
            wb = new XSSFWorkbook(inStr);  //2007+
        }else{
            throw new Exception("解析的文件格式有误！");
        }
        return wb;
    }

    /**
     * 描述：对表格中数值进行格式化
     * @param cell
     * @return
     */
    public  Object getCellValue(Cell cell){
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
