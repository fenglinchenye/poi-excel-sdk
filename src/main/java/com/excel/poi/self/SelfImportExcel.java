package com.excel.poi.self;

import com.excel.poi.utils.CellUtils;
import com.excel.poi.utils.GeneralFieldValueByFieldsUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

/**
 *
 * 自定义导入 (目前只支持2007 版本  不支持2003 版本)
 *
 */
@Slf4j
public class SelfImportExcel<T> {

    /**
     * 2003- 版本的excel
     */
    private final static String excel2003L =".xls";
    /**
     * 2007+ 版本的excel
     */
    private final static String excel2007U =".xlsx";

    private List<String> columnNameList;

    public void setColumnNameList(List<String> columnNameList) {
        this.columnNameList = columnNameList;
    }

    public SelfImportExcel() {
    }

    public SelfImportExcel(List<String> columnNameList) {
        this.columnNameList = columnNameList;
    }

    /**
     * 获取IO流中的数据，组装成List<T>对象
     * @param in
     * @param fileName
     * @param targetClazz
     * @return
     * @throws Exception
     */
    public List<T> getBankListByExcel(InputStream in, String fileName, Class<T> targetClazz){

        try {
            //创建Excel工作薄
            Workbook work = this.getWorkbook(in, fileName);
            if (null == work) {
                throw new Exception("创建Excel工作薄为空！");
            }
            Sheet sheet = null;
            Row row = null;

            sheet = work.getSheetAt(0);
            if (sheet == null) {
                return null;
            }
            // 目标对象集合
            List<T> targetList = new ArrayList<T>();

            //遍历当前sheet中的所有行
            for (int j = sheet.getFirstRowNum(); j <= sheet.getLastRowNum(); j++) {
                row = sheet.getRow(j);
                if (row == null || row.getFirstCellNum() == j) {
                    continue;
                }
                List<String> valueList = new ArrayList<String>();

                for (short i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                    Cell cell = row.getCell(i);
                    valueList.add(CellUtils.getCellValue(cell).toString());
                }

                T target = GeneralFieldValueByFieldsUtils.createObjectByFields(targetClazz, columnNameList, valueList);

                targetList.add(target);
            }
            return targetList;
        }catch (Exception e){
            log.error("【自定义导入数据】导入数据异常,exception={}",e);
            return Collections.EMPTY_LIST;
        }finally {
            if (in!=null){
                try {
                    in.close();
                } catch (IOException e) {
                    log.error("【自定义导入数据】输入流关闭异常,exception={}",e);
                }
            }
        }
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
            //2003-
            wb = new HSSFWorkbook(inStr);
        }else if(excel2007U.equals(fileType)){
            //2007+
            wb = new XSSFWorkbook(inStr);
        }else{
            throw new Exception("解析的文件格式有误！");
        }
        return wb;
    }
}
