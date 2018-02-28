package com.excel.poi.utils;

import com.excel.poi.enums.EnumDataModel;

/**
 * 根据enum的键获取其值
 */
public class EnumConstantsUtil {


    /**
     *
     * @param sourceEnumClass 源枚举类
     * @param code 键值
     * @param <T>
     * @return
     */
    public static <T extends EnumDataModel> String valueBy(Class<T> sourceEnumClass,int code){

        T[] constants = sourceEnumClass.getEnumConstants();

        if(constants.length!=0) {

            for (int i = 0; i < constants.length; i++) {

                if(constants[i].getClass().isEnum() && ((EnumDataModel)constants[i]).getCode()==code){
                    return ((EnumDataModel)constants[i]).getMessage();
                }
            }
        }
        return "";
    }
}
