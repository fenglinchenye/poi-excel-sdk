package com.excel.poi.utils;


import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 *
 * 根据属性名获得属性值
 * 属性名形如 ：patient.name 等 多级属性用点分割
 *
 */
public class GetGenernalFieldValueByFields {

    /**
     * 获得属性值
     * @param field 属性名
     * @param sourceObject 源对象 （要求必须提供get办法）
     * @return
     * @throws IllegalAccessException
     * @throws IllegalArgumentException
     * @throws InvocationTargetException
     * @throws NoSuchMethodException
     * @throws SecurityException
     */
    public static Object invoke(String field ,Object sourceObject) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException{

        String getMethodName = "get";

        if(StringUtils.isNotBlank(field)){

            int seat = field.lastIndexOf(".");

            if( seat == -1){
                getMethodName+= field.substring(0,1).toUpperCase()+field.substring(1, field.length());
                //创建方法
                Method method = sourceObject.getClass().getMethod(getMethodName, null);
                //通过方法调用获得返回值
                Object obj = method.invoke(sourceObject, null);

                return obj;
            }else{
                String[] splits = field.split("\\.");
                Object child = sourceObject;
                for (String string : splits) {
                    String methodName = getMethodName + string.substring(0,1).toUpperCase()+string.substring(1, string.length());
                    Method method = child.getClass().getMethod(methodName, null);
                    child = method.invoke(child, null);
                }
                return child;
            }
        }
        else return null;
    }
}
