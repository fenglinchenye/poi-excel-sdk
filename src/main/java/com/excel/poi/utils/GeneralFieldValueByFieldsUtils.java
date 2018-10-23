package com.excel.poi.utils;


import com.excel.poi.enums.Constant;
import com.excel.poi.enums.FieldType;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Locale;

/**
 * @author zx
 * 根据属性名获得属性值
 * 属性名形如 ：patient.name 等 多级属性用点分割
 */
@Slf4j
public class GeneralFieldValueByFieldsUtils {

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
    public static Object getFieldValue(String field ,Object sourceObject) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException, NoSuchMethodException, SecurityException{

        String getMethodName = "get";

        if(StringUtils.isNotBlank(field)){

            int seat = field.lastIndexOf(Constant.POINT_SEPARATOR);

            if( seat == -1){
                getMethodName+= field.substring(0,1).toUpperCase()+field.substring(1);
                //创建方法
                Method method = sourceObject.getClass().getMethod(getMethodName, null);
                //通过方法调用获得返回值
                Object obj = method.invoke(sourceObject, null);

                return obj;
            }else{
                String[] splits = field.split("\\.");
                Object child = sourceObject;
                for (String string : splits) {
                    String methodName = getMethodName + string.substring(0,1).toUpperCase()+string.substring(1);
                    Method method = child.getClass().getMethod(methodName, null);
                    child = method.invoke(child, null);
                }
                return child;
            }
        } else {
            return null;
        }
    }


    /**
     * set属性的值到Bean
     * @param sourceClazz
     * @param fieldNames
     * @param fieldValues
     *
     */
    public static <T> T createObjectByFields(Class<T> sourceClazz, List<String> fieldNames,List<String> fieldValues) throws IllegalAccessException, InstantiationException {

        T bean = sourceClazz.newInstance();

        // 取出bean里的所有方法
        Method[] methods = sourceClazz.getDeclaredMethods();

        for (int i = 0; i < fieldNames.size(); i++) {
            try {

                // 获得set的name 属性
                String fieldSetName = parSetName(fieldNames.get(i));
                if (!checkSetMet(methods, fieldSetName)) {
                    continue;
                }

                String value = fieldValues.get(i);

                // 执行set 注入方法属性
                invokeSetMethod(sourceClazz.newInstance(),fieldNames.get(i),value,bean);

            } catch (Exception e) {
                continue;
            }
        }
        return bean;
    }

    /**
     * 格式化string为Date
     * @param dateStr
     * @return date
     */
    private static Date parseDate(String dateStr) {
        if (null == dateStr || "".equals(dateStr)) {
            return null;
        }
        try {
            String formatStr = null;
            if (dateStr.indexOf(Constant.COLON_SEPARATOR) > 0) {
                formatStr = Constant.TIME_ALL_PATTERN;
            } else {
                formatStr = Constant.TIME_SIMPLE_PATTERN;
            }
            SimpleDateFormat sdf = new SimpleDateFormat(formatStr, Locale.CHINESE);
            return sdf.parse(dateStr);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 日期转化为String
     * @param date
     * @return date string
     */
    private static String formatDate(Date date) {
        if (null == date) {
            return null;
        }
        try {
            SimpleDateFormat sdf = new SimpleDateFormat(Constant.TIME_ALL_PATTERN, Locale.CHINESE);
            return sdf.format(date);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * 判断是否存在某属性的 set方法
     * @param methods
     * @param fieldSetMet
     * @return boolean
     */
    private static boolean checkSetMet(Method[] methods, String fieldSetMet) {
        for (Method met : methods) {
            if (fieldSetMet.equals(met.getName())) {
                return true;
            }
        }
        return false;
    }

    /**
     * 判断是否存在某属性的 get方法
     * @param methods
     * @param fieldGetMet
     * @return boolean
     */
    private static boolean checkGetMethod(Method[] methods, String fieldGetMet) {
        for (Method met : methods) {
            if (fieldGetMet.equals(met.getName())) {
                return true;
            }
        }
        return false;
    }

    /**
     * 拼接某属性的 get方法
     * @param fieldName
     * @return String
     */
    private static String parGetName(String fieldName) {
        if (StringUtils.isBlank(fieldName)) {
            return null;
        }

        if(fieldName.lastIndexOf(Constant.POINT_SEPARATOR)!=-1){
            return "get" + fieldName.substring(0, 1).toUpperCase()
                    + fieldName.substring(1,fieldName.indexOf(Constant.POINT_SEPARATOR));
        }

        return "get" + fieldName.substring(0, 1).toUpperCase()
                + fieldName.substring(1);
    }
    /**
     * 拼接在某属性的 set方法
     * @param fieldName
     * @return String
     */
    private static String parSetName(String fieldName) {

        if (StringUtils.isBlank(fieldName)) {
            return null;
        }

        if(fieldName.lastIndexOf(Constant.POINT_SEPARATOR)!=-1){
            return "set" + fieldName.substring(0, 1).toUpperCase()
                    + fieldName.substring(1,fieldName.indexOf(Constant.POINT_SEPARATOR));
        }

        return "set" + fieldName.substring(0, 1).toUpperCase()
                + fieldName.substring(1);
    }

    /**
     * 获得属性的类型
     * @param sourceClazz 源数据类
     * @param fieldName 属性名
     * @param <T>
     * @return
     * @throws NoSuchFieldException
     */
    private static <T>Class<?> getLastFieldType(Class<T> sourceClazz,String fieldName) throws NoSuchFieldException {

        if(fieldName.indexOf(Constant.POINT_SEPARATOR)!=-1){
            String endFiledName = fieldName.substring(fieldName.indexOf(Constant.POINT_SEPARATOR)+1);
            return getLastFieldType(sourceClazz.getDeclaredField(fieldName.substring(0,fieldName.indexOf(Constant.POINT_SEPARATOR))).getType(),endFiledName);
        }else{
            return sourceClazz.getDeclaredField(fieldName).getType();
        }
    }

    /**
     *  获得倒数第二个的类型
     * @param sourceClazz
     * @param fieldName
     * @param <T>
     * @return
     */
    private static <T> Class<?> getReverseSecondFieldType(Class<T> sourceClazz,String fieldName) throws NoSuchFieldException {

        if(fieldName.indexOf(Constant.POINT_SEPARATOR) == -1){
            return sourceClazz;
        }else{
            Field field = sourceClazz.getDeclaredField(fieldName.substring(0, fieldName.indexOf(Constant.POINT_SEPARATOR)));
            return getReverseSecondFieldType(field.getType(),fieldName.substring(fieldName.indexOf(Constant.POINT_SEPARATOR)+1));
        }
    }

    /**
     * 获得最顶级的属性类型
     * @param sourceClazz
     * @param fieldName
     * @param <T>
     * @return
     * @throws NoSuchFieldException
     */
    private static <T>Class getFirstFieldType(Class<T> sourceClazz,String fieldName) throws NoSuchFieldException {

        if(fieldName.indexOf(Constant.POINT_SEPARATOR)!=-1){
            String endFiledName = fieldName.substring(0,fieldName.indexOf(Constant.POINT_SEPARATOR));
            return sourceClazz.getDeclaredField(endFiledName).getType();
        }else{
            return sourceClazz.getDeclaredField(fieldName).getType();
        }
    }

    /**
     * 执行基本类型及String类型的数据转换处理
     * @param t
     * @param fileTypeName
     * @param fieldSetMet
     * @param value
     * @param <T>
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     */
    private static <T> void invokeBaseType(T t,String fileTypeName,Method fieldSetMet,String value) throws InvocationTargetException, IllegalAccessException {

        if (StringUtils.isNotBlank(value)) {
            if (FieldType.STRING.equals(fileTypeName)) {
                fieldSetMet.invoke(t, value);
            } else if (FieldType.DATE.equals(fileTypeName)) {
                Date temp = parseDate(value);
                fieldSetMet.invoke(t, temp);
            } else if (FieldType.INTEGER.equals(fileTypeName) || FieldType.INT.equals(fileTypeName)) {
                fieldSetMet.invoke(t, Integer.parseInt(value));
            } else if (FieldType.LONG.equalsIgnoreCase(fileTypeName)) {
                Long temp = Long.parseLong(value);
                fieldSetMet.invoke(t, temp);
            } else if (FieldType.DOUBLE.equalsIgnoreCase(fileTypeName)) {
                Double temp = Double.parseDouble(value);
                fieldSetMet.invoke(t, temp);
            } else if (FieldType.BOOLEAN.equalsIgnoreCase(fileTypeName)) {
                Boolean temp = Boolean.parseBoolean(value);
                fieldSetMet.invoke(t, temp);
            } else {
                log.info("not supper type" + fileTypeName);
            }
        }
    }

    /**
     *  执行set方法注入
     * @param t
     * @param fieldName
     * @param fieldValue
     * @param sourceBean
     * @param <T>
     * @param <E>
     * @throws NoSuchFieldException
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    private static <T,E> void invokeSetMethod(E t,String fieldName,Object fieldValue,T sourceBean) throws NoSuchFieldException, NoSuchMethodException, InvocationTargetException, IllegalAccessException, InstantiationException {

        Class<?> sourceClass = sourceBean.getClass();
        Method fieldGetMet = sourceClass.getMethod(parGetName(fieldName),null);

        Object value = fieldGetMet.invoke(sourceBean, null);
        if(value != null){
            if(fieldName.indexOf(Constant.POINT_SEPARATOR)!=-1){
                // 递归调用
                invokeSetMethod(value,fieldName.substring(fieldName.indexOf(Constant.POINT_SEPARATOR)+1),fieldValue,sourceClass);
            }else{
                Method fieldSetMet = t.getClass().getMethod(parSetName(fieldName), getFirstFieldType(t.getClass(), fieldName));
                invokeBaseType(t,getFirstFieldType(sourceClass, fieldName).getSimpleName(),fieldSetMet,fieldValue.toString());
            }
        } else {
            if(fieldName.indexOf(Constant.POINT_SEPARATOR)!=-1) {
                Object reverseSecondField = getReverseSecondFieldType(sourceClass, fieldName).newInstance();
                String lastFieldName = fieldName.substring(fieldName.lastIndexOf(Constant.POINT_SEPARATOR) + 1);
                Method fieldSetMet = reverseSecondField.getClass().getMethod(parSetName(lastFieldName), reverseSecondField.getClass().getDeclaredField(fieldName.substring(fieldName.lastIndexOf(Constant.POINT_SEPARATOR)+1)).getType());
                invokeBaseType(reverseSecondField,getFirstFieldType(reverseSecondField.getClass(), lastFieldName).getSimpleName(),fieldSetMet,fieldValue.toString());

                invokeObjectType(fieldName.substring(0,fieldName.lastIndexOf(Constant.POINT_SEPARATOR)),reverseSecondField,sourceBean);
            } else {

                Method fieldSetMet = t.getClass().getMethod(parSetName(fieldName), getFirstFieldType(sourceClass, fieldName));
                invokeBaseType(sourceBean,getFirstFieldType(sourceClass, fieldName).getSimpleName(),fieldSetMet,fieldValue.toString());
            }
        }
    }

    /**
     * 执行自定义对象类型的数据处理
     * @param fieldName
     * @param fieldValue
     * @param source
     * @param <T>
     * @param <E>
     * @throws NoSuchFieldException
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    private static <T,E>void invokeObjectType(String fieldName,E fieldValue,T source) throws NoSuchFieldException, NoSuchMethodException, InvocationTargetException, IllegalAccessException, InstantiationException {

        if(fieldName.indexOf(Constant.POINT_SEPARATOR)==-1){
            String setMet = parSetName(fieldName);
            Method fieldSetMet = source.getClass().getMethod(setMet, getFirstFieldType(source.getClass(), fieldName));
            if(checkSetMet(source.getClass().getMethods(), setMet)){
                fieldSetMet.invoke(source,fieldValue);
            }
        }else{
            Object obj = getReverseSecondFieldType(source.getClass(), fieldName).newInstance();

            String setMet = parSetName(fieldName.substring(fieldName.lastIndexOf(Constant.POINT_SEPARATOR) + 1));
            Method fieldSetMet = obj.getClass().getMethod(setMet, getLastFieldType(source.getClass(), fieldName));

            if(checkSetMet(obj.getClass().getMethods(),setMet)){
                fieldSetMet.invoke(obj,fieldValue);
            }

            // 递归调用
            invokeObjectType(fieldName.substring(0,fieldName.lastIndexOf(Constant.POINT_SEPARATOR)),obj,source);
        }
    }
}
