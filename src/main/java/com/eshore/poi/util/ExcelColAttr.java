package com.eshore.poi.util;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.lang.String;

/**
 * 对象属性和excel列的对应关系控制类
 * @version 1.0 
 * Created by wangqian on 2016/10/28.
 */
@Target(ElementType.METHOD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColAttr {

    public enum  DataType{INT,STRING,DOUBLE,DATE};

    /**
     * 属性导入excel的colIndex列
     * @return int
     */
    int colIndex();//colIndex列

    /**
     * 属性对应excel的列名
     * @return String
     */
    String colName();//colIndex列的列名称

    /**
     * 数据类型 <br>
     * DataType.INT,DataType.STRING,DataType.DOUBLE,DataType.DATE
     * @return DataType
     */
    DataType dataType();

    /**
     * 日期类型数据的字符串格式设置，默认为"yyyy-MM-dd HH:mm:ss"
     * @return String
     */
    String datePattern() default "yyyy-MM-dd HH:mm:ss";

    /**
     * double类型数据格式设置，默认为"0.00"
     * @return String
     */
    String doublePattern() default "0.00";
}
