package com.eshore.poi.test;


import java.util.Date;

import com.eshore.poi.util.ExcelColAttr;

/**
 * Created by wangqian on 2016/10/28.
 */
public class SimpleGrid {
    public SimpleGrid(){

    }

    public  SimpleGrid(int index,String name,Date date,Double point){
        this.index = index;
        this.name  = name;
        this.date  = date;
        this.point = point;
    }

    @ExcelColAttr(colIndex = 0,colName = "工号",dataType = ExcelColAttr.DataType.INT)
    public int getIndex() {
        return index;
    }

    @ExcelColAttr(colIndex = 1,colName = "名字",dataType = ExcelColAttr.DataType.STRING)
    public String getName() {
        return name;
    }

    @ExcelColAttr(colIndex = 2,colName = "出生时间",dataType = ExcelColAttr.DataType.DATE)
    public Date getDate() {
        return date;
    }

    @ExcelColAttr(colIndex = 3,colName = "得分",dataType = ExcelColAttr.DataType.DOUBLE)
    public double getPoint() {
        return point;
    }

    @Override
    public String toString(){
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("工号:"+index+",");
        stringBuilder.append("名字:"+name+",");
        stringBuilder.append("出生时间:"+date+",");
        stringBuilder.append("得分:"+point);
        return stringBuilder.toString();
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public void setIndex(Integer index){
        this.index = index.intValue();
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public void setPoint(double point) {
        this.point = point;
    }

    public void setPoint(Double point){
        this.point = point.doubleValue();
    }

    private int     index;
    private String  name;
    private Date    date;
    private double  point;
}
