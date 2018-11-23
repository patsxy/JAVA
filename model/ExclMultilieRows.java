package com.framework.excel.model;

import java.io.Serializable;
//Excl 多行类
public class ExclMultilieRows implements Serializable {
    //值
    private  String value;
    //跨行数
    private  int rowSize=0;
   //跨列数
    private  int  colSize=0;

    public ExclMultilieRows(){

    }

    public ExclMultilieRows(String value){
        this.value=value;
    }

    public ExclMultilieRows(String value,int rowSize){
        this.value=value;
        this.rowSize=rowSize;

    }

    public ExclMultilieRows(String value,int rowSize,int colSize){
        this.value=value;
        this.rowSize=rowSize;
        this.colSize=colSize;

    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public int getRowSize() {
        return rowSize;
    }

    public void setRowSize(int rowSize) {
        this.rowSize = rowSize;
    }


    public int getColSize() {
        return colSize;
    }

    public void setColSize(int colSize) {
        this.colSize = colSize;
    }
}
