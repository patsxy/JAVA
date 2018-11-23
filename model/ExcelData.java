package com.framework.excel.model;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

/**
 * excl 导入导出类
 */
public class ExcelData implements Serializable {


    private static final long serialVersionUID = 4444017239100620999L;



    //表标题
    private List head;

    //标题合并
    private  List<ExcelMergedRegion>  headMergedRegions;
    //头部样式
    private List<ExcelFontStyle> headStyle;

    // 表头
    private List titles;

    //表头合并
    private  List<ExcelMergedRegion>  titlesMergedRegions;

    //表头样式
    private List<ExcelFontStyle> titleStyle;
    // 数据
    private List rows;
    //行列合并
    private  List<ExcelMergedRegion>  rowsMergedRegions;
    //行内样式
    private List<ExcelFontStyle> rowStyle;
    //多行多列
    private  int  is_Multiline=0;
    //多表头
    private  int  is_MultiTitles=0;
    //列长度
    private  int colSize=0;

    //每列宽带
    private List<Integer>  colsWidth;


    //高度
    private short colsHight=0;

    // 页签名称
    private String name;

    //是否自动按照 title 设置头款
    private  boolean  is_autoWith=false;


    public boolean isIs_autoWith() {
        return is_autoWith;
    }

    public List<Integer> getColsWidth() {
        return colsWidth;
    }


    public short getColsHight() {
        return colsHight;
    }

    public void setColsHight(short colsHight) {
        this.colsHight = colsHight;
    }

    public void setColsWidth(List<Integer> colsWidth) {
        this.colsWidth = colsWidth;
    }

    public void setIs_autoWith(boolean is_autoWith) {
        this.is_autoWith = is_autoWith;
    }

    public List getHead() {
        return head;
    }

    public void setHead(List head) {
        this.head = head;
    }

    public List<ExcelMergedRegion> getHeadMergedRegions() {
        return headMergedRegions;
    }

    public void setHeadMergedRegions(List<ExcelMergedRegion> headMergedRegions) {
        this.headMergedRegions = headMergedRegions;
    }

    public List<ExcelFontStyle> getHeadStyle() {
        return headStyle;
    }

    public void setHeadStyle(List<ExcelFontStyle> headStyle) {
        this.headStyle = headStyle;
    }

    public List getTitles() {
        return titles;
    }

    public void setTitles(List titles) {
        this.titles = titles;
    }

    public List<ExcelMergedRegion> getTitlesMergedRegions() {
        return titlesMergedRegions;
    }

    public void setTitlesMergedRegions(List<ExcelMergedRegion> titlesMergedRegions) {
        this.titlesMergedRegions = titlesMergedRegions;
    }

    public List getRows() {
        return rows;
    }

    public void setRows(List rows) {
        this.rows = rows;
    }

    public List<ExcelMergedRegion> getRowsMergedRegions() {
        return rowsMergedRegions;
    }

    public void setRowsMergedRegions(List<ExcelMergedRegion> rowsMergedRegions) {
        this.rowsMergedRegions = rowsMergedRegions;
    }

    public int getIs_Multiline() {
        return is_Multiline;
    }

    public void setIs_Multiline(int is_Multiline) {
        this.is_Multiline = is_Multiline;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getIs_MultiTitles() {
        return is_MultiTitles;
    }

    public void setIs_MultiTitles(int is_MultiTitles) {
        this.is_MultiTitles = is_MultiTitles;
    }


    public List<ExcelFontStyle> getRowStyle() {
        return rowStyle;
    }

    public void setRowStyle(List<ExcelFontStyle> rowStyle) {
        this.rowStyle = rowStyle;
    }

    public int getColSize() {
        return colSize;
    }

    public void setColSize(int colSize) {
        this.colSize = colSize;
    }


    public List<ExcelFontStyle> getTitleStyle() {
        return titleStyle;
    }

    public void setTitleStyle(List<ExcelFontStyle> titleStyle) {
        this.titleStyle = titleStyle;
    }
}
