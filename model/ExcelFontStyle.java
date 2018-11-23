package com.framework.excel.model;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.awt.*;
import java.io.Serializable;

public class ExcelFontStyle implements Serializable {
    //行号  -99 时 为全部
    private  int row=0;
    //字体
    private  String fontName;
    //是否加粗
    private  boolean isBold=false;
    //字号
    private  short fontHeightInPoints;
    //颜色 IndexedColors
    private short color;
    //横向格式
    private HorizontalAlignment horizontalAlignment;
    //纵向格式
    private VerticalAlignment   verticalAlignment;
    //填充背景色
    private Color fillForegroundColor;
    //单元格填充样式
    private   FillPatternType fillPatternType;
    //边框
    private   boolean isBorder=false;
    //边框颜色
    private  Color borderColor;
    //边框粗细
    private BorderStyle  borderStyle;


    //单元格  \n  换行
    private  boolean isWrapText=false;
    //行高
    private  Float   height;


    public Color getBorderColor() {
        return borderColor;
    }

    public void setBorderColor(Color borderColor) {
        this.borderColor = borderColor;
    }

    public BorderStyle getBorderStyle() {
        return borderStyle;
    }

    public void setBorderStyle(BorderStyle borderStyle) {
        this.borderStyle = borderStyle;
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public boolean isBold() {
        return isBold;
    }

    public void setBold(boolean bold) {
        isBold = bold;
    }

    public short getFontHeightInPoints() {
        return fontHeightInPoints;
    }

    public void setFontHeightInPoints(short fontHeightInPoints) {
        this.fontHeightInPoints = fontHeightInPoints;
    }

    public short getColor() {
        return color;
    }

    public void setColor(short color) {
        this.color = color;
    }

    public HorizontalAlignment getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public void setHorizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public Color getFillForegroundColor() {
        return fillForegroundColor;
    }

    public void setFillForegroundColor(Color fillForegroundColor) {
        this.fillForegroundColor = fillForegroundColor;
    }

    public FillPatternType getFillPatternType() {
        return fillPatternType;
    }

    public void setFillPatternType(FillPatternType fillPatternType) {
        this.fillPatternType = fillPatternType;
    }

    public boolean isBorder() {
        return isBorder;
    }

    public void setBorder(boolean border) {
        isBorder = border;
    }

    public boolean isWrapText() {
        return isWrapText;
    }

    public void setWrapText(boolean wrapText) {
        isWrapText = wrapText;
    }

    public Float getHeight() {
        return height;
    }

    public void setHeight(Float height) {
        this.height = height;
    }


}
