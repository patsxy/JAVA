package com.zy.wordtodb.model;

public class ExcelCellType {
    //单元格类型
    private  int  cellType;
    //单元格列数 第几列 从0开始
    private  int  columnIndex;
    //格式内容
    private  String  cellFormula;

    public int getCellType() {
        return cellType;
    }

    public void setCellType(int cellType) {
        this.cellType = cellType;
    }

    public int getColumnIndex() {
        return columnIndex;
    }

    public void setColumnIndex(int columnIndex) {
        this.columnIndex = columnIndex;
    }

    public String getCellFormula() {
        return cellFormula;
    }

    public void setCellFormula(String cellFormula) {
        this.cellFormula = cellFormula;
    }
}
