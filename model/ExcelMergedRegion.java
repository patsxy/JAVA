package com.framework.excel.model;

import java.io.Serializable;
//excel 合并行列
public class ExcelMergedRegion implements Serializable {
    //是否合并
    private  boolean isMerged=false;

    //列标记
    private  boolean isColMerged=false;
    //firstRow
    private  int firstRow;
    //lastRow
    private  int lastRow;
    //firstCol
    private  int firstCol;
    //lastCol
    private int lastCol;

    public void setMerged(boolean merged) {
        isMerged = merged;
    }

    public boolean isMerged() {
        return isMerged;
    }

    public int getFirstRow() {
        return firstRow;
    }

    public void setFirstRow(int firstRow) {
        this.firstRow = firstRow;
    }

    public int getLastRow() {
        return lastRow;
    }

    public void setLastRow(int lastRow) {
        this.lastRow = lastRow;
    }

    public int getFirstCol() {
        return firstCol;
    }

    public void setFirstCol(int firstCol) {
        this.firstCol = firstCol;
    }

    public int getLastCol() {
        return lastCol;
    }

    public void setLastCol(int lastCol) {
        this.lastCol = lastCol;
    }

    public boolean isColMerged() {
        return isColMerged;
    }

    public void setColMerged(boolean colMerged) {
        isColMerged = colMerged;
    }
}
