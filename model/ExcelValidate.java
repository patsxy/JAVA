package com.framework.excel.model;

import java.io.Serializable;
import java.util.List;

/**
 * excel 验证类
 */
public class ExcelValidate implements Serializable {
    //验证状态
    private  boolean state;
    //验证结果
    private  String  msg;
    //写入类数据
    private   List   arrlist;

    public ExcelValidate(){

    }
    public ExcelValidate(boolean state, String msg) {
        this.state = state;
        this.msg = msg;
    }

    public boolean getState() {
        return state;
    }

    public void setState(boolean state) {
        this.state = state;
    }

    public String getMsg() {
        return msg;
    }

    public void setMsg(String msg) {
        this.msg = msg;
    }

    public boolean isState() {
        return state;
    }

    public List getArrlist() {
        return arrlist;
    }

    public void setArrlist(List arrlist) {
        this.arrlist = arrlist;
    }
}
