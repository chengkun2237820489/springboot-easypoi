package com.chengkun.entity;

import cn.afterturn.easypoi.handler.inter.IExcelDataModel;
import cn.afterturn.easypoi.handler.inter.IExcelModel;

/**
 * FileName: ExcelVerifyInfo
 * Author:   ck
 * Date:     2022/10/1 14:07
 * Description: Excel统一校验类
 */
public class ExcelVerifyInfo implements IExcelModel, IExcelDataModel {

    private String errorMsg;

    private int rowNum;

    @Override
    public Integer getRowNum() {
        return rowNum;
    }

    @Override
    public void setRowNum(Integer rowNum) {
        this.rowNum = rowNum;
    }

    @Override
    public String getErrorMsg() {
        return errorMsg;
    }

    @Override
    public void setErrorMsg(String errorMsg) {
        this.errorMsg = errorMsg;
    }
}
