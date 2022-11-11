package com.chengkun.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

/**
 * FileName: BusinessData
 * Author:   ck
 * Date:     2022/10/2 17:29
 * Description: 业务数据
 */
@Data
public class BusinessData extends ExcelVerifyInfo {

    private Integer id;

    // 数据
    @Excel(name = "数据1", orderNum = "1", width = 30)
    private String data1;

    // 数据
    @Excel(name = "数据2", orderNum = "1", width = 30)
    private String data2;

    // 数据
    @Excel(name = "数据3", orderNum = "1", width = 30)
    private String data3;

    // 数据
    @Excel(name = "数据4", orderNum = "1", width = 30)
    private String data4;
}
