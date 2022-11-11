package com.chengkun.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import cn.afterturn.easypoi.excel.annotation.ExcelCollection;
import lombok.Data;

import java.util.List;

/**
 * FileName: Good
 * Author:   ck
 * Date:     2022/10/2 17:26
 * Description: 商品
 */
@Data
public class Good extends ExcelVerifyInfo {

    private Integer id;

    //日期
    @Excel(name = "日期", orderNum = "1", width = 30)
    private String dt;

    //产品PV
    @Excel(name = "产品PV", orderNum = "2", width = 30)
    private String pv;

    // 业务数据
    @ExcelCollection(name = "业务数据", orderNum = "3")
    private List<BusinessData> businessData;
}
