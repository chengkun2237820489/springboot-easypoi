package com.chengkun.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

@Data
public class InsightIec104Mapping {

    @Excel(name = "通道号", orderNum = "1", width=30)
    private long channelId;

    @Excel(name = "设备编号", orderNum = "2", width=30)
    private long uuid;

    @Excel(name = "测点类型", orderNum = "3", width=30)
    private String pointType;

    @Excel(name = "测点名称", orderNum = "4", width=30)
    private String pointName;

    @Excel(name = "IEC测点", orderNum = "5", width=30)
    private long iecPoint;

    @Excel(name = "设备测点", orderNum = "6", width=30)
    private long pointId;

    @Excel(name = "状态(1：启用，0：停用)", orderNum = "7", width=30)
    private String isEnable;

    @Excel(name = "修正系数", orderNum = "8", width=30)
    private String iecParameter;

    @Excel(name = "偏移量", orderNum = "9", width=30)
    private String iecOffset;

    @Excel(name = "取反(1：是，0：否)", orderNum = "10", width=30)
    private String iecNegate;

}
