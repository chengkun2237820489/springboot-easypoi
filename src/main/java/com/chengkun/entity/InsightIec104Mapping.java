package com.chengkun.entity;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotNull;
import javax.validation.constraints.Pattern;
import java.util.Objects;

@Data
public class InsightIec104Mapping extends ExcelVerifyInfo {

    @NotBlank(message = "通道号不能为空")
    @NotNull
    @Pattern(regexp = "[\\d+]*", message = "设备通道号必须为数字")
    @Excel(name = "通道号", orderNum = "1", width = 30)
    private String channelId;

    @Excel(name = "设备编号", orderNum = "2", width = 30)
    private Integer uuid;

    @Excel(name = "测点类型", orderNum = "3", width = 30)
    private String pointType;

    @Excel(name = "测点名称", orderNum = "4", width = 30)
    private String pointName;

    @Excel(name = "IEC测点", orderNum = "5", width = 30)
    private Integer iecPoint;

    @Excel(name = "设备测点", orderNum = "6", width = 30)
    private Integer pointId;

    @Excel(name = "状态(1：启用，0：停用)", orderNum = "7", width = 30)
    private String isEnable;

    @Excel(name = "修正系数", orderNum = "8", width = 30)
    private String iecParameter;

    @Excel(name = "偏移量", orderNum = "9", width = 30)
    private String iecOffset;

    @Excel(name = "取反(1：是，0：否)", orderNum = "10", width = 30)
    private String iecNegate;

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        InsightIec104Mapping that = (InsightIec104Mapping) o;
        return Objects.equals(channelId, that.channelId);
    }

    @Override
    public int hashCode() {
        return Objects.hash(channelId);
    }

    @Override
    public void setRowNum(Integer rowNum) {
        super.setRowNum(rowNum + 1); //没有把标题行计算上去
    }
}
