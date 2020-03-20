package com.chengkun.entity;
/**
 * sungrow all right reserved
 **/

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

/**
 * @Description 导入电站Excel对象
 * @Author chengkun
 * @Date 2020/3/19 11:17
 **/
@Data
public class PowerDeviceExcelPojo {

    @Excel(name = "设备编号", orderNum = "1", width = 30)
    private String uuid;

    @Excel(name = "所属上级设备编号", orderNum = "2", width = 30)
    private String up_uuid;

    @Excel(name = "通道号", orderNum = "3", width = 30)
    private String channel_id;

    @Excel(name = "设备名称", orderNum = "4", width = 30)
    private String device_name;

    @Excel(name = "所属设备类型", orderNum = "5", width = 30)
    private String device_type;

    @Excel(name = "所属设备型号", orderNum = "6", width = 30)
    private String device_model;

    @Excel(name = "所属电站编码", orderNum = "7", width = 30)
    private String ps_id;

    @Excel(name = "装机功率", orderNum = "8", width = 30)
    private String install_power;

    @Excel(name = "额定功率", orderNum = "9", width = 30)
    private String rate_power;
}