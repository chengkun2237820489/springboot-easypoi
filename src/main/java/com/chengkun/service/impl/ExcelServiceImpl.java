package com.chengkun.service.impl;
/**
 * sungrow all right reserved
 **/

import com.chengkun.entity.InsightIec104Mapping;
import com.chengkun.mapper.ExcelMapper;
import com.chengkun.service.ExcelService;
import com.chengkun.utils.EasyPoiUtils;
import com.chengkun.utils.style.ExportExcelStyle;
import com.google.common.collect.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.http.HttpServletResponse;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * @Description Excel导入导出业务代码
 * @Author chengkun
 * @Date 2020/3/17 14:45
 **/
@Service
public class ExcelServiceImpl implements ExcelService {

    @Autowired
    private ExcelMapper excelMapper;

    /**
     * @Description 查询所有，返回实体类集合
     * @Author chengkun
     * @Date 2020/3/18 13:30
     * @Param
     * @Return java.util.List<com.chengkun.entity.InsightIec104Mapping>
     **/
    @Override
    public List<InsightIec104Mapping> findAll() {
        return excelMapper.findAll();
    }

    /**
     * @Description 查询所有，返回map集合
     * @Author chengkun
     * @Date 2020/3/18 13:31
     * @Param
     * @Return java.util.List<java.util.Map < java.lang.String, java.lang.Object>>
     **/
    @Override
    public List<Map<String, Object>> findAllByMap() {
        return excelMapper.findAllByMap();
    }

    /**
     * @param params
     * @Description 根据测点类型查询
     * @Author chengkun
     * @Date 2020/3/18 15:40
     * @Param point_type
     * @Return java.util.List<java.util.Map < java.lang.String, java.lang.Object>>
     */
    @Override
    public List<Map<String, Object>> findAllByPointType(Map<String, String> params) {
        return excelMapper.findAllByPointType(params);
    }

    /**
     * @param params
     * @Description 根据测点类型查询, 返回实体类
     * @Author chengkun
     * @Date 2020/3/18 15:40
     * @Param params
     * @Return java.util.List<java.util.Map < java.lang.String, java.lang.Object>>
     */
    @Override
    public List<InsightIec104Mapping> findAllByPointTypeEntity(Map<String, String> params) {
        return excelMapper.findAllByPointTypeEntity(params);
    }

    /**
     * @param response
     * @Description 导出iec映射
     * @Author chengkun
     * @Date 2020/3/27 17:15
     * @Param response
     * @Return void
     */
    @Override
    public void exportExcelIecMapping(HttpServletResponse response) {
        // 根据电站id获取设备
        List<Map<String, Object>> list = excelMapper.findDeviceByPsId(ImmutableMap.of("ps_id", "1"));
        ArrayListMultimap<String, Map<String, Object>> listMultimap = ArrayListMultimap.create();
        for (Map<String, Object> map : list) {
            String deviceType = String.valueOf(map.get("DEVICE_TYPE"));
            if (!"11".equals(deviceType) && !"3".equals(deviceType) && !"17".equals(deviceType)) {
                listMultimap.put(String.valueOf(map.get("DEVICE_TYPE")), map);
            }
        }
        //获取所有通道
        List<Map<String, Object>> channelList = excelMapper.findAllChannel();
        List<String> channelIds = Lists.newArrayList();
        for (Map<String, Object> map : channelList) {
            channelIds.add(String.valueOf(map.get("CHNNL_ID")));
        }
        int index = listMultimap.values().size() / channelList.size();
        List<Map<String, Object>> dataList = Lists.newArrayList();
        Map<String, Object> data;
        int counte = 0;
        int x = 1;
        int y = 16385;
        String channelId = "0";
        Set<String> channelSet = Sets.newLinkedHashSet();
        for (String key : listMultimap.keySet()) {
            List<Map<String, Object>> deviceList = listMultimap.get(key);
            //根据设备类型和测点类型查询测点
            List<Map<String, Object>> pointList1 = excelMapper.findPointList(ImmutableMap.of("device_type", key, "point_type", "1")); //遥信
            List<Map<String, Object>> pointList2 = excelMapper.findPointList(ImmutableMap.of("device_type", key, "point_type", "2")); //遥测
            for (int i = 0; i < deviceList.size(); i++) {
                if (counte / index >= channelIds.size()) {
                    channelId = channelIds.get(channelIds.size() - 1);
                } else {
                    channelId = channelIds.get(counte / index);
                }

                if (!channelSet.contains(channelId) && channelSet.size() > 0) {
                    x = 1;
                    y = 16385;
                }
                channelSet.add(channelId);
                for (Map<String, Object> point : pointList1) {
                    data = Maps.newLinkedHashMap();
                    data.put("channel_id", channelId);
                    data.put("uuid", deviceList.get(i).get("UUID"));
                    data.put("point_type", "遥信信息");
                    data.put("point_name", point.get("POINT_NAME"));
                    data.put("iec_point", x++);
                    data.put("point_id", point.get("POINT_ID"));
                    data.put("is_enable", "1");
                    data.put("iec_negate", "0");
                    dataList.add(data);
                }
                for (Map<String, Object> point : pointList2) {
                    data = Maps.newLinkedHashMap();
                    data.put("channel_id", channelId);
                    data.put("uuid", deviceList.get(i).get("UUID"));
                    data.put("point_type", "遥测信息");
                    data.put("point_name", point.get("POINT_NAME"));
                    data.put("iec_point", y++);
                    data.put("point_id", point.get("POINT_ID"));
                    data.put("is_enable", "1");
                    data.put("iec_parameter", "1");
                    data.put("iec_offset", "0");
                    dataList.add(data);
                }
                counte++;
            }

        }

        Map<String, Object> headerMap = Maps.newLinkedHashMap(); //定义标题行
        headerMap.put("通道号", "channel_id");
        headerMap.put("设备编号", "uuid");
        headerMap.put("测点类型", "point_type");
        headerMap.put("测点名称", "point_name");
        headerMap.put("IEC测点", "iec_point");
        headerMap.put("设备测点", "point_id");
        headerMap.put("状态(1：启用，0：停用)", "is_enable");
        headerMap.put("修正系数", "iec_parameter");
        headerMap.put("偏移量", "iec_offset");
        headerMap.put("取反(1：是，0：否)", "iec_negate");
        EasyPoiUtils.exportExcelForMap(dataList, "easypoi导出map功能", "Export", headerMap, "easypoi导出map功能.xlsx", null, ExportExcelStyle.class, response);
    }

}