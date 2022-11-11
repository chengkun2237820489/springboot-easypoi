package com.chengkun.utils.handler;
/**
 * sungrow all right reserved
 **/

import cn.afterturn.easypoi.exception.excel.ExcelImportException;
import cn.afterturn.easypoi.exception.excel.enums.ExcelImportEnum;
import cn.afterturn.easypoi.handler.inter.IReadHandler;
import com.chengkun.entity.InsightIec104Mapping;
import com.chengkun.entity.SaxExcelImportResult;
import com.google.common.collect.Maps;
import lombok.Data;
import lombok.extern.log4j.Log4j2;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.StringJoiner;

/**
 * @Description map导入使用sax数据处理（包含校验规则）
 * @Author chengkun
 * @Date 2020/3/19 13:25
 **/
@Log4j2
@Data
public class DefaultMapSaxVerifyHandler implements IReadHandler<Map<String, Object>> {

    // 解析的数据
    private SaxExcelImportResult<Map<String, Object>> result;

    //自定义标题行，用于校验格式是否正确
    public Map<String, Object> headerMap;

    //数据map，将标题替换为对应key
    private Map<String, Object> dataMap;

    private Map<String, Object> dbDate;

    public DefaultMapSaxVerifyHandler(Map<String, Object> headerMap, Map<String, Object> dbDate) {
        this.headerMap = headerMap;
        this.dbDate = dbDate;
    }

    /**
     * 处理解析对象
     *
     * @param map
     */
    @Override
    public void handler(Map<String, Object> map) {
        //判断Excel中标题行与定义的标题行数量是否一致
        if (headerMap.size() != map.keySet().size()) {
            throw new ExcelImportException(ExcelImportEnum.IS_NOT_A_VALID_TEMPLATE);
        }
        dataMap = Maps.newLinkedHashMap();
        for (String key : map.keySet()) {
            //判断与定义的标题行是否一致
            if (!headerMap.containsKey(key)) {
                throw new ExcelImportException(ExcelImportEnum.IS_NOT_A_VALID_TEMPLATE);
            }
            dataMap.put(String.valueOf(headerMap.get(key)), map.get(key) == null ? "" : map.get(key));
        }

        StringJoiner joiner = new StringJoiner(",");
        List<Integer> pointTypeList = (List<Integer>) dbDate.get("point_type");
        if (!pointTypeList.contains(Integer.parseInt((String) dataMap.get("point_type")))) {
            joiner.add("测点类型不存在");
        }

        if (dataMap.get("uuid") == null) {
            joiner.add("设备编号不能为空");
        }

        List<Map<String, Object>> list = result.getList();
        if (list == null) {
            list = new ArrayList<>();
        }
        List<Map<String, Object>> failList = result.getFailList();
        if (failList == null) {
            failList = new ArrayList<>();
        }
        // map校验是没有行号，我们自己记录，框架是在校验过后添加行号字段的
        // 如果有标题第一条数据就是第2行，这里是有标题所以是3
        int lineNumber = list.size() + failList.size() + 3;
        dataMap.put("excelRowNum", lineNumber);
        // 添加本行数据
        list.forEach(e -> {
            if (e.get("channel_id").equals(dataMap.get("channel_id"))) {
                joiner.add("第" + e.get("excelRowNum") + "行与" + lineNumber + "行数据重复");
            }
        });
        if (!joiner.toString().isEmpty()) {
            dataMap.put("excelErrorMsg", joiner.toString());
            failList.add(dataMap);
            result.setFailList(failList);
        } else {
            list.add(dataMap);
            result.setList(list);
        }
        if (result.getFailList().size() != 0) {
            result.setVerifyFail(false);
        } else {
            result.setVerifyFail(true);
        }

    }

    /**
     * 处理完成之后的业务
     */
    @Override
    public void doAfterAll() {
        log.info("从Excel读取正确数据一共 {} 行", result.getList().size());
        log.info("从Excel读取错误数据一共 {} 行", result.getFailList().size());
    }
}