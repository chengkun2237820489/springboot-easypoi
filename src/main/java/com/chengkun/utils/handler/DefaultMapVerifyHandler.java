package com.chengkun.utils.handler;

import cn.afterturn.easypoi.excel.entity.result.ExcelVerifyHandlerResult;
import cn.afterturn.easypoi.handler.inter.IExcelVerifyHandler;
import com.chengkun.entity.InsightIec104Mapping;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.StringJoiner;

/**
 * FileName: MapVerifyHandler
 * Author:   ck
 * Date:     2022/10/3 17:27
 * Description: map数据校验处理器
 */
@Data
public class DefaultMapVerifyHandler implements IExcelVerifyHandler<Map<String, Object>> {

    // 存放数据，比对和之前数据是否有重复
    private final ThreadLocal<List<Map<String, Object>>> threadLocal = new ThreadLocal<>();

    private Map<String, Object> dbDate;

    @Override
    public ExcelVerifyHandlerResult verifyHandler(Map<String, Object> obj) {
        // 缓存数据校验是否有重复数据，同是记录行号
        List<Map<String, Object>> threadLocalVal = threadLocal.get();
        if (threadLocalVal == null) {
            threadLocalVal = new ArrayList<>();
        }
        // map校验是没有行号，我们自己记录，框架是在校验过后添加行号字段的
        // 如果有标题第一条数据就是第2行，这里是有标题所以是3
        int lineNumber = threadLocalVal.size() + 3;
        obj.put("excelRowNum", lineNumber);
        StringJoiner joiner = new StringJoiner(",");
        List<Integer> pointTypeList = (List<Integer>) dbDate.get("point_type");
        if (!pointTypeList.contains(Integer.parseInt((String) obj.get("point_type")))) {
            joiner.add("第" + lineNumber + "行测点类型不存在");
        }

        if (obj.get("uuid") == null) {
            joiner.add("第" + lineNumber + "行设备编号不能为空");
        }

        threadLocalVal.forEach(e -> {
            if (e.get("channel_id").equals(obj.get("channel_id"))) {
                joiner.add("第" + e.get("excelRowNum") + "行与" + lineNumber + "行数据重复");
            }
        });
        // 添加本行数据对象到ThreadLocal中
        threadLocalVal.add(obj);
        threadLocal.set(threadLocalVal);
        if (!joiner.toString().isEmpty()) {
            return new ExcelVerifyHandlerResult(false, joiner.toString());
        } else {
            return new ExcelVerifyHandlerResult(true, "导入成功");
        }
    }

    public ThreadLocal<List<Map<String, Object>>> getThreadLocal() {
        return threadLocal;
    }
}
