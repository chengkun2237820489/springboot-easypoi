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
 * FileName: InsightIec104MappingVerifyHandler
 * Author:   ck
 * Date:     2022/10/1 14:10
 * Description: InsightIec104Mapping校验处理器
 */
@Data
public class InsightIec104MappingVerifyHandler implements IExcelVerifyHandler<InsightIec104Mapping> {

    // 存放数据，比对和之前数据是否有重复
    private final ThreadLocal<List<InsightIec104Mapping>> threadLocal = new ThreadLocal<>();

    private Map<String, Object> dbDate;

    @Override
    public ExcelVerifyHandlerResult verifyHandler(InsightIec104Mapping obj) {
        StringJoiner joiner = new StringJoiner(",");
        List<Integer> pointTypeList = (List<Integer>) dbDate.get("point_type");
        if (!pointTypeList.contains(Integer.parseInt(obj.getPointType()))) {
            joiner.add("第" + obj.getRowNum() + "行测点类型不存在");
        }

        if (obj.getUuid() == null) {
            joiner.add("第" + obj.getRowNum() + "行设备编号不能为空");
        }

        List<InsightIec104Mapping> threadLocalVal = threadLocal.get();
        if (threadLocalVal == null) {
            threadLocalVal = new ArrayList<>();
        }

        threadLocalVal.forEach(e -> {
            if (e.equals(obj)) {
                joiner.add("数据与第" + e.getRowNum() + "行重复");
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

    public ThreadLocal<List<InsightIec104Mapping>> getThreadLocal() {
        return threadLocal;
    }
}
