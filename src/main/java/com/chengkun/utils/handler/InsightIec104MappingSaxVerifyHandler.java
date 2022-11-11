package com.chengkun.utils.handler;

import cn.afterturn.easypoi.handler.inter.IReadHandler;
import com.chengkun.entity.InsightIec104Mapping;
import com.chengkun.entity.SaxExcelImportResult;
import lombok.Data;
import lombok.extern.log4j.Log4j2;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.StringJoiner;

/**
 * FileName: InsightIec104MappingSaxVerifyHandler
 * Author:   ck
 * Date:     2022/10/5 14:00
 * Description: InsightIec104Mapping的sax读取校验类
 */
@Data
@Log4j2
public class InsightIec104MappingSaxVerifyHandler implements IReadHandler<InsightIec104Mapping> {

    private SaxExcelImportResult<InsightIec104Mapping> result;

    private Map<String, Object> dbDate;

    @Override
    public void handler(InsightIec104Mapping obj) {
        StringJoiner joiner = new StringJoiner(",");
        List<Integer> pointTypeList = (List<Integer>) dbDate.get("point_type");
        if (!pointTypeList.contains(Integer.parseInt(obj.getPointType()))) {
            joiner.add("测点类型不存在");
        }

        if (obj.getUuid() == null) {
            joiner.add("设备编号不能为空");
        }

        List<InsightIec104Mapping> list = result.getList();
        if (list == null) {
            list = new ArrayList<>();
        }
        List<InsightIec104Mapping> failList = result.getFailList();
        if (failList == null) {
            failList = new ArrayList<>();
        }
        // 添加本行数据
        obj.setRowNum(list.size() + failList.size() + 2);
        list.forEach(e -> {
            if (e.equals(obj)) {
                joiner.add("第" + e.getRowNum() + "行与" + obj.getRowNum() + "行数据重复");
            }
        });
        if (!joiner.toString().isEmpty()) {
            obj.setErrorMsg(joiner.toString());
            failList.add(obj);
            result.setFailList(failList);
        } else {
            list.add(obj);
            result.setList(list);
        }
        if (result.getFailList().size() != 0) {
            result.setVerifyFail(false);
        } else {
            result.setVerifyFail(true);
        }
    }

    @Override
    public void doAfterAll() {
        log.info("从Excel读取正确数据一共 {} 行", result.getList().size());
        log.info("从Excel读取错误数据一共 {} 行", result.getFailList().size());
    }
}
