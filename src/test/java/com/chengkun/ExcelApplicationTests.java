package com.chengkun;
/**
 * sungrow all right reserved
 **/

import com.chengkun.entity.InsightIec104Mapping;
import com.chengkun.mapper.ExcelMapper;
import com.google.common.collect.ImmutableMap;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @Description 测试类
 * @Author chengkun
 * @Date 2020/3/17 15:19
 **/

@RunWith(SpringRunner.class)
@SpringBootTest(classes = ExcelApplication.class)
public class ExcelApplicationTests {

    @Autowired
    ExcelMapper excelMapper;

    /**
     * @Description 查询所有返回实体类集合
     * @Author chengkun
     * @Date 2020/3/18 9:00
     * @Param
     * @Return void
     **/
    @Test
    public void test1() {
        List<InsightIec104Mapping> list = excelMapper.findAll();
        System.out.println(list);
    }

    /**
     * @Description 查询所有返回map集合
     * @Author chengkun
     * @Date 2020/3/18 9:00
     * @Param
     * @Return void
     **/
    @Test
    public void test2() {
        List<Map<String, Object>> list = excelMapper.findAllByMap();
        System.out.println(list);
    }

    @Test
    public void test3() {
        List<Map<String, Object>> list = excelMapper.findAllByPointType(ImmutableMap.of("point_type", "1"));
        System.out.println(list);
    }

    @Test
    public void test4() {
        List<InsightIec104Mapping> list = new ArrayList<>();
        for (int i = 0; i < 20; i++) {
            InsightIec104Mapping iec104Mapping = new InsightIec104Mapping();
            iec104Mapping.setChannelId("" + i);
            iec104Mapping.setIecNegate("" + i);
            iec104Mapping.setIecOffset("" + i);
            iec104Mapping.setIecParameter("" + i);
            iec104Mapping.setIecPoint(i);
            iec104Mapping.setIsEnable("" + i);
            iec104Mapping.setPointId(i);
            iec104Mapping.setUuid(i);
            iec104Mapping.setPointName("test" + i);
            iec104Mapping.setPointType("" + i);
            list.add(iec104Mapping);
        }
        excelMapper.insertPointList(list);
    }
}