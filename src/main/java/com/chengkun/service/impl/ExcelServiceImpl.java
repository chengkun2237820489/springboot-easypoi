package com.chengkun.service.impl;
/**
 * sungrow all right reserved
 **/

import com.chengkun.entity.InsightIec104Mapping;
import com.chengkun.mapper.ExcelMapper;
import com.chengkun.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Map;

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
}