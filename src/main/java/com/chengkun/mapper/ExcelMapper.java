package com.chengkun.mapper;

import com.chengkun.entity.InsightIec104Mapping;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Component;

import java.util.List;
import java.util.Map;

/**
 * sungrow all right reserved
 **/
@Component
@Mapper
public interface ExcelMapper {

    /**
     * @Description 查询所有，返回实体类集合
     * @Author chengkun
     * @Date 2020/3/18 13:30
     * @Param
     * @Return java.util.List<com.chengkun.entity.InsightIec104Mapping>
     **/
    List<InsightIec104Mapping> findAll();

    /**
     * @Description 查询所有，返回map集合
     * @Author chengkun
     * @Date 2020/3/18 13:31
     * @Param
     * @Return java.util.List<java.util.Map < java.lang.String, java.lang.Object>>
     **/
    List<Map<String, Object>> findAllByMap();

    /**
     * @Description 根据测点类型查询
     * @Author chengkun
     * @Date 2020/3/18 15:42
     * @Param params
     * @Return java.util.List<java.util.Map < java.lang.String, java.lang.Object>>
     **/
    List<Map<String, Object>> findAllByPointType(Map<String, String> params);

    /**
     * @Description 根据测点类型查询, 返回实体类
     * @Author chengkun
     * @Date 2020/3/18 16:18
     * @Param params
     * @Return java.util.List<com.chengkun.entity.InsightIec104Mapping>
     **/
    List<InsightIec104Mapping> findAllByPointTypeEntity(Map<String, String> params);
}
