package com.chengkun.utils.handler;
/**
 * sungrow all right reserved
 **/

import cn.afterturn.easypoi.exception.excel.enums.ExcelImportEnum;
import cn.afterturn.easypoi.handler.inter.IReadHandler;
import com.chengkun.utils.exception.ExcelImportException;
import com.google.common.collect.Maps;
import lombok.Data;
import lombok.extern.log4j.Log4j2;

import java.util.List;
import java.util.Map;

/**
 * @Description map导入使用sax数据处理
 * @Author chengkun
 * @Date 2020/3/19 13:25
 **/
@Log4j2
@Data
public class MapIReadHandler implements IReadHandler<Map<String, Object>> {

    //解析的数据
    public List<Map<String, Object>> result;

    //自定义标题行，用于校验格式是否正确
    public Map<String, Object> headerMap;

    //数据map，将标题替换为对应key
    private Map<String, Object> dataMap;

    public MapIReadHandler(List<Map<String, Object>> result, Map<String, Object> headerMap) {
        this.result = result;
        this.headerMap = headerMap;
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
        result.add(dataMap);
    }

    /**
     * 处理完成之后的业务
     */
    @Override
    public void doAfterAll() {
        log.info("从Excel导入数据一共 {} 行", result.size());
    }
}