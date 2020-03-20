package com.chengkun.utils.handler;
/**
 * sungrow all right reserved
 **/

import cn.afterturn.easypoi.exception.excel.enums.ExcelImportEnum;
import cn.afterturn.easypoi.handler.inter.IExcelDataHandler;
import com.chengkun.utils.exception.ExcelImportException;
import lombok.Data;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;

import java.util.Map;

/**
 * @Description easypoi使用map导入数据处理
 * @Author chengkun
 * @Date 2020/3/18 19:57
 **/
@Data
public class MapImportHandler implements IExcelDataHandler<Map<String, Object>> {

    //自定义标题行，用于校验格式是否正确
    public Map<String, Object> headerMap;

    public MapImportHandler(Map<String, Object> headerMap) {
        this.headerMap = headerMap;
    }

    /**
     * 导出处理方法
     *
     * @param obj   当前对象
     * @param name  当前字段名称
     * @param value 当前值
     * @return
     */
    @Override
    public Object exportHandler(Map<String, Object> obj, String name, Object value) {
        return null;
    }

    /**
     * 获取需要处理的字段,导入和导出统一处理了, 减少书写的字段
     *
     * @return
     */
    @Override
    public String[] getNeedHandlerFields() {
        return new String[0];
    }

    /**
     * 导入处理方法 当前对象,当前字段名称,当前值
     *
     * @param obj   当前对象
     * @param name  当前字段名称
     * @param value 当前值
     * @return
     */
    @Override
    public Object importHandler(Map<String, Object> obj, String name, Object value) {
        return null;
    }

    /**
     * 设置需要处理的属性列表
     *
     * @param fields
     */
    @Override
    public void setNeedHandlerFields(String[] fields) {

    }

    /**
     * @Description 替换map中key值
     * @Author chengkun
     * @Date 2020/3/18 20:00
     * @Param map       数据map
     * @Param originKey 原始表头key
     * @Param value     key对应的值
     * @Return void
     **/
    @Override
    public void setMapValue(Map<String, Object> map, String originKey, Object value) {
        map.put(getRealKey(originKey), value == null ? "" : value);
    }

    /**
     * 获取这个字段的 Hyperlink ,07版本需要,03版本不需要
     *
     * @param creationHelper
     * @param obj
     * @param name
     * @param value
     * @return
     */
    @Override
    public Hyperlink getHyperlink(CreationHelper creationHelper, Map<String, Object> obj, String name, Object value) {
        return null;
    }


    /**
     * @Description 进行国际化替换
     * @Author chengkun
     * @Date 2020/3/18 20:01
     * @Param originKey
     * @Return java.lang.String
     **/
    private String getRealKey(String originKey) {
        if (headerMap != null && headerMap.size() != 0) {
            if (headerMap.containsKey(originKey)) {
                return String.valueOf(headerMap.get(originKey));
            }
            //定义字段不存在，说明模板不对
            throw new ExcelImportException(ExcelImportEnum.IS_NOT_A_VALID_TEMPLATE);
        }

        return originKey;
    }
}