package com.chengkun.entity;

import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import lombok.Data;

import java.util.List;

/**
 * FileName: SaxExcelImportResult
 * Author:   ck
 * Date:     2022/10/5 13:34
 * Description: sax读取数据实体
 */
@Data
public class SaxExcelImportResult<T> extends ExcelImportResult<T> {

    /**
     * 结果集
     **/
    private List<T> list;
    /**
     * 失败数据
     */
    private List<T> failList;

    /**
     * 是否存在校验失败
     */
    private boolean verifyFail;
}
