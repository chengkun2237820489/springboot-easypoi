package com.chengkun.utils.handler;

import cn.afterturn.easypoi.handler.inter.IExcelVerifyHandler;
import cn.afterturn.easypoi.handler.inter.IReadHandler;
import com.chengkun.entity.SaxExcelImportResult;
import lombok.Data;

/**
 * FileName: DefaultSaxReadVerifyHandler
 * Author:   ck
 * Date:     2022/10/5 15:11
 * Description: 默认sax读取校验类
 */
@Data
public class DefaultSaxReadVerifyHandler<T> {

    /**
     * 返回结果集
     **/
    private SaxExcelImportResult<T> result;

    /**
     *sax读取数据处理类（包含校验）
     **/
    private IReadHandler<T> iReadHandler;

    /**
     * xls调用sax读取使用原始校验处理器
     **/
    private IExcelVerifyHandler<T> verifyHandler;
}
