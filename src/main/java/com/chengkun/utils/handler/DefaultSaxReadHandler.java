package com.chengkun.utils.handler;

import cn.afterturn.easypoi.handler.inter.IReadHandler;
import lombok.Data;
import lombok.extern.log4j.Log4j2;

import java.util.List;

/**
 * FileName: DefaultSaxReadHandler
 * Author:   ck
 * Date:     2022/10/2 16:20
 * Description: 实体类导入sax读取数据处理类
 */
@Data
@Log4j2
public class DefaultSaxReadHandler<T> implements IReadHandler<T> {

    // 存储读取的数据
    private List<T> dataList;

    @Override
    public void handler(T o) {
        dataList.add(o);
    }

    @Override
    public void doAfterAll() {
        log.info("从Excel导入数据一共 {} 行", dataList.size());
    }
}
