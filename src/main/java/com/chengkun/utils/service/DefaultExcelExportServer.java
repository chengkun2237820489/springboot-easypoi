package com.chengkun.utils.service;

import cn.afterturn.easypoi.handler.inter.IExcelExportServer;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * FileName: DefaultExcelExportServer
 * Author:   ck
 * Date:     2022/10/1 17:10
 * Description: 大数据导出处理
 */
@Data
public class DefaultExcelExportServer implements IExcelExportServer {

    // 导出数据
    private List<?> dataList;

    // 每次处理数据
    private Integer pageSize;

    @Override
    public List<Object> selectListForExcelExport(Object queryParams, int page) {
        // 到达最后一页结束处理
        if (((int) queryParams) < page) {
            return null;
        }
        // 对数据进行分页
        return getList(dataList, pageSize, page);
    }

    /**
     * 循环截取某页列表进行分页
     *
     * @param dataList    分页数据
     * @param pageSize    页面大小
     * @param curPage 当前页面
     */
    private static List<Object> getList(List<?> dataList, int pageSize, int curPage) {
        List<Object> currentPageList = new ArrayList<>();
        if (dataList != null && dataList.size() > 0) {
            int currIdx = (curPage > 1 ? (curPage - 1) * pageSize : 0);
            for (int i = 0; i < pageSize && i < dataList.size() - currIdx; i++) {
                currentPageList.add(dataList.get(currIdx + i));
            }
        }
        return currentPageList;
    }
}
