package com.chengkun.controller;
/**
 * sungrow all right reserved
 **/

import cn.afterturn.easypoi.entity.vo.NormalExcelConstants;
import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import cn.afterturn.easypoi.excel.export.ExcelExportService;
import cn.afterturn.easypoi.handler.inter.IReadHandler;
import cn.afterturn.easypoi.util.PoiMergeCellUtil;
import com.alibaba.fastjson.JSONObject;
import com.chengkun.entity.InsightIec104Mapping;
import com.chengkun.service.ExcelService;
import com.chengkun.utils.*;
import com.chengkun.utils.handler.DefaultMapImportHandler;
import com.chengkun.utils.style.DefaultExportExcelStyle;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

/**
 * @Description Excel导入导出Controller
 * @Author chengkun
 * @Date 2020/3/17 14:42
 **/
@Api(tags = "Excel导入导出控制器")
@RestController
@RequestMapping("/excel")
@Log4j2
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @ApiOperation(value = "Excel导出")
    @GetMapping("/exportExcel")
    public void exportExcel(HttpServletResponse response) {
        System.out.println(1);
        // 模拟从数据库获取需要导出的数据
        List<InsightIec104Mapping> personList = excelService.findAll();
        // 导出操作
        ExcelUtils.exportExcel(personList, "easypoi导出功能", "Export", InsightIec104Mapping.class, "测试.xls", response);
    }

    @ApiOperation(value = "Excel导出样式测试")
    @GetMapping("/exportExcelForStyle")
    public void exportExcelForStyle(HttpServletResponse response) {
        System.out.println(1);
        // 模拟从数据库获取需要导出的数据
        List<InsightIec104Mapping> personList = excelService.findAll();
        // 导出操作
        //exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName, Class<?> style, HttpServletResponse response)
        ExcelUtils.exportExcel(personList, "easypoi导出样式功能", "Export", InsightIec104Mapping.class, "excel样式测试.xls", DefaultExportExcelStyle.class, response);
    }

    @ApiOperation(value = "使用map导出")
    @GetMapping("/exportExcelForMap")
    public void exportExcelForMap(HttpServletResponse response) {
        System.out.println(1);
        // 模拟从数据库获取需要导出的数据
        List<Map<String, Object>> list = excelService.findAllByMap();
        //定义 excel 导出工具类集合
        List<ExcelExportEntity> colList = new ArrayList<>();
        // 构造对象等同于@Excel
        ExcelExportEntity exportEntity = new ExcelExportEntity("通道号", "channel_id", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("设备编号", "uuid", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("测点类型", "point_type", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("测点名称", "point_name", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("IEC测点", "iec_point", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("设备测点", "point_id", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("状态(1：启用，0：停用)", "is_enable", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("修正系数", "iec_parameter", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("偏移量", "iec_offset", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("取反(1：是，0：否)", "iec_negate", 30);
        colList.add(exportEntity);

        // 把我们构造好的bean对象放到params就可以了
        ExcelUtils.exportExcel(list, "easypoi导出map功能", "Export", colList, "easypoi导出map功能.xls", DefaultExportExcelStyle.class, response);
    }

    @ApiOperation(value = "Excel导出合并单元格")
    @GetMapping("/exportExcelForMergeCells")
    public void exportExcelForMergeCells(HttpServletResponse response) {
        System.out.println(1);
        // 模拟从数据库获取需要导出的数据
        List<InsightIec104Mapping> personList = excelService.findAll();
        // 导出操作
        Workbook workbook = ExcelExportUtil.exportExcel(new ExportParams("easypoi导出功能", "Export"), InsightIec104Mapping.class, personList);
        Sheet sheet = workbook.getSheetAt(0);
        /**合并单元格**/
        PoiMergeCellUtil.mergeCells(sheet, 1, 8, 0, 1);
        if (workbook != null) {
            ExcelUtils.downLoadExcel("合并单元格测试.xls", response, workbook);
        }

    }

    @ApiOperation(value = "Excel导出复合表头")
    @GetMapping("/exportExcelForComposite")
    public void exportExcelForComposite(HttpServletResponse response) {
        System.out.println(1);
        //表头设置
        List<ExcelExportEntity> colList = new ArrayList<>();

        ExcelExportEntity colEntity = new ExcelExportEntity("日期", "dt");
        colEntity.setNeedMerge(true);
        colList.add(colEntity);

        colEntity = new ExcelExportEntity("产品PV", "pv");
        colEntity.setNeedMerge(true);
        colList.add(colEntity);

        colEntity = new ExcelExportEntity("产品UV", "uv");
        colEntity.setNeedMerge(true);
        colList.add(colEntity);

        ExcelExportEntity group_1 = new ExcelExportEntity("业务数据", "businessData");
        List<ExcelExportEntity> exportEntities = new ArrayList<>();
        for (int i = 1; i < 5; i++) {
            exportEntities.add(new ExcelExportEntity("数据" + i, "data" + i));
        }
        group_1.setList(exportEntities);
        colList.add(group_1);

        //文件数据
        List<Map<String, Object>> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Map<String, Object> valMap = new HashMap<String, Object>();
            valMap.put("dt", "日期" + i);
            valMap.put("pv", "pv" + i);
            valMap.put("uv", "uv" + i);
            List<Map<String, Object>> list_1 = new ArrayList<Map<String, Object>>();
            Map<String, Object> valMap_1 = new HashMap<String, Object>();
            for (int j = 1; j < 5; j++) {
                valMap_1.put("data" + j, "数据" + j);
            }
            list_1.add(valMap_1);
            valMap.put("businessData", list_1);
            list.add(valMap);
        }
        // 把我们构造好的bean对象放到params就可以了
        ExcelUtils.exportExcel(list, "easypoi导出复合标题", "Export", colList, "easypoi导出复合标题.xls", DefaultExportExcelStyle.class, response);
    }

    @ApiOperation(value = "Excel导出多sheet页,自定义标题")
    @GetMapping("/exportExcelForSheets")
    public void exportExcelForSheets(HttpServletResponse response) {
        System.out.println(1);
        // 模拟从数据库获取需要导出的数据
        //sheet1 遥信
        List<Map<String, Object>> list1 = excelService.findAllByPointType(ImmutableMap.of("point_type", "1"));
        List<Map<String, Object>> list2 = excelService.findAllByPointType(ImmutableMap.of("point_type", "2"));
        //sheet2  遥测
        //定义 excel 导出工具类集合
        List<ExcelExportEntity> colList = new ArrayList<>();
        // 构造对象等同于@Excel
        ExcelExportEntity exportEntity = new ExcelExportEntity("id", "id", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("通道号", "channel_id", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("设备编号", "uuid", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("测点类型", "point_type", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("测点名称", "point_name", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("IEC测点", "iec_point", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("设备测点", "point_id", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("状态(1：启用，0：停用)", "is_enable", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("修正系数", "iec_parameter", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("偏移量", "iec_offset", 30);
        colList.add(exportEntity);
        exportEntity = new ExcelExportEntity("取反(1：是，0：否)", "iec_negate", 30);
        colList.add(exportEntity);

        // 创建参数对象
        ExportParams exportParams1 = new ExportParams("遥信", "遥信");
        exportParams1.setStyle(DefaultExportExcelStyle.class);
        ExportParams exportParams2 = new ExportParams("遥测", "遥测");
        exportParams2.setStyle(DefaultExportExcelStyle.class);

        // sheet1设置
        Map<String, Object> sheet1ExportMap = new HashMap<>();
        sheet1ExportMap.put(NormalExcelConstants.CLASS, ExcelExportEntity.class);
        sheet1ExportMap.put(NormalExcelConstants.DATA_LIST, list1);
        sheet1ExportMap.put(NormalExcelConstants.PARAMS, exportParams1);
        //这边为了方便，sheet1和sheet2用同一个表头(实际使用中可自行调整)
        sheet1ExportMap.put(NormalExcelConstants.MAP_LIST, colList);

        //Sheet2设置
        Map<String, Object> sheet2ExportMap = new HashMap<>();
        sheet2ExportMap.put(NormalExcelConstants.CLASS, ExcelExportEntity.class);
        sheet2ExportMap.put(NormalExcelConstants.DATA_LIST, list2);
        sheet2ExportMap.put(NormalExcelConstants.PARAMS, exportParams2);
        //这边为了方便，sheet1和sheet2用同一个表头(实际使用中可自行调整)
        sheet2ExportMap.put(NormalExcelConstants.MAP_LIST, colList);

        // 将sheet1、sheet2使用得map进行包装
        List<Map<String, Object>> sheetsList = new ArrayList<>();
        sheetsList.add(sheet1ExportMap);
        sheetsList.add(sheet2ExportMap);

        // 执行方法
        Workbook workbook = new HSSFWorkbook();

        for (Map<String, Object> map : sheetsList) {
            ExcelExportService server = new ExcelExportService();
            ExportParams param = (ExportParams) map.get("params");
            @SuppressWarnings("unchecked")
            List<ExcelExportEntity> entity = (List<ExcelExportEntity>) map.get("mapList");
            Collection<?> data = (Collection<?>) map.get("data");
            server.createSheetForMap(workbook, param, entity, data);
        }
        if (workbook != null) {
            ExcelUtils.downLoadExcel("多sheet页导出.xls", response, workbook);
        }
    }

    @ApiOperation(value = "Excel导出多sheet页,使用实体类")
    @GetMapping("/exportExcelForSheets1")
    public void exportExcelForSheets1(HttpServletResponse response) {
        System.out.println(1);
        // 模拟从数据库获取需要导出的数据
        //sheet1 遥信
        List<InsightIec104Mapping> list1 = excelService.findAllByPointTypeEntity(ImmutableMap.of("point_type", "1"));
        List<InsightIec104Mapping> list2 = excelService.findAllByPointTypeEntity(ImmutableMap.of("point_type", "2"));

        // 创建参数对象
        ExportParams exportParams1 = new ExportParams("遥信", "遥信");
        ExportParams exportParams2 = new ExportParams("遥测", "遥测");

        // 创建sheet1使用得map
        Map<String, Object> deptDataMap = new HashMap<>(4);
        // title的参数为ExportParams类型
        deptDataMap.put("title", exportParams1);
        // 模版导出对应得实体类型
        deptDataMap.put("entity", InsightIec104Mapping.class);
        // sheet中要填充得数据
        deptDataMap.put("data", list1);

        // 创建sheet2使用得map
        Map<String, Object> userDataMap = new HashMap<>(4);
        userDataMap.put("title", exportParams2);
        userDataMap.put("entity", InsightIec104Mapping.class);
        userDataMap.put("data", list2);
        // 将sheet1和sheet2使用得map进行包装
        List<Map<String, Object>> sheetsList = new ArrayList<>();
        sheetsList.add(deptDataMap);
        sheetsList.add(userDataMap);

        // 把我们构造好的bean对象放到params就可以了
        // 执行方法
        Workbook workbook = ExcelExportUtil.exportExcel(sheetsList, ExcelType.HSSF);
        if (workbook != null) {
            ExcelUtils.downLoadExcel("多sheet页导出", response, workbook);
        }
    }

    @ApiOperation(value = "Excel导入")
    @GetMapping("/importExcel")
    public String importExcel(HttpServletResponse response) {
        ImportParams importParams = new ImportParams();
        // 数据处理
        importParams.setHeadRows(1);
        importParams.setTitleRows(1);
        // 需要验证
        importParams.setNeedVerify(true);
        File file = new File("E:\\Chrom Downloads\\IEC测点映射20200317102927.xls");
        try {
            ExcelImportResult<InsightIec104Mapping> result = ExcelImportUtil.importExcelMore(new FileInputStream(file), InsightIec104Mapping.class,
                    importParams);
            List<InsightIec104Mapping> list = result.getList();
            for (InsightIec104Mapping iec : list) {
                log.info("从Excel导入数据到数据库的详细为 ：{}", JSONObject.toJSONString(iec));
                //TODO 将导入的数据做保存数据库操作
            }
            log.info("从Excel导入数据一共 {} 行 ", list.size());
        } catch (IOException e) {
            log.error("导入失败：{}", e.getMessage());
        } catch (Exception e1) {
            log.error("导入失败：{}", e1.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Excel使用map导入")
    @GetMapping("/importExcelForMap")
    public String importExcelForMap(HttpServletResponse response) {
        ImportParams importParams = new ImportParams();
        // 数据处理
        importParams.setHeadRows(1);
        importParams.setTitleRows(1);
        // 需要验证
        importParams.setNeedVerify(true);
        importParams.setImportFields(new String[]{"通道号", "测点类型", "测点名称", "IEC测点", "设备测点", "状态(1：启用，0：停用)", "修正系数", "偏移量", "取反(1：是，0：否)"});
        importParams.setDataHandler(new DefaultMapImportHandler(Maps.newLinkedHashMap()));
        try {
            long start = System.currentTimeMillis() / 1000;
            ExcelImportResult<Object> result = ExcelImportUtil.importExcelMore(
                    new FileInputStream("E:\\Chrom Downloads\\IEC测点映射20200317102927.xls"), Map.class, importParams);
            long end = System.currentTimeMillis() / 1000;
            log.info("从Excel导入数据一共 {} 行 ，消耗时间{}秒", result.getList().size(), end - start);
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Excel使用sax导入,使用实体类")
    @GetMapping("/importExcelBySaxForPojo")
    public String importExcelBySaxForPojo(HttpServletResponse response) {
        ImportParams importParams = new ImportParams();
        // 数据处理
        importParams.setHeadRows(1);
        importParams.setTitleRows(1);

        List<InsightIec104Mapping> result = Lists.newArrayList();
        try {
            long start = System.currentTimeMillis() / 1000;
            ExcelImportUtil.importExcelBySax(
                    new FileInputStream(
                            new File("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx")),
                    InsightIec104Mapping.class, importParams, new IReadHandler<InsightIec104Mapping>() {

                        /**
                         * 处理解析对象
                         *
                         * @param insightIec104Mapping
                         */
                        @Override
                        public void handler(InsightIec104Mapping insightIec104Mapping) {
                            result.add(insightIec104Mapping);
                        }

                        /**
                         * 处理完成之后的业务
                         */
                        @Override
                        public void doAfterAll() {
                            log.info("从Excel导入数据一共 {} 行", result.size());
                        }
                    });
            long end = System.currentTimeMillis() / 1000;
            log.info("消耗时间{}秒", end - start);
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Excel使用sax导入,map导入")
    @GetMapping("/importExcelBySax")
    public String importExcelBySax(HttpServletResponse response) {
        ImportParams importParams = new ImportParams();
        // 数据处理
        importParams.setHeadRows(1);
        importParams.setTitleRows(1);

        List<Map<String, Object>> result = Lists.newArrayList();
        try {
            long start = System.currentTimeMillis() / 1000;
            ExcelImportUtil.importExcelBySax(new FileInputStream(new File("E:\\google_downloads\\06DCE110.xlsx")), Map.class, importParams, new IReadHandler<Map<String, Object>>() {

                @Override
                public void handler(Map<String, Object> map) {
                    Map<String, Object> stringObjectMap = Maps.newLinkedHashMap();
                    for (String key : map.keySet()) {
                        String value = (String) map.get(key);
                        if ("设备编号".equals(key)) {
                            stringObjectMap.put("uuid", "null".equals(value) ? "" : value);
                        }
                        if ("所属上级设备编号".equals(key)) {
                            stringObjectMap.put("up_uuid", "null".equals(value) ? "" : value);
                        }
                        if ("通道号".equals(key)) {
                            stringObjectMap.put("channel_id", "null".equals(value) ? "" : value);
                        }
                        if ("设备名称".equals(key)) {
                            stringObjectMap.put("device_name", "null".equals(value) ? "" : value);
                        }
                        if ("所属设备类型".equals(key)) {
                            stringObjectMap.put("device_type", "null".equals(value) ? "" : value);
                        }
                        if ("所属设备型号".equals(key)) {
                            stringObjectMap.put("device_model", "null".equals(value) ? "" : value);
                        }
                        if ("所属电站编码".equals(key)) {
                            stringObjectMap.put("ps_id", "null".equals(value) ? "" : value);
                        }
                        if ("装机功率".equals(key)) {
                            stringObjectMap.put("install_power", "null".equals(value) ? "" : value);
                        }
                        if ("额定功率".equals(key)) {
                            stringObjectMap.put("rate_power", "null".equals(value) ? "" : value);
                        }
                    }
                    result.add(stringObjectMap);
                }

                @Override
                public void doAfterAll() {
                    log.info("从Excel导入数据一共 {} 行", result.size());
                }
            });
            long end = System.currentTimeMillis() / 1000;
            log.info("消耗时间{}秒", end - start);
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }
}