package com.chengkun.controller;
/**
 * sungrow all right reserved
 **/

import cn.afterturn.easypoi.util.PoiMergeCellUtil;
import com.alibaba.fastjson.JSONObject;
import com.chengkun.entity.InsightIec104Mapping;
import com.chengkun.service.ExcelService;
import com.chengkun.utils.EasyPoiUtils;
import com.chengkun.utils.style.ExportExcelStyle;
import com.google.common.collect.*;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import lombok.extern.log4j.Log4j2;
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
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @Description 自定义工具类控制器
 * 在ExcelController基础做修改，更方便开发
 * @Author chengkun
 * @Date 2020/3/20 9:21
 **/
@Api(tags = "ExcelUtils导入导出控制器")
@RestController
@RequestMapping("/excelutil")
@Log4j2
public class ExcelUtilController {
    @Autowired
    private ExcelService excelService;

    @ApiOperation(value = "Excel实体类导出")
    @GetMapping("/exportExcelForPojo")
    public void exportExcel(HttpServletResponse response) {
        // 模拟从数据库获取需要导出的数据
        List<InsightIec104Mapping> list = excelService.findAll();
        // 导出操作
        EasyPoiUtils.exportExcel(list, "easypoi实体类导出", "Export", InsightIec104Mapping.class, "测试.xls", null, response);
    }

    @ApiOperation(value = "Excel导出样式测试")
    @GetMapping("/exportExcelByStyle")
    public void exportExcelByStyle(HttpServletResponse response) {
        // 模拟从数据库获取需要导出的数据
        List<InsightIec104Mapping> personList = excelService.findAll();
        // 导出操作
        //exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName, Class<?> style, HttpServletResponse response)
        EasyPoiUtils.exportExcel(personList, "easypoi导出样式功能", "Export", InsightIec104Mapping.class, "excel样式测试.xls", null, ExportExcelStyle.class, response);
    }

    @ApiOperation(value = "使用map导出")
    @GetMapping("/exportExcelByMap")
    public void exportExcelByMap(HttpServletResponse response) {
        // 模拟从数据库获取需要导出的数据
        List<Map<String, Object>> list = excelService.findAllByMap();
        Map<String, Object> headerMap = Maps.newLinkedHashMap(); //定义标题行
        headerMap.put("通道号", "channel_id");
        headerMap.put("设备编号", "uuid");
        headerMap.put("测点类型", "point_type");
        headerMap.put("测点名称", "point_name");
        headerMap.put("IEC测点", "iec_point");
        headerMap.put("设备测点", "point_id");
        headerMap.put("状态(1：启用，0：停用)", "is_enable");
        headerMap.put("修正系数", "iec_parameter");
        headerMap.put("偏移量", "iec_offset");
        headerMap.put("取反(1：是，0：否)", "iec_negate");
        EasyPoiUtils.exportExcelForMap(list, "easypoi导出map功能", "Export", headerMap, "easypoi导出map功能.xlsx", null, ExportExcelStyle.class, response);

    }

    @ApiOperation(value = "Excel导出合并单元格")
    @GetMapping("/exportExcelByMergeCells")
    public void exportExcelByMergeCells(HttpServletResponse response) {
        System.out.println(1);
        // 模拟从数据库获取需要导出的数据
        List<InsightIec104Mapping> personList = excelService.findAll();
        // 导出操作 //List<?> list, String title, String sheetName, Class<?> pojoClass,  Class<?> style
        Workbook workbook = EasyPoiUtils.exportExcel(personList, "easypoi导出功能", "Export", InsightIec104Mapping.class, ExportExcelStyle.class);
        Sheet sheet = workbook.getSheetAt(0);
        /**合并单元格**/
        PoiMergeCellUtil.mergeCells(sheet, 1, 8, 0, 1);
        if (workbook != null) {
            EasyPoiUtils.downLoadExcel("合并单元格测试.xls", null, response, workbook);
        }

    }

    @ApiOperation(value = "Excel导出复合表头")
    @GetMapping("/exportExcelByComposite")
    public void exportExcelByComposite(HttpServletResponse response) {
        Table<String, String, Object> table = HashBasedTable.create();
        table.put("日期", "dt", Maps.newLinkedHashMap());
        table.put("产品PV", "pv", Maps.newLinkedHashMap());
        Map<String, Object> groupMap = Maps.newLinkedHashMap();
        for (int i = 0; i < 5; i++) {
            groupMap.put("数据" + i, "data" + i);
        }
        table.put("业务数据", "businessData", groupMap);

        //文件数据
        List<Map<String, Object>> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Map<String, Object> valMap = new HashMap<String, Object>();
            valMap.put("dt", "日期" + i);
            valMap.put("pv", "pv" + i);
            valMap.put("uv", "uv" + i);
            List<Map<String, Object>> list_1 = new ArrayList<Map<String, Object>>();
            Map<String, Object> valMap_1 = new HashMap<String, Object>();
            for (int j = 0; j < 5; j++) {
                valMap_1.put("data" + j, "数据" + j);
            }
            list_1.add(valMap_1);
            valMap.put("businessData", list_1);
            list.add(valMap);
        }
        //List<?> list, String title, String sheetName, Table<String, String, Object> heaters, String fileName, String filePath, Class<?> style, HttpServletResponse response
        EasyPoiUtils.exportHeadersExcelForMap(list, "easypoi导出复合标题", "Export", table, "easypoi导出复合标题.xls", null, ExportExcelStyle.class, response);
    }

    @ApiOperation(value = "Excel导出多sheet页,自定义标题")
    @GetMapping("/exportExcelBySheetsForMap")
    public void exportExcelBySheets(HttpServletResponse response) {
        // 模拟从数据库获取需要导出的数据
        //sheet1 遥信
        List<Map<String, Object>> list1 = excelService.findAllByPointType(ImmutableMap.of("point_type", "1"));
        //sheet2  遥测
        List<Map<String, Object>> list2 = excelService.findAllByPointType(ImmutableMap.of("point_type", "2"));
        //定义 sheet1内容
        List<Map<String, Object>> sheets = Lists.newArrayList(); //sheet页集合
        Map<String, Object> sheet1 = Maps.newLinkedHashMap(); //sheet1页内容
        Map<String, Object> headerMap = Maps.newLinkedHashMap(); //定义标题行
        headerMap.put("通道号", "channel_id");
        headerMap.put("设备编号", "uuid");
        headerMap.put("测点类型", "point_type");
        headerMap.put("测点名称", "point_name");
        headerMap.put("IEC测点", "iec_point");
        headerMap.put("设备测点", "point_id");
        headerMap.put("状态(1：启用，0：停用)", "is_enable");
        headerMap.put("修正系数", "iec_parameter");
        headerMap.put("偏移量", "iec_offset");
        headerMap.put("取反(1：是，0：否)", "iec_negate");
        sheet1.put("header", headerMap); //sheet1标题行
        sheet1.put("title", "遥信");   //sheet1表头
        sheet1.put("sheetName", "遥信"); //sheet1工作簿名称
        sheet1.put("dataList", list1);//sheet1数据
        sheets.add(sheet1);

        //定义 sheet2内容
        Map<String, Object> sheet2 = Maps.newLinkedHashMap(); //sheet2页内容
        sheet2.put("header", headerMap); //sheet2标题行
        sheet2.put("title", "遥测");   //sheet2表头
        sheet2.put("sheetName", "遥测"); //sheet2工作簿名称
        sheet2.put("dataList", list2);//sheet2数据
        sheets.add(sheet2);
        //List<Map<String, Object>> sheets, String fileName, String filePath, Class<?> style, HttpServletResponse response
        EasyPoiUtils.exportSheetsExcelForMap(sheets, "多sheet页map导出.xlsx", null, ExportExcelStyle.class, response);
    }

    @ApiOperation(value = "Excel导出多sheet页,使用实体类")
    @GetMapping("/exportExcelBySheetsForPojo")
    public void exportExcelBySheets1(HttpServletResponse response) {
        System.out.println(1);
        // 模拟从数据库获取需要导出的数据
        //sheet1 遥信
        List<InsightIec104Mapping> list1 = excelService.findAllByPointTypeEntity(ImmutableMap.of("point_type", "1"));
        List<InsightIec104Mapping> list2 = excelService.findAllByPointTypeEntity(ImmutableMap.of("point_type", "2"));

        // 创建参数对象
        //定义 sheet1内容
        List<Map<String, Object>> sheets = Lists.newArrayList(); //sheet页集合
        Map<String, Object> sheet1 = Maps.newLinkedHashMap(); //sheet1页内容

        sheet1.put("header", InsightIec104Mapping.class); //sheet1标题行
        sheet1.put("title", "遥信");   //sheet1表头
        sheet1.put("sheetName", "遥信"); //sheet1工作簿名称
        sheet1.put("dataList", list1);//sheet1数据
        sheets.add(sheet1);

        //定义 sheet2内容
        Map<String, Object> sheet2 = Maps.newLinkedHashMap(); //sheet2页内容
        sheet2.put("header", InsightIec104Mapping.class); //sheet1标题行
        sheet2.put("title", "遥测");   //sheet2表头
        sheet2.put("sheetName", "遥测"); //sheet2工作簿名称
        sheet2.put("dataList", list2);//sheet2数据
        sheets.add(sheet2);
        //List<Map<String, Object>> sheets, String fileName, String filePath, Class<?> style, HttpServletResponse response
        EasyPoiUtils.exportSheetsExcelForPoJo(sheets, "多sheet页pojo导出.xlsx", null, ExportExcelStyle.class, response);
    }

    @ApiOperation(value = "Excel导入")
    @GetMapping("/importExcel")
    public String importExcel(HttpServletResponse response) {
        Map<String, Object> params = Maps.newLinkedHashMap();
        params.put("headRows", 1); //表头行数
        params.put("titleRows", 1); //表格标题行数
        //导入时校验数据标题
        Map<String, Object> headerMap = Maps.newLinkedHashMap(); //定义标题行
        headerMap.put("通道号", "channel_id");
        headerMap.put("设备编号", "uuid");
        headerMap.put("测点类型", "point_type");
        headerMap.put("测点名称", "point_name");
        headerMap.put("IEC测点", "iec_point");
        headerMap.put("设备测点", "point_id");
        headerMap.put("状态(1：启用，0：停用)", "is_enable");
        headerMap.put("修正系数", "iec_parameter");
        headerMap.put("偏移量", "iec_offset");
        headerMap.put("取反(1：是，0：否)", "iec_negate");
        params.put("header", headerMap);

        File file = new File("E:\\google_downloads\\IEC测点映射20200317102927.xls");
        try {
            Map<String, Object> result = EasyPoiUtils.importExcel(new FileInputStream(file), InsightIec104Mapping.class, params);
            List<InsightIec104Mapping> list = (List<InsightIec104Mapping>) result.get("dataList");
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
    @GetMapping("/importExcelByMap")
    public String importExcelByMap(HttpServletResponse response) {
        Map<String, Object> params = Maps.newLinkedHashMap();
        params.put("headRows", 1); //表头行数
        params.put("titleRows", 1); //表格标题行数
        //导入时校验数据标题
        Map<String, Object> headerMap = Maps.newLinkedHashMap(); //定义标题行
        headerMap.put("通道号", "channel_id");
        headerMap.put("设备编号", "uuid");
        headerMap.put("测点类型", "point_type");
        headerMap.put("测点名称", "point_name");
        headerMap.put("IEC测点", "iec_point");
        headerMap.put("设备测点", "point_id");
        headerMap.put("状态(1：启用，0：停用)", "is_enable");
        headerMap.put("修正系数", "iec_parameter");
        headerMap.put("偏移量", "iec_offset");
//        headerMap.put("设备类型", "device_type");
        headerMap.put("取反(1：是，0：否)", "iec_negate");
        params.put("header", headerMap);

        try {
            long start = System.currentTimeMillis() / 1000;
            Map<String, Object> map = EasyPoiUtils.importExcelForMap(new FileInputStream("E:\\google_downloads\\IEC测点映射20200317102927.xls"), params);
            List<Map<String, Object>> dataList = (List<Map<String, Object>>) map.get("dataList");
            long end = System.currentTimeMillis() / 1000;
            log.info("从Excel导入数据一共 {} 行 ，消耗时间{}秒", dataList.size(), end - start);
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Excel使用sax导入,使用实体类")
    @GetMapping("/importExcelBySaxForPojo")
    public String importExcelBySaxForPojo(HttpServletResponse response) {
        Map<String, Object> params = Maps.newLinkedHashMap();
        params.put("headRows", 1); //表头行数
        params.put("titleRows", 1); //表格标题行数
        try {
            long start = System.currentTimeMillis() / 1000;
            List<InsightIec104Mapping> result = EasyPoiUtils.importExcelBySaxForPojo(new FileInputStream(
                    new File("E:\\google_downloads\\IEC测点映射20200317102927.xlsx")), InsightIec104Mapping.class, params);
            long end = System.currentTimeMillis() / 1000;
            log.info("消耗时间{}秒", end - start);
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Excel使用sax导入,map导入")
    @GetMapping("/importExcelBySaxForMap")
    public String importExcelBySax(HttpServletResponse response) {
        Map<String, Object> params = Maps.newLinkedHashMap();
        params.put("headRows", 1); //表头行数
        params.put("titleRows", 1); //表格标题行数
        //导入时校验数据标题
        Map<String, Object> headerMap = Maps.newLinkedHashMap(); //定义标题行
        headerMap.put("通道号", "channel_id");
        headerMap.put("设备编号", "uuid");
        headerMap.put("测点类型", "point_type");
        headerMap.put("测点名称", "point_name");
        headerMap.put("IEC测点", "iec_point");
        headerMap.put("设备测点", "point_id");
//        headerMap.put("设备类型", "device_type");
        headerMap.put("状态(1：启用，0：停用)", "is_enable");
        headerMap.put("修正系数", "iec_parameter");
        headerMap.put("偏移量", "iec_offset");
        headerMap.put("取反(1：是，0：否)", "iec_negate");
        params.put("header", headerMap);
        try {
            long start = System.currentTimeMillis() / 1000;
            Map<String, Object> result = EasyPoiUtils.importExcelBySaxForMap(new FileInputStream(
                    new File("E:\\google_downloads\\IEC测点映射20200317102927.xlsx")), params);
            List<Map<String, Object>> dataList = (List<Map<String, Object>>) result.get("dataList");
            long end = System.currentTimeMillis() / 1000;
            log.info("消耗时间{}秒",end - start);
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }
}