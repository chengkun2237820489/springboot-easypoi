package com.chengkun.controller;
/**
 * sungrow all right reserved
 **/

import cn.afterturn.easypoi.excel.entity.result.ExcelImportResult;
import cn.afterturn.easypoi.util.PoiMergeCellUtil;
import com.alibaba.fastjson.JSONObject;
import com.chengkun.entity.BusinessData;
import com.chengkun.entity.Good;
import com.chengkun.entity.InsightIec104Mapping;
import com.chengkun.entity.SaxExcelImportResult;
import com.chengkun.service.ExcelService;
import com.chengkun.utils.EasyPoiUtils;
import com.chengkun.utils.handler.*;
import com.chengkun.utils.style.DefaultExportExcelStyle;
import com.google.common.collect.*;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileNotFoundException;
import java.util.*;

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
        EasyPoiUtils.exportExcel(list, "easypoi实体类导出", "Export", InsightIec104Mapping.class, "测试.xls","D:\\Download\\easypoi\\easypoi实体类导出测试.xls");
        EasyPoiUtils.exportExcel(list, "easypoi实体类导出", "Export", InsightIec104Mapping.class, "easypoi实体类导出测试.xls", response);
    }

    @ApiOperation(value = "Excel导出样式测试")
    @GetMapping("/exportExcelByStyle")
    public void exportExcelByStyle(HttpServletResponse response) {
        // 模拟从数据库获取需要导出的数据
        List<InsightIec104Mapping> personList = excelService.findAll();
        // 导出操作
        EasyPoiUtils.exportExcel(personList, "easypoi导出样式功能", "Export", InsightIec104Mapping.class, "excel样式测试.xls", "D:\\Download\\easypoi\\easypoi导出样式功能.xls", DefaultExportExcelStyle.class);
        EasyPoiUtils.exportExcel(personList, "easypoi导出样式功能", "Export", InsightIec104Mapping.class, "easypoi导出样式功能测试.xls", DefaultExportExcelStyle.class, response);
    }

    @ApiOperation(value = "Excel实体类导出(大数据)")
    @GetMapping("/exportBigExcelForPojo")
    public void exportBigExcelForPojo(HttpServletResponse response) {
        // 模拟从数据库获取需要导出的数据
        List<InsightIec104Mapping> personList = excelService.findAll();
        // 导出操作
        EasyPoiUtils.exportBigExcel(personList, "easypoi大数据实体类导出", "Export", InsightIec104Mapping.class, "easypoi大数据实体类导出.xls", "D:\\Download\\easypoi\\easypoi大数据实体类导出.xls",DefaultExportExcelStyle.class, 20);
        EasyPoiUtils.exportBigExcel(personList, "easypoi大数据实体类导出", "Export", InsightIec104Mapping.class, "easypoi大数据实体类导出.xls", DefaultExportExcelStyle.class, 20, response);
    }

    @ApiOperation(value = "csv实体类导出")
    @GetMapping("/exportCsv")
    public void exportCsv(HttpServletResponse response) {
        // 模拟从数据库获取需要导出的数据
        List<InsightIec104Mapping> personList = excelService.findAll();
        // 导出操作
        EasyPoiUtils.exportCsv(personList, InsightIec104Mapping.class, "导出csv.csv", "D:\\Download\\easypoi\\导出csv.csv");
        EasyPoiUtils.exportCsv(personList, InsightIec104Mapping.class, "导出csv.csv", response);
    }

    @ApiOperation(value = "使用map导出")
    @GetMapping("/exportExcelForMap")
    public void exportExcelForMap(HttpServletResponse response) {
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
        EasyPoiUtils.exportExcelForMap(list, "easypoi导出map功能", "Export", headerMap, "easypoi导出map功能.xlsx", "D:\\Download\\easypoi\\easypoi导出map功能.xlsx",DefaultExportExcelStyle.class);
        EasyPoiUtils.exportExcelForMap(list, "easypoi导出map功能", "Export", headerMap, "easypoi导出map功能.xlsx", DefaultExportExcelStyle.class, response);

    }

    @ApiOperation(value = "使用map导出(大数据)")
    @GetMapping("/exportBigExcelForMap")
    public void exportBigExcelForMap(HttpServletResponse response) {
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
        // 导出操作
        EasyPoiUtils.exportBigExcelForMap(list, "easypoi导出大数据map功能", "Export", headerMap, "excel样式测试.xls", "D:\\Download\\easypoi\\excel样式测试.xls",DefaultExportExcelStyle.class, 20);
        EasyPoiUtils.exportBigExcelForMap(list, "easypoi导出大数据map功能", "Export", headerMap, "excel样式测试.xls", DefaultExportExcelStyle.class, 20, response);
    }

    @ApiOperation(value = "使用map导出csv")
    @GetMapping("/exportCsvForMap")
    public void exportCsvForMap(HttpServletResponse response) {
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
        EasyPoiUtils.exportCsvForMap(list, headerMap, "map导出csv.csv","D:\\Download\\easypoi\\map导出csv.csv");
        EasyPoiUtils.exportCsvForMap(list, headerMap, "map导出csv.csv", response);

    }

    @ApiOperation(value = "Excel导出合并单元格")
    @GetMapping("/exportExcelForMergeCells")
    public void exportExcelForMergeCells(HttpServletResponse response) {
        System.out.println(1);
        // 模拟从数据库获取需要导出的数据
        List<InsightIec104Mapping> personList = excelService.findAll();
        // 导出操作 //List<?> list, String title, String sheetName, Class<?> pojoClass,  Class<?> style
        Workbook workbook = EasyPoiUtils.exportExcel(personList, "easypoi导出功能", "Export", InsightIec104Mapping.class, DefaultExportExcelStyle.class);
        Sheet sheet = workbook.getSheetAt(0);
        /**合并单元格**/
        PoiMergeCellUtil.mergeCells(sheet, 1, 8, 0, 1);
        if (workbook != null) {
            EasyPoiUtils.downLoadExcel("合并单元格测试.xls", null, response, workbook);
        }

    }

    @ApiOperation(value = "Excel导出复合表头使用map")
    @GetMapping("/exportExcelByCompositeForMap")
    public void exportExcelByCompositeForMap(HttpServletResponse response) {
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
//            valMap.put("uv", "uv" + i);
            List<Map<String, Object>> list_1 = new ArrayList<Map<String, Object>>();
            Map<String, Object> valMap_1 = new HashMap<String, Object>();
            for (int j = 0; j < 5; j++) {
                valMap_1.put("data" + j, "数据" + j);
            }
            list_1.add(valMap_1);
            valMap.put("businessData", list_1);
            list.add(valMap);
        }
        EasyPoiUtils.exportHeadersExcelForMap(list, "easypoi导出复合标题", "Export", table, "easypoi导出复合标题.xls", "D:\\Download\\easypoi\\easypoi导出复合标题.xls", DefaultExportExcelStyle.class);
        EasyPoiUtils.exportHeadersExcelForMap(list, "easypoi导出复合标题", "Export", table, "easypoi导出复合标题.xls", DefaultExportExcelStyle.class, response);
    }

    @ApiOperation(value = "Excel导出复合表头使用注解")
    @GetMapping("/exportExcelByCompositeForPojo")
    public void exportExcelByCompositeForPojo(HttpServletResponse response) {
        List<Good> goods = new ArrayList<>();
        //文件数据
        Good good;
        for (int i = 0; i < 10; i++) {
            good = new Good();
            good.setId(i);
            good.setDt("日期" + i);
            good.setPv("pv" + i);
            List<BusinessData> data = new ArrayList<>();
            BusinessData businessData = new BusinessData();
            businessData.setData1("数据1");
            businessData.setData2("数据2");
            businessData.setData3("数据3");
            businessData.setData4("数据4");
            data.add(businessData);
            good.setBusinessData(data);
            goods.add(good);
        }
        EasyPoiUtils.exportExcel(goods, "Excel导出复合表头使用注解", "Export", Good.class, "测试.xls", "D:\\Download\\easypoi\\测试.xls");
        EasyPoiUtils.exportExcel(goods, "Excel导出复合表头使用注解", "Export", Good.class, "测试.xls", response);
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
        EasyPoiUtils.exportSheetsExcelForMap(sheets, "多sheet页map导出.xlsx","D:\\Download\\easypoi\\多sheet页map导出.xlsx", DefaultExportExcelStyle.class);
        EasyPoiUtils.exportSheetsExcelForMap(sheets, "多sheet页map导出.xlsx", DefaultExportExcelStyle.class, response);
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
        EasyPoiUtils.exportSheetsExcel(sheets, "多sheet页pojo导出.xlsx","D:\\Download\\easypoi\\多sheet页pojo导出.xlsx", DefaultExportExcelStyle.class);
        EasyPoiUtils.exportSheetsExcel(sheets, "多sheet页pojo导出.xlsx", DefaultExportExcelStyle.class, response);
    }

    @ApiOperation(value = "根据模板导出Excel")
    @GetMapping("/exportExcelForTemplate")
    public void exportExcelForTemplate(HttpServletResponse response) throws FileNotFoundException {
        //封装模板数据
        Map<String, Object> dataMap = Maps.newLinkedHashMap();
        dataMap.put("title", "模板导出测试");
        dataMap.put("nowTime", new Date());
        dataMap.put("unitName", "悟耘信息");
        dataMap.put("order", new Date().getTime());
        List<Map<String, Object>> mapList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            Map<String, Object> testMap = new HashMap<>();
            testMap.put("name", "小明" + i);
            testMap.put("nums", i);
            testMap.put("type", "食品");
            testMap.put("remark", "甜食");
            mapList.add(testMap);
        }
        dataMap.put("list", mapList);
        //获取模板文件路径
        File file = ResourceUtils.getFile(ResourceUtils.CLASSPATH_URL_PREFIX + "excelTemplate/excelTemplate.xlsx");
        String templatePath = file.getPath();
        EasyPoiUtils.exportExcelForTemplate(dataMap, templatePath, "模板导出测试.xlsx", "D:\\Download\\easypoi\\模板导出测试.xlsx");
        EasyPoiUtils.exportExcelForTemplate(dataMap, templatePath, "模板导出测试.xlsx", response);
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

        try {
            Map<String, Object> result = EasyPoiUtils.importExcel("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx", InsightIec104Mapping.class, params);
//            Map<String, Object> result = EasyPoiUtils.importExcel(new FileInputStream("E:\\Chrom Downloads\\IEC测点映射20200317102927.xls"), InsightIec104Mapping.class, params);
            if ((int) result.get("flag") == 1) {
                List<InsightIec104Mapping> list = (List<InsightIec104Mapping>) result.get("dataList");
                for (InsightIec104Mapping iec : list) {
                    log.info("从Excel导入数据到数据库的详细为 ：{}", JSONObject.toJSONString(iec));
                    //TODO 将导入的数据做保存数据库操作
                }
                log.info("从Excel导入数据一共 {} 行 ", list.size());
            } else {
                log.error("导入失败：{}", result.get("msg"));
            }

        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Excel导入校验")
    @GetMapping("/importVerifyExcel")
    public String importVerifyExcel(HttpServletResponse response) {
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
        //校验处理类
        Map<String, Object> dbDate = Maps.newHashMap(); // 模拟数据数据
        List<Integer> pointTypeList = ImmutableList.of(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);
        dbDate.put("point_type", pointTypeList);
        InsightIec104MappingVerifyHandler verifyHandler = new InsightIec104MappingVerifyHandler();
        verifyHandler.setDbDate(dbDate); // 如果需要跟数据库数据对比，将这些数据传入
        try {
            Map<String, Object> result = EasyPoiUtils.importVerifyExcel("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx", InsightIec104Mapping.class, params, verifyHandler);
//            Map<String, Object> result = EasyPoiUtils.importExcel(new FileInputStream("E:\\Chrom Downloads\\IEC测点映射20200317102927.xls"), InsightIec104Mapping.class, params);
            if ((int) result.get("flag") == 1) {
                ExcelImportResult excelImportResult = (ExcelImportResult) result.get("dataList");

                log.info("从Excel导入数据一共 {} 行 ", excelImportResult.getList().size());
                log.info("从Excel导入失败一共 {} 行 ", excelImportResult.getFailList().size());
            } else {
                log.error("导入失败：{}", result.get("msg"));
            }

        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        } finally {
            // 清除threadLocal 防止内存泄漏
            ThreadLocal<List<InsightIec104Mapping>> threadLocal = verifyHandler.getThreadLocal();
            if (threadLocal != null) {
                threadLocal.remove();
            }
        }
        return "导入成功";
    }

    @ApiOperation(value = "Excel使用map导入")
    @GetMapping("/importExcelForMap")
    public String importExcelForMap(HttpServletResponse response) {
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
            Map<String, Object> map = EasyPoiUtils.importExcelForMap("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx", params);
//            Map<String, Object> map = EasyPoiUtils.importExcelForMap(new FileInputStream("E:\\Chrom Downloads\\IEC测点映射20200317102927.xls"), params);
            if ((int) map.get("flag") == 1) {
                List<Map<String, Object>> list = (List<Map<String, Object>>) map.get("dataList");
                for (Map<String, Object> iec : list) {
                    log.info("从Excel导入数据到数据库的详细为 ：{}", JSONObject.toJSONString(iec));
                    //TODO 将导入的数据做保存数据库操作
                }
                long end = System.currentTimeMillis() / 1000;
                log.info("从Excel导入数据一共 {} 行 ，消耗时间{}秒", list.size(), end - start);
            } else {
                log.error("导入失败：{}", map.get("msg"));
            }
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Excel使用map导入校验")
    @GetMapping("/importVerifyExcelForMap")
    public String importVerifyExcelForMap(HttpServletResponse response) {
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
        //校验处理类
        Map<String, Object> dbDate = Maps.newHashMap(); // 模拟数据数据
        List<Integer> pointTypeList = ImmutableList.of(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);
        dbDate.put("point_type", pointTypeList);
        DefaultMapVerifyHandler verifyHandler = new DefaultMapVerifyHandler();
        verifyHandler.setDbDate(dbDate); // 如果需要跟数据库数据对比，将这些数据传入

        try {
            long start = System.currentTimeMillis() / 1000;
            Map<String, Object> map = EasyPoiUtils.importVerifyExcelForMap("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx", params, verifyHandler);
//            Map<String, Object> map = EasyPoiUtils.importExcelForMap(new FileInputStream("E:\\Chrom Downloads\\IEC测点映射20200317102927.xls"), params);
            if ((int) map.get("flag") == 1) {
                ExcelImportResult excelImportResult = (ExcelImportResult) map.get("dataList");

                log.info("从Excel导入数据一共 {} 行 ", excelImportResult.getList().size());
                log.info("从Excel导入失败一共 {} 行 ", excelImportResult.getFailList().size());
            } else {
                log.error("导入失败：{}", map.get("msg"));
            }
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        } finally {
            // 清除threadLocal 防止内存泄漏
            ThreadLocal<List<Map<String, Object>>> threadLocal = verifyHandler.getThreadLocal();
            threadLocal = threadLocal;
            if (threadLocal != null) {
                threadLocal.remove();
            }
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
            Map<String, Object> map = EasyPoiUtils.importExcelBySaxForPojo("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx", InsightIec104Mapping.class, params);
//            Map<String, Object> map = EasyPoiUtils.importExcelBySaxForPojo(new FileInputStream("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx"), InsightIec104Mapping.class, params);
            if ((int) map.get("flag") == 1) {
                List<InsightIec104Mapping> list = (List<InsightIec104Mapping>) map.get("dataList");
                for (InsightIec104Mapping iec : list) {
                    log.info("从Excel导入数据到数据库的详细为 ：{}", JSONObject.toJSONString(iec));
                    //TODO 将导入的数据做保存数据库操作
                }
                long end = System.currentTimeMillis() / 1000;
                log.info("从Excel导入数据一共 {} 行 ，消耗时间{}秒", list.size(), end - start);
            } else {
                log.error("导入失败：{}", map.get("msg"));
            }
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Excel使用sax导入校验,使用实体类")
    @GetMapping("/importVerifyExcelBySaxForPojo")
    public String importVerifyExcelBySaxForPojo(HttpServletResponse response) {
        Map<String, Object> params = Maps.newLinkedHashMap();
        params.put("headRows", 1); //表头行数
        params.put("titleRows", 1); //表格标题行数
        try {
            long start = System.currentTimeMillis() / 1000;
            //校验处理类
            Map<String, Object> dbDate = Maps.newHashMap(); // 模拟数据数据
            DefaultSaxReadVerifyHandler<InsightIec104Mapping> verifyHandler = new DefaultSaxReadVerifyHandler<>(); // 数据校验类，包含校验处理类和结果集
            SaxExcelImportResult<InsightIec104Mapping> saxExcelImportResult = new SaxExcelImportResult<>(); // 返回数据
            InsightIec104MappingSaxVerifyHandler iecVerifyHandler = new InsightIec104MappingSaxVerifyHandler();
            List<Integer> pointTypeList = ImmutableList.of(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);
            dbDate.put("point_type", pointTypeList);
            iecVerifyHandler.setDbDate(dbDate);
            iecVerifyHandler.setResult(saxExcelImportResult);
            verifyHandler.setIReadHandler(iecVerifyHandler); // 添加校验规则
            verifyHandler.setResult(iecVerifyHandler.getResult()); // 添加校验结果
            Map<String, Object> map = EasyPoiUtils.importVerifyExcelBySaxForPojo("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx", InsightIec104Mapping.class, params, verifyHandler);
//            Map<String, Object> map = EasyPoiUtils.importVerifyExcelBySaxForPojo(new FileInputStream("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx"), InsightIec104Mapping.class, params, verifyHandler);
            if ((int) map.get("flag") == 1) {
                SaxExcelImportResult<InsightIec104Mapping> excelImportResult = (SaxExcelImportResult<InsightIec104Mapping>) map.get("dataList");

                log.info("从Excel导入数据一共 {} 行 ", excelImportResult.getList().size());
                log.info("从Excel导入失败一共 {} 行 ", excelImportResult.getFailList().size());
            } else {
                log.error("导入失败：{}", map.get("msg"));
            }
            long end = System.currentTimeMillis() / 1000;
            log.info("消耗时间{}秒", end - start);
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Excel使用sax导入,map导入")
    @GetMapping("/importExcelBySaxForMap")
    public String importExcelBySaxForMap(HttpServletResponse response) {
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
            Map<String, Object> result = EasyPoiUtils.importExcelBySaxForMap("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx", params);
//            Map<String, Object> result = EasyPoiUtils.importExcelBySaxForMap(new FileInputStream("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx"), params);
            List<Map<String, Object>> dataList = (List<Map<String, Object>>) result.get("dataList");
            long end = System.currentTimeMillis() / 1000;
            log.info("消耗时间{}秒", end - start);
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Excel使用sax导入校验,map导入")
    @GetMapping("/importVerifyExcelBySaxForMap")
    public String importVerifyExcelBySaxForMap(HttpServletResponse response) {
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
        //校验处理类
        Map<String, Object> dbDate = Maps.newHashMap(); // 模拟数据数据
        List<Integer> pointTypeList = ImmutableList.of(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);
        dbDate.put("point_type", pointTypeList);
        DefaultMapVerifyHandler defaultMapVerifyHandler = new DefaultMapVerifyHandler(); // 如果是xlx使用普通校验处理类
        try {
            long start = System.currentTimeMillis() / 1000;
            DefaultSaxReadVerifyHandler<Map<String, Object>> verifyHandler = new DefaultSaxReadVerifyHandler<Map<String, Object>>();
            DefaultMapSaxVerifyHandler mapIReadVerifyHandler = new DefaultMapSaxVerifyHandler(headerMap, dbDate);
            SaxExcelImportResult<Map<String, Object>> saxExcelImportResult = new SaxExcelImportResult<>(); // 返回数据
            mapIReadVerifyHandler.setResult(saxExcelImportResult);
            verifyHandler.setIReadHandler(mapIReadVerifyHandler); // 添加校验规则
            verifyHandler.setResult(mapIReadVerifyHandler.getResult()); // 添加校验结果

            defaultMapVerifyHandler.setDbDate(dbDate); // 如果需要跟数据库数据对比，将这些数据传入
            verifyHandler.setVerifyHandler(defaultMapVerifyHandler);
            Map<String, Object> map = EasyPoiUtils.importVerifyExcelBySaxForMap("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx", params, verifyHandler);
//            Map<String, Object> result = EasyPoiUtils.importExcelBySaxForMap(new FileInputStream("E:\\Chrom Downloads\\IEC测点映射20200317102927.xlsx"), params);
            if ((int) map.get("flag") == 1) {
                ExcelImportResult<Map<String, Object>> excelImportResult = (ExcelImportResult<Map<String, Object>>) map.get("dataList");
                log.info("从Excel导入数据一共 {} 行 ", excelImportResult.getList().size());
                log.info("从Excel导入失败一共 {} 行 ", excelImportResult.getFailList().size());
            } else {
                log.error("导入失败：{}", map.get("msg"));
            }
            long end = System.currentTimeMillis() / 1000;
            log.info("消耗时间{}秒", end - start);
        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        } finally {
            // 清除threadLocal 防止内存泄漏
            ThreadLocal<List<Map<String, Object>>> threadLocal = defaultMapVerifyHandler.getThreadLocal();
            threadLocal = threadLocal;
            if (threadLocal != null) {
                threadLocal.remove();
            }
        }
        return "导入成功";
    }

    @ApiOperation(value = "Csv导入")
    @GetMapping("/importCsv")
    public String importCsv(HttpServletResponse response) {
        Map<String, Object> params = Maps.newLinkedHashMap();
        params.put("headRows", 1); //表头行数
        params.put("titleRows", 0); //表格标题行数
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

        try {
            Map<String, Object> result = EasyPoiUtils.importCsv("E:\\Chrom Downloads\\实体类导出csv.csv", InsightIec104Mapping.class, params);
//            Map<String, Object> result = EasyPoiUtils.importExcel(new FileInputStream("E:\\Chrom Downloads\\实体类导出csv.csv"), InsightIec104Mapping.class, params);
            if ((int) result.get("flag") == 1) {
                List<InsightIec104Mapping> list = (List<InsightIec104Mapping>) result.get("dataList");
                for (InsightIec104Mapping iec : list) {
                    log.info("从Excel导入数据到数据库的详细为 ：{}", JSONObject.toJSONString(iec));
//                    TODO 将导入的数据做保存数据库操作
                }
                log.info("从Excel导入数据一共 {} 行 ", list.size());
            } else {
                log.error("导入失败：{}", result.get("msg"));
            }

        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Csv导入map")
    @GetMapping("/importCsvForMap")
    public String importCsvForMap(HttpServletResponse response) {
        Map<String, Object> params = Maps.newLinkedHashMap();
        params.put("headRows", 1); //表头行数
        params.put("titleRows", 0); //表格标题行数
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

        try {
            Map<String, Object> result = EasyPoiUtils.importCsv("E:\\Chrom Downloads\\map导出csv.csv", Map.class, params);
//            Map<String, Object> result = EasyPoiUtils.importExcel(new FileInputStream("E:\\Chrom Downloads\\实体类导出csv.csv"), InsightIec104Mapping.class, params);
            if ((int) result.get("flag") == 1) {
                List<Map<String, Object>> list = (List<Map<String, Object>>) result.get("dataList");
                for (Map<String, Object> iec : list) {
                    log.info("从Excel导入数据到数据库的详细为 ：{}", JSONObject.toJSONString(iec));
//                    TODO 将导入的数据做保存数据库操作
                }
                log.info("从Excel导入数据一共 {} 行 ", list.size());
            } else {
                log.error("导入失败：{}", result.get("msg"));
            }

        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Csv导入校验使用实体类")
    @GetMapping("/importVerifyCsvForPojo")
    public String importVerifyCsvForPojo(HttpServletResponse response) {
        Map<String, Object> params = Maps.newLinkedHashMap();
        params.put("headRows", 1); //表头行数
        params.put("titleRows", 0); //表格标题行数
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

        //校验处理类
        Map<String, Object> dbDate = Maps.newHashMap(); // 模拟数据数据
        DefaultSaxReadVerifyHandler<InsightIec104Mapping> verifyHandler = new DefaultSaxReadVerifyHandler<>(); // 数据校验类，包含校验处理类和结果集
        SaxExcelImportResult<InsightIec104Mapping> saxExcelImportResult = new SaxExcelImportResult<>(); // 返回数据
        InsightIec104MappingSaxVerifyHandler iecVerifyHandler = new InsightIec104MappingSaxVerifyHandler();
        List<Integer> pointTypeList = ImmutableList.of(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);
        dbDate.put("point_type", pointTypeList);
        iecVerifyHandler.setDbDate(dbDate);
        iecVerifyHandler.setResult(saxExcelImportResult);
        verifyHandler.setIReadHandler(iecVerifyHandler); // 添加校验规则
        verifyHandler.setResult(iecVerifyHandler.getResult()); // 添加校验结果

        try {
            long start = System.currentTimeMillis() / 1000;
            Map<String, Object> result = EasyPoiUtils.importVerifyCsv("E:\\Chrom Downloads\\实体类导出csv.csv", InsightIec104Mapping.class, params, verifyHandler);
//            Map<String, Object> result = EasyPoiUtils.importExcel(new FileInputStream("E:\\Chrom Downloads\\实体类导出csv.csv"), InsightIec104Mapping.class, params);
            if ((int) result.get("flag") == 1) {
                SaxExcelImportResult<InsightIec104Mapping> excelImportResult = (SaxExcelImportResult<InsightIec104Mapping>) result.get("dataList");

                log.info("从Excel导入数据一共 {} 行 ", excelImportResult.getList().size());
                log.info("从Excel导入失败一共 {} 行 ", excelImportResult.getFailList().size());
            } else {
                log.error("导入失败：{}", result.get("msg"));
            }
            long end = System.currentTimeMillis() / 1000;
            log.info("消耗时间{}秒", end - start);

        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "Csv导入校验使用map")
    @GetMapping("/importVerifyCsvForMap")
    public String importVerifyCsvForMap(HttpServletResponse response) {
        Map<String, Object> params = Maps.newLinkedHashMap();
        params.put("headRows", 1); //表头行数
        params.put("titleRows", 0); //表格标题行数
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

        //校验处理类
        Map<String, Object> dbDate = Maps.newHashMap(); // 模拟数据数据
        List<Integer> pointTypeList = ImmutableList.of(1, 2, 3, 4, 5, 6, 7, 8, 9, 10);
        dbDate.put("point_type", pointTypeList);
        DefaultSaxReadVerifyHandler<Map<String, Object>> verifyHandler = new DefaultSaxReadVerifyHandler<Map<String, Object>>();
        DefaultMapSaxVerifyHandler mapIReadVerifyHandler = new DefaultMapSaxVerifyHandler(headerMap, dbDate);
        SaxExcelImportResult<Map<String, Object>> saxExcelImportResult = new SaxExcelImportResult<>(); // 返回数据
        mapIReadVerifyHandler.setResult(saxExcelImportResult);
        verifyHandler.setIReadHandler(mapIReadVerifyHandler); // 添加校验规则
        verifyHandler.setResult(mapIReadVerifyHandler.getResult()); // 添加校验结果

        try {
            long start = System.currentTimeMillis() / 1000;
            Map<String, Object> result = EasyPoiUtils.importVerifyCsv("E:\\Chrom Downloads\\实体类导出csv.csv", Map.class, params, verifyHandler);
//            Map<String, Object> result = EasyPoiUtils.importExcel(new FileInputStream("E:\\Chrom Downloads\\实体类导出csv.csv"), InsightIec104Mapping.class, params);
            if ((int) result.get("flag") == 1) {
                ExcelImportResult<Map<String, Object>> excelImportResult = (ExcelImportResult<Map<String, Object>>) result.get("dataList");

                log.info("从Excel导入数据一共 {} 行 ", excelImportResult.getList().size());
                log.info("从Excel导入失败一共 {} 行 ", excelImportResult.getFailList().size());
            } else {
                log.error("导入失败：{}", result.get("msg"));
            }
            long end = System.currentTimeMillis() / 1000;
            log.info("消耗时间{}秒", end - start);

        } catch (Exception e) {
            log.error("导入失败：{}", e.getMessage());
        }
        return "导入成功";
    }

    @ApiOperation(value = "使用map导出iec映射")
    @GetMapping("/exportExcelIecMapping")
    public void exportExcelIecMapping(HttpServletResponse response) {
        excelService.exportExcelIecMapping(response);
    }
}