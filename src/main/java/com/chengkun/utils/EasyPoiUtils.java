package com.chengkun.utils;
/**
 * sungrow all right reserved
 **/

import cn.afterturn.easypoi.entity.vo.NormalExcelConstants;
import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.ExcelImportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import cn.afterturn.easypoi.excel.entity.enmus.ExcelType;
import cn.afterturn.easypoi.excel.entity.params.ExcelExportEntity;
import cn.afterturn.easypoi.excel.export.ExcelExportService;
import cn.afterturn.easypoi.handler.inter.IReadHandler;
import com.chengkun.utils.handler.MapIReadHandler;
import com.chengkun.utils.handler.MapImportHandler;
import com.chengkun.utils.sax.SaxReadExcel;
import com.chengkun.utils.style.ExportExcelStyle;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Table;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.ResponseEntity;
import org.springframework.web.client.RestTemplate;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.*;

/**
 * @Description easypoi导入导出工具类
 * @Author chengkun
 * @Date 2020/3/20 9:24
 **/
@Log4j2
@SuppressWarnings("all")
public class EasyPoiUtils {
    /**
     * 功能描述：复杂导出Excel，包括文件名以及表名。创建表头
     *
     * @param list           导出的实体类
     * @param title          表头名称
     * @param sheetName      sheet表名
     * @param pojoClass      映射的实体类
     * @param fileName       文件名称
     * @param filePath       文件路径 + 文件名称
     * @param style          文件样式
     * @param isCreateHeader 是否创建表头
     * @param response
     * @return
     * @Author chengkun
     * @Date 2020/3/17 19:40
     */
    public static void exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName, String filePath, Class<?> style, boolean isCreateHeader, HttpServletResponse response) {
        ExportParams exportParams = new ExportParams(title, sheetName);
        exportParams.setCreateHeadRows(isCreateHeader);
        exportParams.setStyle(style);
        defaultExport(list, pojoClass, fileName, filePath, response, exportParams);
    }

    /**
     * @Description 自定义样式导出
     * @Author chengkun
     * @Date 2020/3/18 9:05
     * @Param list          导出的实体类
     * @Param title         表头名称
     * @Param sheetName     sheet表名
     * @Param pojoClass     映射的实体类
     * @Param fileName      文件名称
     * @Param filePath      文件路径 + 文件名称
     * @Param style         导出样式
     * @Param response
     * @Return void
     **/
    public static void exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName, String filePath, Class<?> style, HttpServletResponse response) {
        ExportParams exportParams = new ExportParams(title, sheetName);
        exportParams.setStyle(style);
        defaultExport(list, pojoClass, fileName, filePath, response, exportParams);
    }

    /**
     * @Description 自定义样式导出
     * @Author chengkun
     * @Date 2020/3/18 9:05
     * @Param list          导出的实体类
     * @Param title         excel大标题
     * @Param sheetName     sheet名
     * @Param header        表头名称  key => 标题行名称  value => 标题行英文标识，返回数据以此标识为key
     * @Param fileName      文件名称
     * @Param filePath      文件路径 + 文件名称
     * @Param style         导出样式
     * @Param response
     * @Return void
     **/
    public static void exportExcelForMap(List<?> list, String title, String sheetName, Map<String, Object> header, String fileName, String filePath, Class<?> style, HttpServletResponse response) {
        ExportParams exportParams = new ExportParams(title, sheetName);
        exportParams.setStyle(style);
        //构造对象等同于@Excel
        List<ExcelExportEntity> colList = new ArrayList<>();
        ExcelExportEntity exportEntity;
        for (String key : header.keySet()) {
            exportEntity = new ExcelExportEntity(key, header.get(key), getColWidth(key));
            colList.add(exportEntity);
        }
        defaultExport(list, colList, fileName, filePath, response, exportParams);
    }

    /**
     * @param isCreateHeader 是否创建表头
     * @Description map导出(是否创建表头)
     * @Author chengkun
     * @Date 2020/3/18 9:05
     * @Param list           导出的实体类
     * @Param title          excel大标题
     * @Param sheetName      sheet名
     * @Param header         表头名称  key => 标题行名称  value => 标题行英文标识，返回数据以此标识为key
     * @Param fileName       文件名称
     * @Param filePath       文件路径 + 文件名称
     * @Param style          导出样式
     * @Param response
     * @Return void
     **/
    public static void exportExcelForMap(List<?> list, String title, String sheetName, Map<String, Object> header, String fileName, String filePath, Class<?> style, boolean isCreateHeader, HttpServletResponse response) {
        ExportParams exportParams = new ExportParams(title, sheetName);
        exportParams.setCreateHeadRows(isCreateHeader); //是否创建表头
        exportParams.setStyle(style);
        //构造对象等同于@Excel
        List<ExcelExportEntity> colList = new ArrayList<>();
        ExcelExportEntity exportEntity;
        for (String key : header.keySet()) {
            exportEntity = new ExcelExportEntity(key, header.get(key), getColWidth(key));
            colList.add(exportEntity);
        }
        defaultExport(list, colList, fileName, filePath, response, exportParams);
    }

    /**
     * @Description 模板导出
     * @Author chengkun
     * @Date 2020/3/18 9:05
     * @Param params         模板封装数据
     * @Param templatePath   模板文件路径
     * @Param fileName      文件名称
     * @Param filePath      文件路径 + 文件名称
     * @Param response
     * @Return void
     **/
    public static void exportExcelForTemplate(Map<String, Object> params, String templatePath, String fileName, String filePath, HttpServletResponse response) {
        TemplateExportParams exportParams = new TemplateExportParams(templatePath); //模板参数，只需模板文件路径
        long start = System.currentTimeMillis() / 1000;
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, params);
        long end = System.currentTimeMillis() / 1000;
        log.info("导出excel处理时间：{}秒", end - start);
        downLoadExcel(fileName, filePath, response, workbook);
    }

    /**
     * @Description 获取单元格长度，为字体长度1.5倍,最大为255字符
     * @Author chengkun
     * @Date 2020/3/20 11:06
     * @Param key 单元格内容
     * @Return int
     **/
    private static int getColWidth(String key) {
        int colWidth = (int) (key.getBytes().length * 1.5); //单元格宽度为字体宽度1.5倍
        if (colWidth < 20) {
            colWidth = 20; //最小宽度为20
        } else if (colWidth > 255) {
            colWidth = 255; ////最大宽度为255
        }
        return colWidth;
    }

    /**
     * @Description 复合表头导出
     * @Author chengkun
     * @Date 2020/3/18 9:05
     * @Param list          导出的实体类
     * @Param title         大标题
     * @Param sheetName     sheet表名
     * @Param heaters       表头table
     * R => 标题行名称     C = >  标题行英文标识     V => Map<String,Object> 该标题行的子标题行与map导出标题行定义一值
     * @Param fileName      文件名称
     * @Param filePath      文件路径
     * @Param style         导出样式
     * @Param response
     * @Return void
     **/
    public static void exportHeadersExcelForMap(List<?> list, String title, String sheetName, Table<String, String, Object> heaters, String fileName, String filePath, Class<?> style, HttpServletResponse response) {
        ExportParams exportParams = new ExportParams(title, sheetName);
        exportParams.setStyle(style);
        //构造对象等同于@Excel
        List<ExcelExportEntity> colList = new ArrayList<>();
        List<ExcelExportEntity> groupList;
        ExcelExportEntity colEntity; //列对象
        ExcelExportEntity groupEntity; //列里面的组对象
        Map<String, Map<String, Object>> rowMap = heaters.rowMap();
        for (String headerName : rowMap.keySet()) {
            Map<String, Object> headerMap = rowMap.get(headerName);
            for (String header : headerMap.keySet()) {
                colEntity = new ExcelExportEntity(headerName, header, getColWidth(headerName));
                Map<String, Object> groupMap = (Map<String, Object>) headerMap.get(header);
                if (groupMap != null && groupMap.size() > 0) {
                    groupList = new ArrayList<>();
                    for (String key : groupMap.keySet()) {
                        groupEntity = new ExcelExportEntity(key, groupMap.get(key), getColWidth(key));
                        groupList.add(groupEntity);
                    }
                    colEntity.setList(groupList);
                }
                colList.add(colEntity);
            }
        }
        defaultExport(list, colList, fileName, filePath, response, exportParams);
    }

    /**
     * @Description 导出多sheet页，map导入
     * @Author chengkun
     * @Date 2020/3/18 9:05
     * @Param sheets        sheet页内容集合
     * sheets中map为每个sheet页内容 ：
     * header = > 标题行，map集合，与导出map定义一致    title = > 大标题   sheetName = > 工作簿名称     dataList = > sheet也数据
     * @Param fileName      文件名称
     * @Param filePath      文件路径 + 文件名称
     * @Param style         导出样式
     * @Param response
     * @Return void
     **/
    public static void exportSheetsExcelForMap(List<Map<String, Object>> sheets, String fileName, String filePath, Class<?> style, HttpServletResponse response) {
        List<Integer> dataSize = Lists.newArrayList();//导出数据大小，根据大小创建对应excel
        String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
        ExcelType type = ExcelType.HSSF; //工作簿类型
        if (suffix.equals("xls")) {
            type = ExcelType.HSSF;
        } else if (suffix.equals("xlsx")) {
            type = ExcelType.XSSF;
        }
        List<Map<String, Object>> sheetExportList = Lists.newArrayList(); //导入sheet集合
        //构造对象等同于@Excel
        ExportParams exportParams; //导入参数
        Map<String, Object> sheetExportMap; //导入sheet
        List<ExcelExportEntity> colList;    //sheet页标题行
        ExcelExportEntity exportEntity;    //标题行实体
        for (Map<String, Object> sheet : sheets) {
            //sheet页表头和sheet名称设置
            String title = String.valueOf(sheet.get("title"));
            String sheetName = String.valueOf(sheet.get("sheetName"));
            exportParams = new ExportParams(title, sheetName);
            exportParams.setStyle(style);
            exportParams.setType(type);
            //sheet页标题行设置
            Map<String, Object> headerMap = (Map<String, Object>) sheet.get("header");
            colList = new ArrayList<>();
            for (String key : headerMap.keySet()) {
                exportEntity = new ExcelExportEntity(key, headerMap.get(key), getColWidth(key));
                colList.add(exportEntity);
            }
            //sheet数据封装
            List<Map<String, Object>> dataList = (List<Map<String, Object>>) sheet.get("dataList");
            sheetExportMap = Maps.newLinkedHashMap();
            sheetExportMap.put(NormalExcelConstants.CLASS, ExcelExportEntity.class);
            sheetExportMap.put(NormalExcelConstants.DATA_LIST, dataList);
            sheetExportMap.put(NormalExcelConstants.PARAMS, exportParams);
            //这边为了方便，sheet1和sheet2用同一个表头(实际使用中可自行调整)
            sheetExportMap.put(NormalExcelConstants.MAP_LIST, colList);
            //放入sheet集合中
            sheetExportList.add(sheetExportMap);
            dataSize.add(dataList.size());
        }
        long start = System.currentTimeMillis() / 1000;
        Workbook workbook = getWorkbook(type, Collections.max(dataSize));
        for (Map<String, Object> map : sheetExportList) {
            ExcelExportService server = new ExcelExportService();
            ExportParams param = (ExportParams) map.get("params");
            @SuppressWarnings("unchecked")
            List<ExcelExportEntity> entity = (List<ExcelExportEntity>) map.get("mapList");
            Collection<?> data = (Collection<?>) map.get("data");
            server.createSheetForMap(workbook, param, entity, data);
        }
        long end = System.currentTimeMillis() / 1000;
        log.info("导出excel处理时间：{}秒", end - start);
        downLoadExcel(fileName, filePath, response, workbook);
    }

    /**
     * @Description 导出多sheet页，实体类导入
     * @Author chengkun
     * @Date 2020/3/18 9:05
     * @Param sheets    sheet页内容集合 List<Map<String, Object>>
     * sheets中map为每个sheet页内容 ：
     * header = > 标题行，为实体类class    title = > 大标题   sheetName = > 工作簿名称     dataList = > sheet也数据
     * @Param fileName      文件名称
     * @Param filePath      文件路径 + 文件名称
     * @Param style         导出样式
     * @Param response
     * @Return void
     **/
    public static void exportSheetsExcelForPoJo(List<Map<String, Object>> sheets, String fileName, String filePath, Class<?> style, HttpServletResponse response) {
        List<Integer> dataSize = Lists.newArrayList();//导出数据大小，根据大小创建对应excel
        String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
        ExcelType type = ExcelType.HSSF; //工作簿类型
        if (suffix.equals("xls")) {
            type = ExcelType.HSSF;
        } else if (suffix.equals("xlsx")) {
            type = ExcelType.XSSF;
        }
        List<Map<String, Object>> sheetExportList = Lists.newArrayList(); //导入sheet集合
        //构造对象等同于@Excel
        ExportParams exportParams; //导入参数
        Map<String, Object> sheetExportMap; //导入sheet
        for (Map<String, Object> sheet : sheets) {
            //sheet页表头和sheet名称设置
            String title = String.valueOf(sheet.get("title"));
            String sheetName = String.valueOf(sheet.get("sheetName"));
            exportParams = new ExportParams(title, sheetName);
            exportParams.setStyle(style);
            exportParams.setType(type);
            //sheet页标题行设置
            Class<?> header = (Class<?>) sheet.get("header");
            //sheet数据封装
            List<Map<String, Object>> dataList = (List<Map<String, Object>>) sheet.get("dataList");
            sheetExportMap = Maps.newLinkedHashMap();
            sheetExportMap.put("title", exportParams);
            sheetExportMap.put("entity", header);
            sheetExportMap.put("data", dataList);
            //放入sheet集合中
            sheetExportList.add(sheetExportMap);
            dataSize.add(dataList.size());
        }
        // 执行方法
        long start = System.currentTimeMillis() / 1000;
        Workbook workbook = ExcelExportUtil.exportExcel(sheetExportList, type);
        long end = System.currentTimeMillis() / 1000;
        log.info("导出excel处理时间：{}秒", end - start);
        downLoadExcel(fileName, filePath, response, workbook);
    }

    /**
     * @Description 根据类型创建workbook
     * @Author chengkun
     * @Date 2020/3/20 14:25
     * @Param type
     * @Param size
     * @Return org.apache.poi.ss.usermodel.Workbook
     **/
    private static Workbook getWorkbook(ExcelType type, int size) {
        if (ExcelType.HSSF.equals(type)) {
            return new HSSFWorkbook();
        } else if (size < 10000) {
            return new XSSFWorkbook();
        } else {
            return new SXSSFWorkbook();
        }
    }

    /**
     * 功能描述 实体类导出，使用默认样式
     *
     * @param list      导出的实体类
     * @param title     表头名称
     * @param sheetName sheet表名
     * @param pojoClass 映射的实体类
     * @param fileName  文件名称
     * @param filePath  文件路径 + 文件名称
     * @param response
     * @return
     * @Author chengkun
     * @Date 2020/3/17 19:40
     */
    public static void exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, String fileName, String filePath, HttpServletResponse response) {
        ExportParams exportParams = new ExportParams(title, sheetName);
        exportParams.setStyle(ExportExcelStyle.class);
        defaultExport(list, pojoClass, fileName, filePath, response, exportParams);
    }

    /**
     * 功能描述：Map 集合导出
     *
     * @return
     * @Param list          导出的实体类
     * @Param title         表头名称
     * @Param sheetName     sheet表名
     * @Param pojoClass     映射的实体类
     * @Param style         导出样式
     * @author chengkun
     * @date 2020/3/18 10:40
     */
    public static Workbook exportExcel(List<?> list, String title, String sheetName, Class<?> pojoClass, Class<?> style) {
        ExportParams exportParams = new ExportParams(title, sheetName);
        exportParams.setStyle(style);
        long start = System.currentTimeMillis() / 1000;
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, pojoClass, list);
        long end = System.currentTimeMillis() / 1000;
        log.info("导出excel处理时间：{}秒", end - start);
        return workbook;
    }

    /**
     * 功能描述：默认导出方法
     *
     * @param list         导出的实体集合
     * @param pojoClass    pojo实体
     * @param fileName     导出的文件名 数据量过大建议使用.xlsx
     * @param filePath     导出文件路径 + 文件名称
     * @param response
     * @param exportParams ExportParams封装实体
     * @return
     * @Author chengkun
     * @Date 2020/3/17 19:40
     */
    private static void defaultExport(List<?> list, Class<?> pojoClass, String fileName, String filePath, HttpServletResponse response, ExportParams exportParams) {
        String type = fileName.substring(fileName.lastIndexOf(".") + 1);
        if (type.equals("xls")) {
            exportParams.setType(ExcelType.HSSF);
        } else if (type.equals("xlsx")) {
            exportParams.setType(ExcelType.XSSF);
        }
        long start = System.currentTimeMillis() / 1000;
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, pojoClass, list);
        long end = System.currentTimeMillis() / 1000;
        log.info("导出excel处理时间：{}秒", end - start);
        downLoadExcel(fileName, filePath, response, workbook);
    }

    /**
     * @param list         导出的实体集合
     * @param colList      定义map实体
     * @param fileName     导出的文件名
     * @param filePath     导出文件路径 + 文件名
     * @param response
     * @param exportParams ExportParams封装实体
     * @Description map导出
     * @Author chengkun
     * @Date 2020/3/20 10:22
     * @Return void
     **/
    private static void defaultExport(List<?> list, List<ExcelExportEntity> colList, String fileName, String filePath, HttpServletResponse response, ExportParams exportParams) {
        String type = fileName.substring(fileName.lastIndexOf(".") + 1);
        if (type.equals("xls")) {
            exportParams.setType(ExcelType.HSSF);
        } else if (type.equals("xlsx")) {
            exportParams.setType(ExcelType.XSSF);
        }
        long start = System.currentTimeMillis() / 1000;
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, colList, list);
        long end = System.currentTimeMillis() / 1000;
        log.info("导出excel处理时间：{}秒", end - start);
        downLoadExcel(fileName, filePath, response, workbook);
    }

    /**
     * 功能描述：Excel导出
     *
     * @param fileName 文件名称
     * @param filePath 下载文件路径 路径+文件名称
     * @param response
     * @param workbook Excel对象
     * @return
     * @Author chengkun
     * @Date 2020/3/17 19:40
     */
    public static void downLoadExcel(String fileName, String filePath, HttpServletResponse response, Workbook workbook) {
        if (workbook != null) {
            if (response != null) {
                try {
                    response.setCharacterEncoding("UTF-8");
                    response.setHeader("content-Type", "application/vnd.ms-excel");
                    response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
                    workbook.write(response.getOutputStream());
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            } else if (filePath != null) {
                OutputStream out = null;
                try {
                    out = new FileOutputStream(filePath);
                    workbook.write(out);
                    out.flush();
                    out.close();
                } catch (IOException e) {
                    log.error("文件路径错误：" + filePath, e);
                } finally {
                    if (out != null) {
                        try {
                            out.close();
                        } catch (IOException e) {
                            log.error("文件路径错误：" + filePath, e);
                        }
                    }
                }
            }
        }

    }

    /**
     * 功能描述：默认导出方法
     *
     * @param list     导出的实体集合
     * @param fileName 导出的文件名
     * @param filePath 导出的文件路径 + 文件名
     * @param response
     * @return
     * @Author chengkun
     * @Date 2020/3/17 19:40
     */
    private static void defaultExport(List<Map<String, Object>> list, String fileName, String filePath, HttpServletResponse response) {
        String type = fileName.substring(fileName.lastIndexOf(".") + 1);
        Workbook workbook = null;
        long start = System.currentTimeMillis() / 1000;
        if (type.equals("xls")) {
            workbook = ExcelExportUtil.exportExcel(list, ExcelType.HSSF);
        } else if (type.equals("xlsx")) {
            workbook = ExcelExportUtil.exportExcel(list, ExcelType.XSSF);
        }
        long end = System.currentTimeMillis() / 1000;
        log.info("导出excel处理时间：{}秒", end - start);
        if (workbook != null) ;
        downLoadExcel(fileName, filePath, response, workbook);
    }

    /**
     * 功能描述：根据文件路径来导入Excel
     *
     * @param filePath  文件路径 + 文件名
     * @param pojoClass Excel实体类
     * @param params    导入参数
     * @return
     * @Author chengkun
     * @Date 2020/3/17 19:40
     */
    public static <T> Map<String, Object> importExcel(String filePath, Class<T> pojoClass, Map<String, Object> params) {
        Map<String, Object> responseMap = Maps.newLinkedHashMap();
        String type = filePath.substring(filePath.lastIndexOf(".") + 1);
        //后缀名不对直接返回
        if (!type.equals("xls") && !type.equals("xlsx")) {
            responseMap.put("flag", 0);
            responseMap.put("msg", "不是合法的Excel模板");
            return responseMap;
        }
        InputStream in = null;
        List<T> result = null;
        try {
            RestTemplate restTemplate = new RestTemplate();
            HttpHeaders headers = new HttpHeaders();
            ResponseEntity<byte[]> entity = restTemplate.exchange(filePath, HttpMethod.GET, new HttpEntity<>(headers), byte[].class);
            byte[] body = entity.getBody();
            in = new ByteArrayInputStream(body);
            responseMap = importExcel(in, pojoClass, params);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        } finally {
            try {
                in.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }
        return responseMap;
    }

    /**
     * 功能描述：根据文件流来导入Excel
     *
     * @param inputStream 文件流
     * @param pojoClass   Excel实体类
     * @param params      导入参数  headRows = > 表头行数 titleRows = > 表格标题行数 header = > 标题行，与map导出定义一致
     * @return
     * @Author chengkun
     * @Date 2020/3/17 19:40
     */
    public static <T> Map<String, Object> importExcel(InputStream inputStream, Class<T> pojoClass, Map<String, Object> params) {
        Map<String, Object> responseMap = Maps.newLinkedHashMap();
        ImportParams importParams = new ImportParams();
        // 数据处理
        importParams.setHeadRows((Integer) params.get("headRows"));
        importParams.setTitleRows((Integer) params.get("titleRows"));
        // 需要验证
        importParams.setNeedVerify(true);
        Map<String, Object> headerMap = (Map<String, Object>) params.get("header");
        Set<String> set = headerMap.keySet();
        String[] importFields = set.toArray(new String[set.size()]);
        importParams.setImportFields(importFields);
        List<T> list = null;
        try {
            long start = System.currentTimeMillis() / 1000;
            list = ExcelImportUtil.importExcel(inputStream, pojoClass, importParams);
            responseMap.put("dataList", list);
            long end = System.currentTimeMillis() / 1000;
            log.info("导入excel处理时间：{}秒", end - start);
            //判断标题行是否一致
        } catch (Exception e) {
            //自定义异常判断
            if (e.getMessage().indexOf("不是合法的Excel模板") != -1) {
                responseMap.put("flag", 0); //easypoi模板校验只能校验字段不一致的，数量少的
                responseMap.put("msg", "不是合法的Excel模板");
            } else {
                log.error("导入失败", e);
                throw new RuntimeException("导入失败", e);
            }
        }
        return responseMap;
    }

    /**
     * @Description map导入, 根据文件路径导入
     * @Author chengkun
     * @Date 2020/3/20 16:23
     * @Param filePath         文件路径 + 文件名
     * @Param params          导入参数
     * @Return java.util.Map<java.lang.String, java.lang.Object>
     **/
    public static Map<String, Object> importExcelForMap(String filePath, Map<String, Object> params) {
        Map<String, Object> responseMap = Maps.newLinkedHashMap();
        String type = filePath.substring(filePath.lastIndexOf(".") + 1);
        //后缀名不对直接返回
        if (!type.equals("xls") && !type.equals("xlsx")) {
            responseMap.put("flag", 0);
            responseMap.put("msg", "不是合法的Excel模板");
            return responseMap;
        }
        InputStream in = null;
        try {
            RestTemplate restTemplate = new RestTemplate();
            HttpHeaders headers = new HttpHeaders();
            ResponseEntity<byte[]> entity = restTemplate.exchange(filePath, HttpMethod.GET, new HttpEntity<>(headers), byte[].class);
            byte[] body = entity.getBody();
            in = new ByteArrayInputStream(body);
            responseMap = importExcelForMap(in, params);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        } finally {
            try {
                in.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }
        return responseMap;
    }


    /**
     * @Description map导入
     * @Author chengkun
     * @Date 2020/3/20 16:23
     * @Param fileInputStream 文件流
     * @Param params          导入参数 headRows = > 表头行数 titleRows = > 表格标题行数 header = > 标题行，与map导出定义一致
     * @Return java.util.Map<java.lang.String, java.lang.Object>
     **/
    public static Map<String, Object> importExcelForMap(InputStream inputStream, Map<String, Object> params) {
        Map<String, Object> responseMap = Maps.newLinkedHashMap();
        ImportParams importParams = new ImportParams();
        // 数据处理
        importParams.setHeadRows((Integer) params.get("headRows"));
        importParams.setTitleRows((Integer) params.get("titleRows"));
        // 需要验证
        importParams.setNeedVerify(true);
        Map<String, Object> headerMap = (Map<String, Object>) params.get("header");
        Set<String> set = headerMap.keySet();
        String[] importFields = set.toArray(new String[set.size()]);
        importParams.setImportFields(importFields);
        importParams.setDataHandler(new MapImportHandler(headerMap));
        List<T> list = null;
        try {
            long start = System.currentTimeMillis() / 1000;
            list = ExcelImportUtil.importExcel(inputStream, Map.class, importParams);
            responseMap.put("dataList", list);
            long end = System.currentTimeMillis() / 1000;
            log.info("导入excel处理时间：{}秒", end - start);
            //判断标题行是否一致
        } catch (Exception e) {
            //自定义异常判断
            if (e.getMessage().indexOf("不是合法的Excel模板") != -1) {
                responseMap.put("flag", 0); //easypoi模板校验只能校验字段不一致的，数量少的
                responseMap.put("msg", "不是合法的Excel模板");
            } else {
                log.error("导入失败", e);
                throw new RuntimeException("导入失败", e);
            }
        }
        return responseMap;
    }

    /**
     * @Description sax导入，实体类导入
     * @Author chengkun
     * @Date 2020/3/20 18:42
     * @Param inputStream
     * @Param pojoClass 实体类class
     * @Param params titleRows = > 表格标题行数 header = > 标题行，与map导出定义一致
     * @Return java.util.List<T>
     **/
    public static <T> List<T> importExcelBySaxForPojo(InputStream inputStream, Class<T> pojoClass, Map<String, Object> params) {
        ImportParams importParams = new ImportParams();
        // 数据处理
        importParams.setHeadRows((Integer) params.get("headRows"));
        importParams.setTitleRows((Integer) params.get("titleRows"));
        List<T> result = Lists.newArrayList();
        long start = System.currentTimeMillis() / 1000;
        new SaxReadExcel().readExcel(inputStream, pojoClass, importParams, new IReadHandler<T>() {
            /**
             * 处理解析对象
             *
             * @param insightIec104Mapping
             */
            @Override
            public void handler(T pojo) {
                result.add(pojo);
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
        log.info("导入excel处理时间：{}秒", end - start);
        return result;
    }

    /**
     * @Description sax导入，map导入
     * sax导入只支持xlsx，如果导入文件为xls使用普通导入
     * @Author chengkun
     * @Date 2020/3/20 21:15
     * @Param filePath
     * @Param params
     * @Return java.util.Map<java.lang.String, java.lang.Object>
     **/
    public static Map<String, Object> importExcelBySaxForMap(String filePath, Map<String, Object> params) {
        Map<String, Object> responseMap = Maps.newLinkedHashMap();
        String type = filePath.substring(filePath.lastIndexOf(".") + 1);
        //后缀名不对直接返回
        if (!type.equals("xls") && !type.equals("xlsx")) {
            responseMap.put("flag", 0);
            responseMap.put("msg", "不是合法的Excel模板");
            return responseMap;
        }
        InputStream in = null;
        try {
            RestTemplate restTemplate = new RestTemplate();
            HttpHeaders headers = new HttpHeaders();
            ResponseEntity<byte[]> entity = restTemplate.exchange(filePath, HttpMethod.GET, new HttpEntity<>(headers), byte[].class);
            byte[] body = entity.getBody();
            in = new ByteArrayInputStream(body);
            //如果是xls调用普通导入方法
            if (type.equals("xls")) {
                responseMap = importExcelForMap(in, params);
            } else {
                responseMap = importExcelBySaxForMap(in, params);
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        } finally {
            try {
                in.close();
            } catch (IOException e) {
                log.error(e.getMessage(), e);
            }
        }
        return responseMap;
    }

    /**
     * @Description sax导入，map导入
     * @Author chengkun
     * @Date 2020/3/20 21:15
     * @Param inputStream
     * @Param params headRows = > 表头行数 titleRows = > 表格标题行数 header = > 标题行，与map导出定义一致
     * @Return java.util.Map<java.lang.String, java.lang.Object>
     **/
    public static Map<String, Object> importExcelBySaxForMap(InputStream inputStream, Map<String, Object> params) {
        Map<String, Object> responseMap = Maps.newLinkedHashMap();
        ImportParams importParams = new ImportParams();
        // 数据处理
        importParams.setHeadRows((Integer) params.get("headRows"));
        importParams.setTitleRows((Integer) params.get("titleRows"));
        Map<String, Object> headerMap = (Map<String, Object>) params.get("header");
        List<Map<String, Object>> dataList = Lists.newArrayList();
        MapIReadHandler readHandler = new MapIReadHandler(dataList, headerMap);
        try {
            long start = System.currentTimeMillis() / 1000;
            new SaxReadExcel().readExcel(inputStream, Map.class, importParams, readHandler);
            responseMap.put("dataList", dataList);
            long end = System.currentTimeMillis() / 1000;
            log.info("导入excel处理时间：{}秒", end - start);
        } catch (Exception e) {
            //自定义异常判断
            if (e.getMessage().indexOf("不是合法的Excel模板") != -1) {
                responseMap.put("flag", 0); //easypoi模板校验只能校验字段不一致的，数量少的
                responseMap.put("msg", "不是合法的Excel模板");
            } else {
                log.error("导入失败", e);
                throw new RuntimeException("导入失败", e);
            }
        }
        return responseMap;
    }
}