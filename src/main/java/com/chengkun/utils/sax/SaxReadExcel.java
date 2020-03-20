package com.chengkun.utils.sax;
/**
 * sungrow all right reserved
 **/

import cn.afterturn.easypoi.excel.entity.ImportParams;
import cn.afterturn.easypoi.excel.imports.sax.parse.ISaxRowRead;
import cn.afterturn.easypoi.excel.imports.sax.parse.SaxRowRead;
import cn.afterturn.easypoi.exception.excel.ExcelImportException;
import cn.afterturn.easypoi.handler.inter.IReadHandler;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.Iterator;

/**
 * @Description 基于SAX Excel大数据读取
 * easypoi的sax有bug，空单元格会出现问题
 * @Author chengkun
 * @Date 2020/3/20 18:56
 **/
@Log4j2
public class SaxReadExcel {
    public void readExcel(InputStream inputstream, Class<?> pojoClass, ImportParams params, IReadHandler handler) {

        try {
            OPCPackage opcPackage = OPCPackage.open(inputstream);
            readExcel(opcPackage, pojoClass, params, null, handler);
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new ExcelImportException(e.getMessage(), e);
        }
    }

    private void readExcel(OPCPackage opcPackage, Class<?> pojoClass, ImportParams params,
                           ISaxRowRead rowRead, IReadHandler handler) {
        try {
            XSSFReader xssfReader = new XSSFReader(opcPackage);
            SharedStringsTable sharedStringsTable = xssfReader.getSharedStringsTable();
            StylesTable stylesTable = xssfReader.getStylesTable();
            if (rowRead == null) {
                rowRead = new SaxRowRead(pojoClass, params, handler);
            }
            XMLReader parser = fetchSheetParser(sharedStringsTable, stylesTable, rowRead);
            Iterator<InputStream> sheets = xssfReader.getSheetsData();
            int sheetIndex = 0;
            while (sheets.hasNext() && sheetIndex < params.getSheetNum() + params.getStartSheetIndex()) {
                if (sheetIndex < params.getStartSheetIndex()) {
                    sheets.next();
                } else {
                    InputStream sheet = sheets.next();
                    InputSource sheetSource = new InputSource(sheet);
                    parser.parse(sheetSource);
                    sheet.close();
                }
                sheetIndex++;
            }
            if (handler != null) {
                handler.doAfterAll();
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
            throw new ExcelImportException(e.getMessage(), e);
        }
    }

    private XMLReader fetchSheetParser(SharedStringsTable sharedStringsTable, StylesTable stylesTable,
                                       ISaxRowRead rowRead) throws SAXException {
        XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
        ContentHandler handler = new SheetHandler(sharedStringsTable, stylesTable, rowRead);
        parser.setContentHandler(handler);
        return parser;
    }
}