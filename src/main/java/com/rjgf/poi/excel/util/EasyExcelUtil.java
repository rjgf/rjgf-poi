package com.rjgf.poi.excel.util;


import cn.hutool.core.date.DateUtil;
import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.rjgf.poi.excel.convert.LocalDateStringConvert;
import com.rjgf.poi.excel.convert.LocalDateTimeStringConvert;
import com.rjgf.poi.excel.util.ExcelConstant;
import com.rjgf.poi.excel.util.ExcelListener;
import com.rjgf.poi.excel.util.ICallback;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.compress.utils.IOUtils;
import org.apache.poi.ss.formula.functions.T;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.Date;
import java.util.List;
import java.util.Set;


/**
 * easyexcel 操作类
 *
 * @author xula
 * @date 2020-06-11
 */
@Slf4j
public class EasyExcelUtil implements ExcelConstant {


    /**
     * 导出excel
     *
     * @param response
     * @param inputStream 输入流
     * @param exportName  导出的文件名
     * @throws IOException
     */
    public static void export(HttpServletResponse response, InputStream inputStream, String exportName) throws IOException {
        response.setContentType("application/vnd.ms-easyExcel");
        response.setCharacterEncoding("utf-8");
        String fileName = new String(new String(exportName + "_" + DateUtil.formatDate(new Date()) + ".xlsx").getBytes("UTF-8"), "ISO8859-1");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName);
        OutputStream outputStream = null;
        try {
            outputStream = response.getOutputStream();
            IOUtils.copy(inputStream, outputStream);
        } catch (Exception e) {
            log.error("导出文件失败！", e);
        } finally {
            if (null != outputStream) {
                outputStream.flush();
                outputStream.close();
            }
            if (null != inputStream) {
                inputStream.close();
            }
        }
    }

    /**
     * 导出excel
     *
     * @param response
     * @param dataList   数据列表
     * @param exportName 使用 utf-8 编码，不需要带后缀,导出的文件名称
     */
    public static <T> void export(HttpServletResponse response, List<T> dataList, String exportName) throws Exception {
        export(response,dataList,null,exportName);
    }

    /**
     * 导出excel
     *
     * @param response
     * @param headers 指定的表头列表
     * @param exportName 导出的文件名
     * @throws IOException
     */
    public static <T> void export(HttpServletResponse response, List<T> dataList, Set<String> headers, String exportName) throws Exception {
        if (CollectionUtils.isEmpty(dataList)) {
            throw new Exception("数据为空，导出无意义！");
        }
        T t = dataList.get(0);
        response.setContentType("application/octet-stream");
        response.setCharacterEncoding("utf-8");
        String fileName = new String(new String(exportName + "_" + DateUtil.formatDate(new Date()) + ".xlsx").getBytes("UTF-8"), "ISO8859-1");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName);
        try {
            // 如果这里想使用03 则 传入excelType参数即可
            ExcelWriterBuilder excelWriterBuilder = new ExcelWriterBuilder();
            excelWriterBuilder.file(response.getOutputStream());
            excelWriterBuilder.head(t.getClass());
            // 只输出，用户指定的表头数据
            if (CollectionUtils.isNotEmpty(headers)) {
                excelWriterBuilder.includeColumnFiledNames(headers);
            }
            // 由于 Java 8 localDate 和 localDateTime 不支持转化，需要添加自定义转化器
            excelWriterBuilder.registerConverter(new LocalDateStringConvert());
            excelWriterBuilder.registerConverter(new LocalDateTimeStringConvert());
            // 不设置表头和内容样式
            HorizontalCellStyleStrategy horizontalCellStyleStrategy =
                    new HorizontalCellStyleStrategy(null, (WriteCellStyle) null);
            excelWriterBuilder.registerWriteHandler(horizontalCellStyleStrategy).sheet().doWrite(dataList);
        } catch (Exception e) {
            log.error("导出文件失败！", e);
        }
    }

    /**
     * 分批读取excel
     *
     * @param inputStream 输入流
     * @param iCallback   分批回调方法
     */
    public static void batchRead(InputStream inputStream, ICallback iCallback) {
        batchReadSheet(inputStream, 0, iCallback, BATCH_COUNT);
    }


    /**
     * 分批读取excel
     *
     * @param inputStream 输入流
     * @param iCallback   分批回调方法
     * @param batchCount  批量执行数据数量
     */
    public static void batchRead(InputStream inputStream, ICallback iCallback, Integer batchCount) {
        batchReadSheet(inputStream, 0, iCallback, batchCount);
    }


    /**
     * 分批读取excel
     *
     * @param inputStream 输入流
     * @param iCallback   分批回调方法
     * @param batchCount  批量执行数据数量
     */
    public static void batchReadSheet(InputStream inputStream, Integer sheetNo, ICallback iCallback, Integer batchCount) {
        ExcelListener excelListenerV2 = new ExcelListener();
        excelListenerV2.setiCallback(iCallback);
        excelListenerV2.setBatchCount(batchCount);
        EasyExcel.read(inputStream, iCallback.getTClass(), excelListenerV2).sheet(sheetNo).doRead();
    }


    /**
     * 分批读取excel
     *
     * @param inputStream 输入流
     * @param iCallback   回调方法
     * @param batchCount  批量执行数据数量
     */
    public static void batchReadAllSheet(InputStream inputStream, ICallback iCallback, Integer batchCount) {
        ExcelListener excelListenerV2 = new ExcelListener();
        excelListenerV2.setiCallback(iCallback);
        excelListenerV2.setBatchCount(batchCount);
        EasyExcel.read(inputStream, iCallback.getTClass(), excelListenerV2).doReadAll();
    }


    /**
     * 读取excel
     *
     * @param inputStream 输入流
     * @param tClass      自定类型
     */
    public static <T> List<T> readAllData(InputStream inputStream, Class<T> tClass) {
        return readAllData(inputStream, tClass, 0);
    }


    /**
     * 读取excel
     *
     * @param inputStream 输入流
     * @param tClass      自定类型
     * @param sheetNo     指定sheet
     */
    public static <T> List<T> readAllData(InputStream inputStream, Class<T> tClass, Integer sheetNo) {
        ExcelListener excelListener = new ExcelListener();
        EasyExcel.read(inputStream, tClass, excelListener).sheet(sheetNo).doRead();
        return excelListener.getDatas();
    }

    /**
     * 读取excel
     *
     * @param inputStream 输入流
     * @param tClass      自定类型
     */
    public static <T> List<T> readAllSheetData(InputStream inputStream, Class<T> tClass) {
        ExcelListener excelListener = new ExcelListener();
        EasyExcel.read(inputStream, tClass, excelListener).doReadAll();
        return excelListener.getDatas();
    }


    /**
     * 读取excel
     *
     * @param inputStream   输入流
     * @param tClass        自定类型
     * @param excelTypeEnum 设置文件类型
     */
    public static <T> List<T> readAllSheetData(InputStream inputStream, Class<T> tClass, ExcelTypeEnum excelTypeEnum) {
        ExcelListener excelListener = new ExcelListener();
        EasyExcel.read(inputStream, tClass, excelListener).excelType(excelTypeEnum).doReadAll();
        return excelListener.getDatas();
    }
}