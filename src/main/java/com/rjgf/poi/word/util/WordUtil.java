package com.rjgf.poi.word.util;

import com.deepoove.poi.XWPFTemplate;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

/**
 * word 文档操作工具类
 *
 * @email: xuliandream@gmail.com
 * @author: xula
 * @date: 2020/2/7
 * @time: 11:13
 */
@Slf4j
public class WordUtil {



    /**
     * 关闭流
     * @param template
     * @param out
     */
    public static void closeIo(XWPFTemplate template, OutputStream out,InputStream in) {
        if (template != null) {
            try {
                template.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        if (out != null) {
            try {
                out.flush();
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        if (in != null) {
            try {
                in.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    /**
     * word 转 pdf
     *
     * @param wordInputFile
     * @param pdfOutputFile
     * @throws IOException
     */
    public static void wordToPdf(String wordInputFile, String pdfOutputFile) {
        FileInputStream in = null;
        OutputStream out = null;
        try {
            in = new FileInputStream(wordInputFile);
            XWPFDocument document = new XWPFDocument(in);
            File outFile = new File(pdfOutputFile);
            out = new FileOutputStream(outFile);
            PdfConverter.getInstance().convert(document, out, null);
        } catch (Exception e) {
          log.error("word 转 pdf 失败！",e);
        } finally {
            closeIo(null,out,in);
        }
    }


    /**
     * 转换word to pdf 返回字节数组
     * @param inputStream
     * @return
     */
    public static byte[] wordToPdfByte(InputStream inputStream) {
        try {
            XWPFDocument document = new XWPFDocument(inputStream);
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            PdfConverter.getInstance().convert(document,byteArrayOutputStream,null);
            return byteArrayOutputStream.toByteArray();
        } catch (IOException e) {
            log.error("word 转 字节数据");
            e.printStackTrace();
        }
        return null;
    }
}
