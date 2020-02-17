package com.rjgf.poi.word.util;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

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


    public static void main(String[] args) throws IOException {
        Map<String, Object> map = new HashMap<>();
        map.put("province", "浙江省");
        map.put("city", "金华市");
        map.put("mapImg", new PictureRenderData(120, 120, "E:\\img\\mapImg.png"));
        map.put("siteData", new PictureRenderData(120, 120, "E:\\img\\siteData.png"));
        // 序号	ENBID	运营商	ear_fcn	推测站点经度	推测站点纬度	备注
        RowRenderData header = RowRenderData.build(new TextRenderData("FFFFFF", "序号"), new TextRenderData("FFFFFF", "ENBID"),
                new TextRenderData("FFFFFF", "运营商"), new TextRenderData("FFFFFF", "ear_fcn"), new TextRenderData("FFFFFF", "推测站点经度"),
                new TextRenderData("FFFFFF", "推测站点纬度"), new TextRenderData("FFFFFF", "备注"));
        RowRenderData row0 = RowRenderData.build("张三", "研究生", "张三", "研究生", "张三", "研究生", "张三");
        map.put("addDataTable", new MiniTableRenderData(header, Arrays.asList(row0)));

        String filePath = "E:\\Users\\xulia\\润建\\app 大数据\\word 模块\\";
        String tmpFileName = "template.docx";
        String fileName = "out_template.docx";
        WordUtil.genFile(filePath, tmpFileName, fileName, map);


        wordToPdf("e://test.docx","e://test1.pdf");
    }


    /**
     * @param filePath    文件路径
     * @param tmpFileName 模板文件名称
     * @param fileName    导出的文件名
     * @param data        代替的数据
     */
    public static void genFile(String filePath, String tmpFileName, String fileName, Map<String, Object> data) {
        XWPFTemplate template = XWPFTemplate.compile(filePath + tmpFileName).render(data);
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(filePath + fileName);
            template.write(out);
        } catch (Exception e) {
            log.error("word 生成出错，原因：{} ", e.getMessage());
            e.printStackTrace();
        } finally {
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
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
