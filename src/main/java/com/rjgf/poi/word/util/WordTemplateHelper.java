package com.rjgf.poi.word.util;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ELMode;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import lombok.Data;
import lombok.experimental.Accessors;
import lombok.extern.slf4j.Slf4j;

import java.io.FileOutputStream;
import java.util.Map;

/**
 * @email: xuliandream@gmail.com
 * @author: xula
 * @date: 2020/2/21
 * @time: 19:46
 */
@Slf4j
@Data
@Accessors(fluent = true)
public class WordTemplateHelper {
    /**
     * 文件路径
     */
    private String filePath;
    /**
     * 模板地址
     */
    private String tmpFileName;
    /**
     * 输出文件名称
     */
    private String fileName;

    /**
     * 自定义插件
     */
    private Map<String,AbstractRenderPolicy> extendRenderPolicy;

    /**
     * @param data  代替的数据
     */
    public void genFile(Object data) {
        Configure.ConfigureBuilder config = Configure.newBuilder().setElMode(ELMode.SPEL_MODE).buildGramer("${", "}");
        extendRenderPolicy.entrySet().forEach(s-> config.bind(s.getKey(),s.getValue()));
        XWPFTemplate template = XWPFTemplate.compile(filePath + tmpFileName, config.build()).render(data);
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(filePath + fileName);
            template.write(out);
        } catch (Exception e) {
            log.error("word 生成出错，原因：{} ", e.getMessage());
            e.printStackTrace();
        } finally {
            WordUtil.closeIo(template, out,null);
        }
    }


}
