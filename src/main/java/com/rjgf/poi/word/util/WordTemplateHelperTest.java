package com.rjgf.poi.word.util;

import com.deepoove.poi.data.PictureRenderData;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.util.List;

/**
 * @email: xuliandream@gmail.com
 * @author: xula
 * @date: 2020/2/21
 * @time: 19:46
 */
@Slf4j
public class WordTemplateHelperTest {

    @Data
    static class NetType {
        private String name;
        private Integer age;
    }

    @Data
    static
    class SpELData {
        private List<NetType> netTypes;
        private String test;
        private PictureRenderData img;
    }

    public static void main(String[] args) throws IOException {
//        SpELData data = new SpELData();
//        data.setTest("测试");
//        List<NetType> list = new ArrayList<>();
//        for (int i = 0; i < 2; i++) {
//            NetType netType = new NetType();
//            netType.setName(RandomUtil.randomString(4));
//            netType.setAge(RandomUtil.randomInt(100));
//            list.add(netType);
//        }
//        data.setNetTypes(list);
//        SpELRenderDataCompute spELRenderDataCompute = new SpELRenderDataCompute(data);
//        System.out.println(spELRenderDataCompute.compute("netTypes[0].age"));


//        map.put("province", "浙江省");
//        map.put("city", "金华市");
//        data.setImg(new PictureRenderData(120, 120, "E:\\Users\\xulia\\润建\\app 大数据\\word 模块\\img\\e23f2f6c-74e0-485d-bb59-c7c8b554abc7.png"));
//        map.put("siteData", new PictureRenderData(120, 120, "E:\\img\\siteData.png"));
//        // 序号	ENBID	运营商	ear_fcn	推测站点经度	推测站点纬度	备注
//        RowRenderData header = RowRenderData.build(new TextRenderData("FFFFFF", "序号"), new TextRenderData("FFFFFF", "ENBID"),
//                new TextRenderData("FFFFFF", "运营商"), new TextRenderData("FFFFFF", "ear_fcn"), new TextRenderData("FFFFFF", "推测站点经度"),
//                new TextRenderData("FFFFFF", "推测站点纬度"), new TextRenderData("FFFFFF", "备注"));
//        RowRenderData row0 = RowRenderData.build("张三", "研究生", "张三", "研究生", "张三", "研究生", "张三");
//        map.put("addDataTable", new MiniTableRenderData(header, Arrays.asList(row0)));


//        Map<String, AbstractRenderPolicy> abstractRenderPolicyMap = new HashMap<>();
//
//        abstractRenderPolicyMap.put("convertDataList",new NetSegRenderPolicy());
//
//        List<OptNetSegDataTemp.convertData> list = new ArrayList<>();
//        for (int i = 0; i < 4 ; i++) {
//            OptNetSegDataTemp.convertData convertData = new OptNetSegDataTemp.convertData();
//            convertData.setConvertRatio(RandomUtil.randomDouble(10));
//            convertData.setIncreaseRatio(RandomUtil.randomDouble(20));
//            convertData.setNetSeg(RandomUtil.randomString(4));
//            list.add(convertData);
//        }
//
//        String filePath = "E:\\Users\\xulia\\润建\\app 大数据\\word 模块\\";
//        String tmpFileName = "区域范围.docx";
//        String fileName = "test-1.docx";
//
//        Test test = new Test();
//        test.setConvertDataList(list);
//
//        WordTemplateHelper wordTemplateHelper = new WordTemplateHelper().
//                fileName(fileName).filePath(filePath).tmpFileName(tmpFileName).extendRenderPolicy(abstractRenderPolicyMap);
//        wordTemplateHelper.genFile(test);
//        wordToPdf("e://test.docx","e://test1.pdf");
    }

}
