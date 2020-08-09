package com.rjgf.poi.excel.util;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.*;

/**
 * @email: xuliandream@gmail.com
 * @author: xula
 * @date: 2020/8/9
 * @time: 18:23
 */
public class ExcelTest {

    public void batchRead() throws FileNotFoundException {

        EasyExcelUtil.batchRead(new FileInputStream(new File("e://test.txt")),new ICallback() {

            @Override
            public void todo(List data) {
                // 处理对应的业务
            }
        },10000); // 设置每次读取的条数
    }

    public void readAll() throws FileNotFoundException {
        EasyExcelUtil.readAllData(new FileInputStream(new File("e://test.txt")), Map.class); // 指定你的类的class
    }

    public void export(HttpServletResponse response) throws Exception {
        Set<String> headers = new HashSet<>();
        headers.add("name");
        headers.add("age");
        List<Map> data = new ArrayList<>();
        EasyExcelUtil.export(response,data,headers,"测试");
    }
}
