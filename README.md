word，excel，pdf 操作工具模块

1: excel的文档操作，读取和写入

2: word基于模板生成最终文件

3: word转pdf功能

使用的技术框架

* excel 操作框架 [easyExcel](https://alibaba-easyexcel.github.io/index.html)
* word 生成框架 [poi-tl](http://deepoove.com/poi-tl/)
* word 转 pdf 框架 [xwpf](https://stackoverflow.com/questions/51440312/docx-to-pdf-converter-in-java)

目前有以下几个功能

1. excel 的导入，实现批量文件处理，实现方式

   ```
   public void batchRead() throws FileNotFoundException { 
           EasyExcelUtil.batchRead(new FileInputStream(new File("e://test.txt")),new ICallback() {
               @Override
               public void todo(List data) {
                   // 处理对应的业务
               }
           },10000); // 设置每次读取的条数
       }
   }
   ```

2. 同样也支持获取文件的所有数据，但是不建议这么做，如果数据量较大会导致内存溢出

   ```
   public void readAll() throws FileNotFoundException {
           EasyExcelUtil.readAllData(new FileInputStream(new File("e://test.txt")), Map.class); // 指定你的类的class
   }
   ```

3. 导出支持设置表头来输出指定数据

   ```
   public void export(HttpServletResponse response) throws Exception {
           Set<String> headers = new HashSet<>(); # 指定输出的数据字段，字段名称需要跟数据实体类的属性一样
           headers.add("name");
           headers.add("age");
           List<Map> data = new ArrayList<>(); # 数据集合
           EasyExcelUtil.export(response,data,headers,"测试");
   }
   ```

   