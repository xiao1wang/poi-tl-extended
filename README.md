# poi-tl使用文档

## poi-tl介绍

源码地址：https://github.com/Sayi/poi-tl

使用文档地址：[http://deepoove.com/poi-tl/#_2min%E5%85%A5%E9%97%A8](http://deepoove.com/poi-tl/#_2min入门)

poi-tl的优缺点

| 优点                                                         | 缺点                                                         |
| ------------------------------------------------------------ | ------------------------------------------------------------ |
| 能够直接在word中设置需要输出的数据类型，后台只需关注对应的数据组装即可。能够做到所见即所输出 | 1、针对表格和列表的输出，需要设置对应的样式。   2、不支持图表数据的输出。   3、无法更新对应的目录结构 |

## 思考

基于平时报告的流程，由运营人员提供模板，有没有一种方式，能够在提供的模板基础上，只更改word的数据部分，保留对应的样式？通过poi-tl提供的生成方式，想到使用更新策略，采用poi-tl提供的插件接口，使用更新策略。

## 扩展程序

扩展的功能：表格和列表数据的更新。图表（目前仅提供二维柱状图、条形图、折线图、饼图、面积图，其他图形poi的API还未提供）和目录（还需要优化）的生成

 

### 图表使用方式：

Word中图表的样式

![img](./images/clip_image002.jpg)

代码使用：

 ```java
public static void main(String[] args) throws Exception {
        // 静态数据
        Integer index = 1;
        String title = "个人金额";
        String[] titleArr = {"姓名","销售额"};
        List<Object[]> list = new ArrayList<>();
        list.add(new Object[]{"僵尸软件",12});
        list.add(new Object[]{"Web攻击",13});
        list.add(new Object[]{"木马程序",15});
        list.add(new Object[]{"蠕虫攻击",16});
        List<ChartTypeData> chartList = new ArrayList<>();
        chartList.add(new ChartTypeData(ChartType.PIE,1,titleArr.length - 1));
        ChartRenderData firstChart = new ChartRenderData(index, null, titleArr, list, chartList);
        Map<String, Object> map = new HashMap<>();
        map.put("firstChart", firstChart);

        Configure.ConfigureBuilder builder = Configure.newBuilder();
        // ‘%’自定义的特殊符号，标记对应的变量为表格数据
        builder.addPlugin('%', new ChartRenderPolicy());

        XWPFTemplate template = XWPFTemplate.compile("D:\\template_chart.docx", builder.build());
        template.render(map);
        FileOutputStream fos = new FileOutputStream("D:\\my_chart.docx");
        template.write(fos);
        fos.flush();
        fos.close();
        template.close();
    }
 ```



### 表格使用方式：

Word中表格的样式

![img](./images/clip_image004.jpg)

代码使用：

 ```java
public static void main(String[] args) throws Exception {
        Map<String,Object> dataMap = new HashMap<>();
        List<Object[]> list = new ArrayList<>();
        list.add(new String[]{"张三", "博士生"});
        list.add(new String[]{"李四", "硕士"});
        list.add(new String[]{"王五", "本科"});
        dataMap.put("table", new TableRenderData(1,list));

        Configure.ConfigureBuilder builder = Configure.newBuilder();
        // ‘&’自定义的特殊符号，标记对应的变量为表格数据
        builder.addPlugin('&', new TableRenderPolicy());

        XWPFTemplate template = XWPFTemplate.compile("D:\\template_table.docx", builder.build());
        template.render(dataMap);
        FileOutputStream fos = new FileOutputStream("D:\\my_table.docx");
        template.write(fos);
        fos.flush();
        fos.close();
        template.close();
    }
 ```



### 列表使用方式：

Word中列表的样式

![1583750671875](./images/1583750671875.png)

代码使用：

 ```java
public static void main(String[] args) throws Exception {
        Map<String,Object> dataMap = new HashMap<>();
        dataMap.put("list", Arrays.asList("mmm","测试数据","测试数据1","测试数据2"));

        Configure.ConfigureBuilder builder = Configure.newBuilder();
        // ‘%’自定义的特殊符号，标记对应的变量为列表数据
        builder.addPlugin('%', new ListRenderPolicy());

        XWPFTemplate template = XWPFTemplate.compile("D:\\template_list.docx", builder.build());
        template.render(dataMap);
        FileOutputStream fos = new FileOutputStream("D:\\my_list.docx");
        template.write(fos);
        fos.flush();
        fos.close();
        template.close();
    }
 ```



### 目录使用方式：

由于目录的生成方式特别，因此提供专门的工具类，用于在文档其他数据生成后，单独调用。目前由于目录自动生成后，导致整体的页数发生变化，需要通过手动设置固定的值来确定总页数

代码如下：

```java
public static void main(String[] args) throws Exception {
        Map<String, Object> map = new HashMap<>();
        XWPFTemplate template = XWPFTemplate.compile("D:\\目录.docx").render(map);
        FileOutputStream fos = new FileOutputStream("D:\\my_目录.docx");

        // 需要在全部数据生成完后，再更新目录，这会牵扯到文档的整体页数变化，只能通过固定数值调整
        NiceXWPFDocument doc = template.getXWPFDocument();
        TOCUtils.updateItem2TOC(doc,4, 2);
        template.write(fos);
        fos.flush();
        fos.close();
        template.close();
    }
```

