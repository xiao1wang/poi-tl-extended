package com.github.xiao1wang.poitlextended;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.github.xiao1wang.poitlextended.renderData.TableRenderData;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * TODO: 文本测试
 */
public class TableTest {

    public static void main(String[] args) throws Exception {
        Map<String, Object> dataMap = new HashMap<>();
        List<Object[]> list = new ArrayList<>();
        list.add(new String[]{"张三", "博士生"});
        list.add(new String[]{"李四", "硕士"});
        list.add(new String[]{"王五", "本科"});
        dataMap.put("table", new TableRenderData(1, list));
        dataMap.put("cc", "-123");

        ConfigureBuilder builder = Configure.newBuilder();
        // 采用spring El语法
        builder.setElMode(Configure.ELMode.SIMPLE_SPEL_MODE);
        // 得到模板文件
        XWPFTemplate template = XWPFTemplate.compile(
                ChartTest.class.getClassLoader().getResource("templates/template_table.docx").getPath(), builder.build());
        template.render(dataMap);
        FileOutputStream fos = new FileOutputStream("D:\\my_table.docx");
        template.write(fos);
        fos.flush();
        fos.close();
        template.close();
    }
}
