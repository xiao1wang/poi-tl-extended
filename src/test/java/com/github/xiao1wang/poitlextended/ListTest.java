package com.github.xiao1wang.poitlextended;

import com.deepoove.poi.XWPFTemplate;

import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

/**
 * TODO: 文本测试
 */
public class ListTest {

    public static void main(String[] args) throws Exception {
        Map<String, Object> dataMap = new HashMap<>();
        dataMap.put("list", Arrays.asList("mmm", "测试数据", "测试数据1", "测试数据2"));

        // 得到模板文件
        XWPFTemplate template = XWPFTemplate.compile(
                ChartTest.class.getClassLoader().getResource("templates/template_list.docx").getPath());
        template.render(dataMap);
        FileOutputStream fos = new FileOutputStream("D:\\my_list.docx");
        template.write(fos);
        fos.flush();
        fos.close();
        template.close();
    }
}
