package com.github.xiao1wang.poitlextended;


import com.deepoove.poi.XWPFTemplate;
import com.github.xiao1wang.poitlextended.renderData.ChartRenderData;
import com.github.xiao1wang.poitlextended.renderData.ChartType;
import com.github.xiao1wang.poitlextended.renderData.ChartTypeData;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 测试生成图表数据
 */
public class ChartTest {

    public static void main(String[] args) throws Exception {
        // 静态数据
        String title = "个人金额";
        String[] titleArr = {"姓名", "销售额"};
        List<Object[]> list = new ArrayList<>();
        list.add(new Object[]{"僵尸软件", 12});
        list.add(new Object[]{"Web攻击", 13});
        list.add(new Object[]{"木马程序", 15});
        list.add(new Object[]{"蠕虫攻击", 16});
        List<ChartTypeData> chartList = new ArrayList<>();
        chartList.add(new ChartTypeData(ChartType.PIE, 1, titleArr.length - 1));
        ChartRenderData firstChart = new ChartRenderData(null, titleArr, list, chartList);
        Map<String, Object> map = new HashMap<>();
        map.put("firstChart", firstChart);
        //map.put("firstChart", null);

        // 得到模板文件
        XWPFTemplate template = XWPFTemplate.compile(
                ChartTest.class.getClassLoader().getResource("templates/template_chart.docx").getPath());
        template.render(map);
        FileOutputStream fos = new FileOutputStream("D:\\my_chart.docx");
        template.write(fos);
        fos.flush();
        fos.close();
        template.close();
    }
}
