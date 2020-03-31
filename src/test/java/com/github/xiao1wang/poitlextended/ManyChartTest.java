package com.github.xiao1wang.poitlextended;

import com.deepoove.poi.XWPFTemplate;
import com.github.xiao1wang.poitlextended.renderData.ChartRenderData;
import com.github.xiao1wang.poitlextended.renderData.ChartType;
import com.github.xiao1wang.poitlextended.renderData.ChartTypeData;
import org.apache.poi.util.IOUtils;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 测试生成图表数据
 */
public class ManyChartTest {

    public static void main(String[] args) throws Exception {
        Map<String, Object> map = new HashMap<>();

        // 静态数据
        String title = "个人金额";
        String[] titleArr = {"姓名", "销售额"};
        List<Object[]> list = new ArrayList<>();
        list.add(new Object[]{"僵尸软件", 12});
        list.add(new Object[]{"Web攻击", 13});
        list.add(new Object[]{"木马程序", 15});
        list.add(new Object[]{"蠕虫攻击", 16});
        List<ChartTypeData> chartList = new ArrayList<>();
        chartList.add(new ChartTypeData(ChartType.DOUGHNUT, 1, 1));
        ChartRenderData firstchart = new ChartRenderData(null, null, list, chartList);
        map.put("firstchart", firstchart);

        String[] secondTitleArr = {"", "数量"};
        List<Object[]> secondList = new ArrayList<>();
        secondList.add(new Object[]{"僵尸软件", 12});
        secondList.add(new Object[]{"Web攻击", 13});
        secondList.add(new Object[]{"木马程序", 15});
        secondList.add(new Object[]{"蠕虫攻击", 16});
        List<ChartTypeData> secondChartList = new ArrayList<>();
        secondChartList.add(new ChartTypeData(ChartType.LINE, 1, 1));
        ChartRenderData secondchart = new ChartRenderData(null, null, secondList, secondChartList);
        map.put("secondchart", secondchart);

        String[] threeTitleArr = { "", "数量" };
        List<Object[]> threeList = new ArrayList<>();
        threeList.add(new Object[] { "僵尸软件", 12 });
        threeList.add(new Object[] { "Web攻击", 13 });
        threeList.add(new Object[] { "木马程序", 15 });
        threeList.add(new Object[] { "蠕虫攻击", 16 });
        List<ChartTypeData> threeChartList = new ArrayList<>();
        threeChartList.add(new ChartTypeData(ChartType.BAR, 1, 1));
        ChartRenderData threechart = new ChartRenderData("测试", threeTitleArr, threeList, threeChartList);
        map.put("threechart", threechart);

        String[] fourTitleArr = { "", "数量" };
        List<Object[]> fourList = new ArrayList<>();
        fourList.add(new Object[] { "僵尸软件", 12 });
        fourList.add(new Object[] { "Web攻击", 13 });
        fourList.add(new Object[] { "木马程序", 15 });
        fourList.add(new Object[] { "蠕虫攻击", 16 });
        List<ChartTypeData> foutChartList = new ArrayList<>();
        foutChartList.add(new ChartTypeData(ChartType.DOUGHNUT, 1, 1));
        ChartRenderData fourchart = new ChartRenderData("测试", fourTitleArr, fourList, foutChartList);
        map.put("fourchart", fourchart);

        // 得到模板文件
        XWPFTemplate template = XWPFTemplate.compile(ManyChartTest.class.getClassLoader().getResource("templates/many_chart.docx").getPath());

        template.render(map);

        FileOutputStream fos = new FileOutputStream("D:\\my_many_chart.docx");
        template.write(fos);
        fos.flush();
        fos.close();
        String context = template.getXWPFDocument().getCharts().get(0).getCTChart().xmlText();
        IOUtils.copy(new ByteArrayInputStream(context.getBytes()), new File("D:\\my_many_chart.xml"));
        template.close();
    }
}
