package com.github.xiao1wang.poitlextended;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;

import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

/**
 * TODO: 文本测试
 */
public class TlTest {

    public static void main(String[] args) throws Exception {
        Map<String, Object> dataMap = new HashMap<>();


        RowRenderData header = RowRenderData.build(
                new TextRenderData("000000", "序号"),
                new TextRenderData("000000", "城市"),
                new TextRenderData("000000", "受害IP个数")
        );

        RowRenderData row0 = RowRenderData.build("1", "济南", "173");
        RowRenderData row1 = RowRenderData.build("2", "青岛", "1");
        RowRenderData row2 = RowRenderData.build("3", "威海", "1");

        dataMap.put("table", new MiniTableRenderData(header, Arrays.asList(row0, row1, row2)));


        // 得到模板文件
        XWPFTemplate template = XWPFTemplate.compile(
                ChartTest.class.getClassLoader().getResource("templates/test-tl.docx").getPath());
        template.render(dataMap);
        FileOutputStream fos = new FileOutputStream("D:\\my_table.docx");
        template.write(fos);
        fos.flush();
        fos.close();
        template.close();
    }
}
