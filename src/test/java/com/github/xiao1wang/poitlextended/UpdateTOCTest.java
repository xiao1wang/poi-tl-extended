package com.github.xiao1wang.poitlextended;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import com.github.xiao1wang.poitlextended.util.TOCUtils;

import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * TODO : 写点注释吧
 */
public class UpdateTOCTest {

    public static void main(String[] args) throws Exception {
        Map<String, Object> map = new HashMap<>();
        XWPFTemplate template = XWPFTemplate.compile(
                ChartTest.class.getClassLoader().getResource("templates/template-toc.docx").getPath());
        FileOutputStream fos = new FileOutputStream("D:\\my_目录.docx");

        // 需要在全部数据生成完后，再更新目录，这会牵扯到文档的整体页数变化，只能通过固定数值调整
        NiceXWPFDocument doc = template.getXWPFDocument();
        TOCUtils.updateItem2TOC(doc, 4, 2);
        template.write(fos);
        fos.flush();
        fos.close();
        template.close();
    }
}
