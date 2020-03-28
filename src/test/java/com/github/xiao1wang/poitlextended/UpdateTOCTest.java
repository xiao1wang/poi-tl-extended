package com.github.xiao1wang.poitlextended;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import com.github.xiao1wang.poitlextended.util.CustomerTOC;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * TODO : 写点注释吧
 */
public class UpdateTOCTest {

    public static void main(String[] args) throws Exception {
        Map<String, Object> map = new HashMap<>();
        ConfigureBuilder builder = Configure.newBuilder();
        // 采用spring El语法
        builder.setElMode(Configure.ELMode.SIMPLE_SPEL_MODE);
        XWPFTemplate template = XWPFTemplate.compile(ChartTest.class.getClassLoader().getResource("templates/test_doc.docx").getPath(), builder.build());

        template.render(map);
        NiceXWPFDocument doc = template.getXWPFDocument();
        CustomerTOC.automaticGenerateTOC(3, "toc", doc, 2);
        OutputStream out = new FileOutputStream("d:\\my_doc_customer.docx");
        doc.write(out);
        out.close();
    }
}
