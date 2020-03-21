package com.deepoove.poi.tl.issue;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.tl.source.XWPFTestSupport;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.HashMap;

import static org.junit.jupiter.api.Assertions.assertEquals;

@DisplayName("Issue111 模板生成模板")
public class Issue111 {

    @SuppressWarnings("serial")
    @Test
    public void testCRBR() throws Exception {
        StringBuffer sb = new StringBuffer();
        for (int i = 0; i < 5; i++) {
            String info = "姓名{{name" + i + "}}，年龄：{{age" + i + "}}。";
            sb.append(info).append("\n");
        }
        final String Info = sb.toString();
        XWPFTemplate templateRule = XWPFTemplate.compile("src/test/resources/issue/111.docx")
                .render(new HashMap<String, Object>() {
                    {
                        put("title", Info);
                    }
                });

        templateRule.writeToFile("out_111temp.docx");
        XWPFTemplate template = XWPFTemplate.compile("out_111temp.docx").render(new HashMap<String, Object>() {
            {
                for (int j = 0; j < 5; j++) {
                    put("name" + j, "测试姓名" + j);
                    put("age" + j, "测试年龄" + j);
                }
            }
        });

        // template.writeToFile("out_issue_111.docx");
        XWPFDocument document = XWPFTestSupport.readNewDocument(template);
        XWPFParagraph paragraph = document.getParagraphArray(0);
        assertEquals(paragraph.getText(), "姓名测试姓名0，年龄：测试年龄0。\n" +
                "姓名测试姓名1，年龄：测试年龄1。\n" +
                "姓名测试姓名2，年龄：测试年龄2。\n" +
                "姓名测试姓名3，年龄：测试年龄3。\n" +
                "姓名测试姓名4，年龄：测试年龄4。\n" +
                "姓名测试姓名0，年龄：测试年龄0。\n" +
                "姓名测试姓名1，年龄：测试年龄1。\n" +
                "姓名测试姓名2，年龄：测试年龄2。\n" +
                "姓名测试姓名3，年龄：测试年龄3。\n" +
                "姓名测试姓名4，年龄：测试年龄4。\n");

        document.close();

        new File("out_111temp.docx").deleteOnExit();
    }

}
