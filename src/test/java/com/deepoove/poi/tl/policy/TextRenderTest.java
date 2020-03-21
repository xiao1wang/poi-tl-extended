package com.deepoove.poi.tl.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.HyperLinkTextRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.policy.TextRenderPolicy;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

@DisplayName("Text Render test case")
public class TextRenderTest {

    @SuppressWarnings("serial")
    @Test
    public void testTextRender() throws Exception {

        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                put("title", "Hello, poi tl.");
                put("text", new TextRenderData("28a745", "我是\n绿色且换行的文字"));

                // 超链接
                HyperLinkTextRenderData hyperLinkTextRenderData = new HyperLinkTextRenderData("Deepoove website.",
                        "http://www.deepoove.com");
                hyperLinkTextRenderData.getStyle().setBold(true);
                put("link", hyperLinkTextRenderData);
                put("maillink", new HyperLinkTextRenderData("发邮件给作者", "mailto:adasai90@gmail.com?subject=poi-tl"));

                // 指定文本样式
                TextRenderData textRenderData = new TextRenderData("just deepoove.");
                Style style = new Style("FF5722");
                textRenderData.setStyle(style);
                style.setBold(true);
                style.setFontSize(48);
                style.setItalic(true);
                style.setStrike(true);
                style.setUnderLine(true);
                style.setFontFamily("微软雅黑");
                style.setCharacterSpacing(20);
                style.setHighlightColor(STHighlightColor.DARK_GREEN);
                put("word", textRenderData);

                // 换行
                put("newline", "hi\n\n\n\n\nhello");

                // 从文件读取文字
                File file = new File("src/test/resources/template/render_text_word.txt");
                FileInputStream in = new FileInputStream(file);
                int size = in.available();
                byte[] buffer = new byte[size];
                in.read(buffer);
                in.close();
                put("newline_from_file", new String(buffer));
            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/render_text.docx").render(datas);

        FileOutputStream out = new FileOutputStream("out_render_text.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();

    }

    @Test
    public void testNewLine() {
        String text = "hello\npoi-tl";
        String text1 = "hello\n\npoi-tl";
        String text2 = "hello\n\n";
        String text3 = "\n\npoi-tl";
        String text4 = "\n\n\n\n";
        String text5 = "hi\n\n\n\nwhat\nis\n\n\nthis";

        String regexLine = TextRenderPolicy.Helper.REGEX_LINE_CHARACTOR;

        assertEquals(Arrays.toString(text.split(regexLine, -1)), "[hello, poi-tl]");
        assertEquals(Arrays.toString(text1.split(regexLine, -1)), "[hello, , poi-tl]");
        assertEquals(Arrays.toString(text2.split(regexLine, -1)), "[hello, , ]");
        assertEquals(Arrays.toString(text3.split(regexLine, -1)), "[, , poi-tl]");
        assertEquals(Arrays.toString(text4.split(regexLine, -1)), "[, , , , ]");
        assertEquals(Arrays.toString(text5.split(regexLine, -1)), "[hi, , , , what, is, , , this]");
    }

}
