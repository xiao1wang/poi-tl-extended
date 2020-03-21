package com.deepoove.poi.tl.render;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.HyperLinkTextRenderData;
import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.data.NumbericRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.Style;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@DisplayName("Foreach template test case")
public class IterableTemplateTest {

    @SuppressWarnings("serial")
    @Test
    public void testIterableWithStyle() throws Exception {
        List<Map<String, Object>> sections = new ArrayList<>();
        sections.add(new HashMap<String, Object>() {
            {
                put("title", "1.1 小节");
                put("word", "这是第一小节的内容");

            }
        });
        sections.add(new HashMap<String, Object>() {
            {
                put("title", "1.2 小节");
                put("word", "这是第二小节的内容");

            }
        });
        List<Map<String, Object>> chapters = new ArrayList<>();
        chapters.add(new HashMap<String, Object>() {
            {
                put("title", "第一章");
                put("name", "Sayi");
                put("sections", sections);

            }
        });
        chapters.add(new HashMap<String, Object>() {
            {
                put("title", "第二章");
                put("name", "Deepoove");

            }
        });
        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                put("title", "poi-tl");
                put("chapters", chapters);

            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/iterable_foreach_withstyle.docx");

        template.render(datas);
        template.writeToFile("out_iterable_foreach_withstyle.docx");
    }

    @SuppressWarnings("serial")
    @Test
    public void testForeach() throws Exception {
        List<Map<String, Object>> addrs = new ArrayList<>();
        addrs.add(new HashMap<String, Object>() {
            {
                put("position", "Hangzhou,China");
            }
        });
        addrs.add(new HashMap<String, Object>() {
            {
                put("position", "Shanghai,China");
            }
        });

        List<Map<String, Object>> users = new ArrayList<>();
        users.add(new HashMap<String, Object>() {
            {
                put("name", "Sayi");
                put("addrs", addrs);
            }
        });
        users.add(new HashMap<String, Object>() {
            {
                put("name", "Deepoove");
            }
        });
        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                put("title", "poi-tl");
                put("users", users);
            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/iterable_foreach1.docx");
        template.render(datas);
        template.writeToFile("out_iterable_foreach1.docx");
    }

    @SuppressWarnings("serial")
    @Test
    public void testHyperField() throws Exception {
        List<Map<String, Object>> addrs = new ArrayList<>();
        addrs.add(new HashMap<String, Object>() {
            {
                put("position", "Hangzhou,China");

            }
        });
        addrs.add(new HashMap<String, Object>() {
            {
                put("position", "Shanghai,China");

            }
        });

        List<Map<String, Object>> users = new ArrayList<>();
        users.add(new HashMap<String, Object>() {
            {
                put("name", "Sayi");
                put("addrs", addrs);

            }
        });
        users.add(new HashMap<String, Object>() {
            {
                put("name", new HyperLinkTextRenderData("Deepoove website.", "http://www.google.com"));

            }
        });
        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                put("users", users);

            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/iterable_hyperlink.docx");
        template.render(datas);
        template.writeToFile("out_iterable_hyperlink.docx");
    }

    @SuppressWarnings("serial")
    @Test
    @DisplayName("using all gramer together")
    public void testTogether() throws Exception {
        RowRenderData row0 = RowRenderData.build(new HyperLinkTextRenderData("张三", "http://deepoove.com"),
                new TextRenderData("1E915D", "研究生"));

        RowRenderData row1 = RowRenderData.build("李四", "博士");

        RowRenderData row2 = RowRenderData.build("王五", "博士后");

        final TextRenderData textRenderData = new TextRenderData("负责生产BUG，然后修复BUG，同时有效实施招聘行为");
        Style style = new Style();
        style.setFontSize(10);
        style.setColor("7F7F7F");
        style.setFontFamily("微软雅黑");
        textRenderData.setStyle(style);
        List<Map<String, Object>> addrs = new ArrayList<>();
        addrs.add(new HashMap<String, Object>() {
            {
                put("position", "Hangzhou,China");

            }
        });
        addrs.add(new HashMap<String, Object>() {
            {
                put("position", "Shanghai,China");

            }
        });

        List<Map<String, Object>> users = new ArrayList<>();
        users.add(new HashMap<String, Object>() {
            {
                put("name", "Sayi");
                put("addrs", addrs);
                put("list", new NumbericRenderData(NumbericRenderData.FMT_DECIMAL,
                        Arrays.asList(textRenderData, textRenderData)));
                put("image", new PictureRenderData(120, 120, "src/test/resources/sayi.png"));
                put("table", new MiniTableRenderData(Arrays.asList(row0, row1, row2)));

            }
        });
        users.add(new HashMap<String, Object>() {
            {
                put("name", "Deepoove");
                put("list", new NumbericRenderData(NumbericRenderData.FMT_DECIMAL,
                        Arrays.asList(textRenderData, textRenderData)));
                put("image", new PictureRenderData(120, 120, "src/test/resources/sayi.png"));

            }
        });
        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                put("title", "poi-tl");
                put("users", users);
                put("thisref", Arrays.asList("good", "people"));

            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/iterable_foreach2.docx");
        template.render(datas);
        template.writeToFile("out_iterable_foreach2.docx");
    }

}
