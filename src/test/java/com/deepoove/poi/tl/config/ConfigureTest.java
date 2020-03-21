package com.deepoove.poi.tl.config;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.config.Configure.AbortHandler;
import com.deepoove.poi.config.Configure.DiscardHandler;
import com.deepoove.poi.config.ConfigureBuilder;
import com.deepoove.poi.data.HyperLinkTextRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.PictureRenderPolicy;
import com.deepoove.poi.policy.TextRenderPolicy;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.tl.source.XWPFTestSupport;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

@DisplayName("Configure test case")
public class ConfigureTest {

    /**
     * [[title]]
     * <p>
     * [[text]]
     * <p>
     * [[%word]]
     * <p>
     * [[姓名]]
     */
    String resource = "src/test/resources/template/config.docx";
    ConfigureBuilder builder = Configure.newBuilder();

    @BeforeEach
    public void init() {
        // 自定义语法以[[开头，以]]结尾
        builder.buildGramer("[[", "]]");
        // 自定义标签text的策略：不是文本，是图片
        builder.bind("text", new PictureRenderPolicy());
        // 添加%语法：%开头的也是文本
        builder.addPlugin('%', new TextRenderPolicy());
    }

    @SuppressWarnings("serial")
    @Test
    public void testPuginAndBind() throws Exception {

        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                put("title", "Hello, poi tl.");
                put("text", new PictureRenderData(100, 120, "src/test/resources/logo.png"));
                put("word", "虽然我是百分号开头，但是我也被自定义成文本了");
            }
        };

        XWPFTemplate template = XWPFTemplate.compile(resource, builder.build());

        template.getElementTemplates().forEach(ele -> {
            assertTrue(ele instanceof RunTemplate);
            RunTemplate runTempalte = (RunTemplate) ele;
            if (runTempalte.getTagName().equals("title")) {
                assertTrue(runTempalte.findPolicy(template.getConfig()) instanceof TextRenderPolicy);
            }
            if (runTempalte.getTagName().equals("text")) {
                assertTrue(runTempalte.findPolicy(template.getConfig()) instanceof PictureRenderPolicy);
            }
            if (runTempalte.getTagName().equals("word")) {
                assertTrue(runTempalte.findPolicy(template.getConfig()) instanceof TextRenderPolicy);
            }
        });

        template.render(datas);

        XWPFTemplate renew = XWPFTestSupport.readNewTemplate(template);
        assertEquals(renew.getElementTemplates().size(), 0);
        renew.close();

    }

    @Test
    public void testDiscardHandler() throws Exception {
        // 没有变量时，保留标签
        builder.setValidErrorHandler(new DiscardHandler());

        XWPFTemplate template = XWPFTemplate.compile(resource, builder.build());
        template.render(new HashMap<String, Object>());

        XWPFTemplate renew = XWPFTestSupport.readNewTemplate(template);
        assertEquals(renew.getElementTemplates().size(), 4);
        renew.close();

    }

    @Test
    public void testAbortHandler() {
        // 没有变量时，无法容忍，抛出异常
        builder.setValidErrorHandler(new AbortHandler());

        assertThrows(RenderException.class,
                () -> XWPFTemplate.compile(resource, builder.build()).render(new HashMap<String, Object>()));
    }

    @Test
    public void testRegex() throws IOException {
        // A~Z,a~z,0~9,_ 组合
        builder.buildGrammerRegex("[\\w]+(\\.[\\w]+)*");

        XWPFTemplate template = XWPFTemplate.compile(resource, builder.build());
        assertEquals(template.getElementTemplates().size(), 3);
        template.close();
    }

    /**
     * {{作者姓名}} {{作者别名}} {{@头像}} {{详情.描述.日期}} {{详情网址}}
     *
     * @throws IOException
     */
    @SuppressWarnings("serial")
    @Test
    public void testSupportChineseAndDot() throws IOException {
        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/config_chinese.docx");
        assertEquals(template.getElementTemplates().size(), 5);

        template.render(new HashMap<String, Object>() {
            {
                put("作者姓名", "Sayi");
                put("作者别名", "卅一");
                put("头像", new PictureRenderData(60, 60, "src/test/resources/sayi.png"));
                put("详情网址", new HyperLinkTextRenderData("http://www.deepoove.com", "http://www.deepoove.com"));
                put("详情", new HashMap<String, Object>() {
                    {
                        put("描述", new HashMap<String, String>() {
                            {
                                put("日期", "2019-05-24");
                            }
                        });
                    }
                });
            }
        });

        XWPFTemplate renew = XWPFTestSupport.readNewTemplate(template);
        assertEquals(renew.getElementTemplates().size(), 0);
        renew.close();
    }

}
