package com.deepoove.poi.tl.issue;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.DocxRenderData;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

@DisplayName("Issue225 子文档形状")
public class Issue225 {

    @SuppressWarnings("serial")
    @Test
    public void testDocxTemplateRender() throws Exception {

        Map<String, Object> datas = new HashMap<String, Object>() {
            {

                put("date", "2019-05-06");
                put("first", new DocxRenderData(new File("src/test/resources/issue/225.docx"),
                        Arrays.asList(1, 2, 3)));

            }
        };

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/issue/244.docx")
                .render(datas);

        FileOutputStream out = new FileOutputStream("out_issue_225.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();

    }

}
