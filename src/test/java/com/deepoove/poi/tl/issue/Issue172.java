package com.deepoove.poi.tl.issue;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.DocxRenderData;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

@DisplayName("Issue172 图表合并")
public class Issue172 {

    @Test
    public void testDocxMerge() throws Exception {

        Map<String, Object> params = new HashMap<String, Object>();

        params.put("docx", new DocxRenderData(new File("src/test/resources/issue/172_MERGE.docx")));

        XWPFTemplate doc = XWPFTemplate.compile("src/test/resources/issue/172.docx");
        doc.render(params);

//        FileOutputStream fos = new FileOutputStream("out_issue_172.docx");
//        doc.write(fos);
//        fos.flush();
//        fos.close();

    }

}
