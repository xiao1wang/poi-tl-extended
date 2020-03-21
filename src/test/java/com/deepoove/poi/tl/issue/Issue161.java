package com.deepoove.poi.tl.issue;

import com.deepoove.poi.XWPFTemplate;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

@DisplayName("Issue161 enter回车")
public class Issue161 {

    @Test
    public void testEmptyRun() throws Exception {
        Map<String, String> dataMap = new HashMap<String, String>();
        dataMap.put("projectName", "t项目名称test");
        dataMap.put("designDeptName", "t设计单位test");
        dataMap.put("applyUnitName", "t申请单位test");
        dataMap.put("ownerDeptName", "t建设单位test");
        dataMap.put("optimizeReason", "t变更原因test");
        dataMap.put("optimizeChangeName", "t审查方案test");
        dataMap.put("changeType", "t变更类型test");
        dataMap.put("moneyChange", "t变更数test");
        XWPFTemplate doc = XWPFTemplate.compile("src/test/resources/issue/161.docx");
        doc.render(dataMap).writeToFile("out_issue_161.docx");

    }

}
