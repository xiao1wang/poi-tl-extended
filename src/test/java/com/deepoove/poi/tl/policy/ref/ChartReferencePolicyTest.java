package com.deepoove.poi.tl.policy.ref;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.HashMap;

@DisplayName("Chart ReferencePolicy test case")
public class ChartReferencePolicyTest {

    @Test
    public void testBarChart() throws Exception {

        Configure configure = Configure.newBuilder().referencePolicy(new MyChartReferenceRenderPolicy()).build();

        XWPFTemplate template = XWPFTemplate.compile("src/test/resources/template/reference_chart.docx", configure)
                .render(new HashMap<>());

        template.writeToFile("out_reference_chart.docx");
    }

}
