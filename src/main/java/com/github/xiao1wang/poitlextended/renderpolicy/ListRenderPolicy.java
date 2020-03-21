package com.github.xiao1wang.poitlextended.renderpolicy;

import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

/**
 * TODO: 更新列表
 */
public class ListRenderPolicy extends AbstractRenderPolicy<List<String>> {

    private static Logger LOGGER = LoggerFactory.getLogger(ListRenderPolicy.class);

    @Override
    protected boolean validate(List<String> data) {
        return true;
    }

    @Override
    protected void afterRender(RenderContext<List<String>> context) {
        //clearPlaceholder(context, true);
    }


    @Override
    public void doRender(RenderContext<List<String>> context) throws Exception {
        List<String> list = context.getData();
        XWPFRun run = context.getRun();
        NiceXWPFDocument document = (NiceXWPFDocument) run.getParent().getDocument();
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
        try {
            XWPFParagraph currentParagraph = (XWPFParagraph) run.getParent();
            CTP ctp = currentParagraph.getCTP();
            XmlObject xmlObject = ctp.copy();
            if (list == null) {
                ctp.set(null);
                return;
            }
            if (null != list) {
                for (int i = 0; i < list.size(); i++) {
                    XWPFParagraph paragraph = null;
                    if (i != list.size() - 1) {
                        paragraph = bodyContainer.insertNewParagraph(run);
                    } else {
                        paragraph = currentParagraph;
                    }
                    CTP newCtp = paragraph.getCTP();
                    newCtp.set(xmlObject);
                    List<CTR> rList = newCtp.getRList();
                    if (rList != null && rList.size() > 0) {
                        rList.get(0).getTArray(0).setStringValue(list.get(i));
                        for (int j = 1; j < rList.size(); j++) {
                            rList.remove(j);
                        }
                    }
                }
            }
        } catch (Exception e) {
            throw new RenderException(" 列表解析失败:" + e.getMessage(), e);
        }
    }
}
