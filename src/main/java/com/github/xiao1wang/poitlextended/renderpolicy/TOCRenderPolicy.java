package com.github.xiao1wang.poitlextended.renderpolicy;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import com.github.xiao1wang.poitlextended.renderData.TOCRenderData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtBlock;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * TODO: 目录插件
 */
public class TOCRenderPolicy extends AbstractRenderPolicy<TOCRenderData> {

    private static Logger LOGGER = LoggerFactory.getLogger(TOCRenderPolicy.class);

    @Override
    protected boolean validate(TOCRenderData data) {
        return true;
    }

    @Override
    protected void afterRender(RenderContext<TOCRenderData> context) {
        clearPlaceholder(context, true);
    }

    @Override
    public void doRender(RenderContext<TOCRenderData> context) throws Exception {
        try {
            TOCRenderData tocRenderData = context.getData();
            // 基本思路，先在document对应的位置生成一个展位区，将占位区的对象传递出去，供文档其余数据生成完后，在设置目录结构
            XWPFRun run = context.getRun();
            NiceXWPFDocument document = (NiceXWPFDocument) run.getParent().getDocument();
            CTSdtBlock block = document.getDocument().getBody().insertNewSdt(0);
            tocRenderData.setBlock(block);
            //tocRenderData.setTOCTitle();
            tocRenderData.setDocument(document);
            tocRenderData.setRun(run);
            tocRenderData.setTOCTitleFirst();
        } catch (Exception e) {
            e.printStackTrace();
            LOGGER.info("填充的数据格式不符合要求，异常信息为{}", e.getMessage());
        }
    }

}
