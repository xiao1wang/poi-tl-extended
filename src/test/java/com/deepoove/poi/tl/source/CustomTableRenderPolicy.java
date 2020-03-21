package com.deepoove.poi.tl.source;

import com.deepoove.poi.data.MiniTableRenderData;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 * 通过
 * <code>
 * Configure.newBuilder().bind("report", new
 * CustomTableRenderPolicy());
 * </code>
 * 将模板report的策略设置成自定义的表格策略
 *
 * @author Sayi
 */
public class CustomTableRenderPolicy extends AbstractRenderPolicy<Object> {

    @Override
    protected void afterRender(RenderContext<Object> context) {
        // 清空模板标签
        clearPlaceholder(context, true);
    }

    @Override
    public void doRender(RenderContext<Object> context) throws Exception {

        XWPFRun run = context.getRun();
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
        // 定义行列
        int row = 10, col = 8;
        // 插入表格
        XWPFTable table = bodyContainer.insertNewTable(run, row, col);

        // 定义表格宽度、边框和样式
        TableTools.widthTable(table, MiniTableRenderData.WIDTH_A4_FULL, col);
        TableTools.borderTable(table, 4);

        // TODO 调用XWPFTable API操作表格：data对象可以包含任意你想要的数据，包括图片文本等
        // TODO 调用MiniTableRenderPolicy.Helper.renderRow方法快速方便的渲染一行数据
        // TODO 调用TableTools类方法操作表格，比如合并单元格
        // ......
        TableTools.mergeCellsHorizonal(table, 0, 0, 7);
        TableTools.mergeCellsVertically(table, 0, 1, 9);

    }

}
