package com.github.xiao1wang.poitlextended.renderpolicy;

import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import com.github.xiao1wang.poitlextended.renderData.TableRenderData;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * TODO : 更新表格策略
 */
public class TableRenderPolicy extends AbstractRenderPolicy<TableRenderData> {

    private static Logger LOGGER = LoggerFactory.getLogger(TableRenderPolicy.class);

    @Override
    protected boolean validate(TableRenderData data) {
        return true;
    }

    @Override
    protected void afterRender(RenderContext<TableRenderData> context) {
        clearPlaceholder(context, true);
    }


    @Override
    public void doRender(RenderContext<TableRenderData> context) throws Exception {
        TableRenderData data = context.getData();
        XWPFRun run = context.getRun();
        NiceXWPFDocument document = (NiceXWPFDocument) run.getParent().getDocument();
        try {
            if (!TableTools.isInsideTable(run)) {
                throw new IllegalStateException(" 变量的定义必须放在表格的单元格内 ");
            }
            XWPFTableCell currentCell = (XWPFTableCell) ((XWPFParagraph) run.getParent()).getBody();
            XWPFTable table = currentCell.getTableRow().getTable();
            if (data == null) {
                return;
            }
            List<Object[]> list = data.getRowList();
            if (null != list) {
                // 存放之前段落的样式
                List<XmlObject> pList = new ArrayList<>();
                List<XWPFTableRow> oldRowList = table.getRows();
                if (oldRowList != null && oldRowList.size() > data.getStart()) {
                    List<XWPFTableCell> cellList = table.getRow(data.getStart()).getTableCells();
                    if (cellList != null && cellList.size() > 0) {
                        for (XWPFTableCell cell : cellList) {
                            CTTc ctTc = cell.getCTTc();
                            CTP ctp = (ctTc.sizeOfPArray() == 0) ? ctTc.addNewP() : ctTc.getPArray(0);
                            XmlObject xmlObject = ctp.copy();
                            pList.add(xmlObject);
                        }
                    }
                    // 删除之前的数据
                    while (table.removeRow(data.getStart())) ;
                }
                for (int i = 0; i < list.size(); i++) {
                    XWPFTableRow insertNewTableRow = table.createRow();
                    Object[] arr = list.get(i);
                    for (int k = 0; k < arr.length; k++) {
                        XWPFTableCell cell = insertNewTableRow.getCell(k);

                        // 处理单元格数据
                        CTTc ctTc = cell.getCTTc();
                        CTP ctP = (ctTc.sizeOfPArray() == 0) ? ctTc.addNewP() : ctTc.getPArray(0);
                        ctP.set(pList.get(k));
                        List<CTR> rList = ctP.getRList();
                        if (rList != null && rList.size() > 0) {
                            rList.get(0).getTArray(0).setStringValue(arr[k].toString());
                            for (int j = 1; j < rList.size(); j++) {
                                rList.remove(j);
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            throw new RenderException("dynamic table error:" + e.getMessage(), e);
        }
    }
}
