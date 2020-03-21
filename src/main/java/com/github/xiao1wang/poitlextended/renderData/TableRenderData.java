package com.github.xiao1wang.poitlextended.renderData;

import com.deepoove.poi.data.RenderData;

import java.util.List;

/**
 * TODO: 表格渲染数据
 */
public class TableRenderData implements RenderData {

    // 数据行从第几行开始，数字从0开始
    private int start = 0;
    // 具体每行数据的内容
    private List<Object[]> rowList = null;

    public TableRenderData(int start, List<Object[]> rowList) {
        this.start = start;
        this.rowList = rowList;
    }

    public int getStart() {
        return start;
    }

    public void setStart(int start) {
        this.start = start;
    }

    public List<Object[]> getRowList() {
        return rowList;
    }

    public void setRowList(List<Object[]> rowList) {
        this.rowList = rowList;
    }
}
