package com.github.xiao1wang.poitlextended.renderData;

import com.deepoove.poi.data.RenderData;

import java.util.List;

/**
 * 图表结构
 */
public class ChartRenderData implements RenderData {

    // 图表上的标题
    private String title = null;
    // excel中中文字段名称
    private String[] colArr = null;
    // excel中每行具体数据
    private List<Object[]> rowList = null;
    // 该图中包含的图表类型(有些图表是复合图)
    private List<ChartTypeData> chartList = null;

    public ChartRenderData(String[] colArr, List<Object[]> rowList, List<ChartTypeData> chartList) {
        this.colArr = colArr;
        this.rowList = rowList;
        this.chartList = chartList;
    }

    public ChartRenderData(String title, String[] colArr, List<Object[]> rowList, List<ChartTypeData> chartList) {
        this.title = title;
        this.colArr = colArr;
        this.rowList = rowList;
        this.chartList = chartList;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String[] getColArr() {
        return colArr;
    }

    public void setColArr(String[] colArr) {
        this.colArr = colArr;
    }

    public List<Object[]> getRowList() {
        return rowList;
    }

    public void setRowList(List<Object[]> rowList) {
        this.rowList = rowList;
    }

    public List<ChartTypeData> getChartList() {
        return chartList;
    }

    public void setChartList(List<ChartTypeData> chartList) {
        this.chartList = chartList;
    }
}
