package com.github.xiao1wang.poitlextended.renderData;

/**
 * 图表数据值范围
 */
public class ChartTypeData {

    // 图表类型
    private ChartType chartType;
    // excel中属于该图表类型数据的起始位置，依ChartRenderData类中的属性colArr为依据，长度从0开始
    private Integer startPosition = null;
    // excel中属于该图表类型数据的结束位置，依ChartRenderData类中的属性colArr为依据
    private Integer endPosition = null;

    public ChartTypeData(ChartType chartType, Integer startPosition, Integer endPosition) {
        this.chartType = chartType;
        this.startPosition = startPosition;
        this.endPosition = endPosition;
    }

    public ChartType getChartType() {
        return chartType;
    }

    public void setChartType(ChartType chartType) {
        this.chartType = chartType;
    }

    public Integer getStartPosition() {
        return startPosition;
    }

    public void setStartPosition(Integer startPosition) {
        this.startPosition = startPosition;
    }

    public Integer getEndPosition() {
        return endPosition;
    }

    public void setEndPosition(Integer endPosition) {
        this.endPosition = endPosition;
    }
}
