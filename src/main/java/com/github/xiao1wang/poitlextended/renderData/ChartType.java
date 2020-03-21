package com.github.xiao1wang.poitlextended.renderData;

/**
 * TODO: 支持的图表类型(其他图形，由于api为提供对应的类)，目前只针对二维做了适配
 */
public enum ChartType {
    /**
     * 柱状图(或条形图)
     */
    BAR,
    /**
     * 折线图
     */
    LINE,
    /**
     * 饼图
     */
    PIE,
    /**
     * 面积图
     */
    AREA,
    /**
     * 环形图
     */
    DOUGHNUT;
}
