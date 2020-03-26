package com.github.xiao1wang.poitlextended.renderpolicy;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import com.github.xiao1wang.poitlextended.renderData.ChartRenderData;
import com.github.xiao1wang.poitlextended.renderData.ChartType;
import com.github.xiao1wang.poitlextended.renderData.ChartTypeData;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAreaChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAreaSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTAxDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBarSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTDoughnutChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTLineSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumDataSource;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTNumVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPieSer;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTSerTx;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrData;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrRef;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTStrVal;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTTitle;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTUnsignedInt;
import org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * TODO: 图表插件
 */
public class ChartRenderPolicy extends AbstractRenderPolicy<ChartRenderData> {

    private final static Logger LOGGER = LoggerFactory.getLogger(ChartRenderPolicy.class);
    private final static String sheetName = "Sheet1";
    private final static Pattern chartIdPattern = Pattern.compile("r:id=\"(.+?)\"");

    @Override
    protected boolean validate(ChartRenderData data) {
        return true;
    }

    @Override
    protected void afterRender(RenderContext<ChartRenderData> context) {
        clearPlaceholder(context, false);
    }

    @Override
    public void doRender(RenderContext<ChartRenderData> context) throws Exception {
        try {
            ChartRenderData chartRenderData = context.getData();
            // 如果当前对象不存在，就将图表清空
            // 动态数据
            String title = chartRenderData == null ? null : chartRenderData.getTitle();
            String[] titleArr = chartRenderData == null ? null : chartRenderData.getColArr();
            List<Object[]> list = chartRenderData == null ? null : chartRenderData.getRowList();
            List<ChartTypeData> chartList = chartRenderData == null ? null : chartRenderData.getChartList();

            /*
            基本思路，就是先编写chart模板，然后替换模板中的图形
            替换分为两步，一个是替换excel中的内容，一个是替换图形的值
             */
            // 得到当前设置填充图表对象所在word的执行位置
            XWPFRun run = context.getRun();
            // 得到图表所在的段落
            XWPFParagraph xwpfParagraph = (XWPFParagraph) run.getParent();
            // 获取图表所在的xml数据，找到图表对应的id值
            String chartId = null;
            String graphicXml = xwpfParagraph.getRuns().get(0).getCTR().getDrawingArray(0).getInlineArray(0).getGraphic().xmlText();
            Matcher matcher = chartIdPattern.matcher(graphicXml);
            if (graphicXml.indexOf("chart") != -1 && matcher.find()) {
                chartId = matcher.group(1);
            }
            if (StringUtils.isNotBlank(chartId)) {
                // 基于文档的对照关系，找到图表对应的对象
                NiceXWPFDocument document = (NiceXWPFDocument) run.getParent().getDocument();
                POIXMLDocumentPart documentPart = document.getRelationById(chartId);
                if (documentPart != null) {
                    XWPFChart chart = (XWPFChart) documentPart;
                    CTChart ctChart = chart.getCTChart();

                    // 如果有标题，就需要设置标题
                    CTTitle ctTitle = ctChart.getTitle();
                    if (title != null && ctTitle.getTx() != null) {
                        CTTextParagraph ctTextParagraph = ctTitle.getTx().getRich().getPArray(0);
                        List<CTRegularTextRun> rList = ctTextParagraph.getRList();
                        if (rList != null && rList.size() > 0) {
                            rList.get(0).setT(title);
                            // 将额外的标题删除掉
                            int length = rList.size();
                            for (int i = 1; i < length; i++) {
                                rList.remove(i);
                            }
                        }
                    }

                    // 设置图表内部的excel数据
                    XSSFWorkbook workbook = this.excelList(list, titleArr);
                    // 由于excel的位置不好找，只能通过与XWPFChart的关联项找
                    List<POIXMLDocumentPart> partList = chart.getRelations();
                    if (partList != null && partList.size() > 0) {
                        for (POIXMLDocumentPart xlsPart : partList) {
                            String contentType = xlsPart.getPackagePart().getContentType();
                            if ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet".equals(contentType)) {
                                OutputStream xlsOut = xlsPart.getPackagePart().getOutputStream();
                                try {
                                    workbook.write(xlsOut);
                                    xlsOut.close();
                                } catch (IOException e) {
                                    e.printStackTrace();
                                } finally {
                                    if (workbook != null) {
                                        try {
                                            workbook.close();
                                        } catch (IOException e) {
                                            e.printStackTrace();
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // 基于传递的参数，得到对应图表的信息
                    CTPlotArea plotArea = ctChart.getPlotArea();
                    if (chartRenderData == null || chartList == null || chartList.size() == 0) {
                        //先清空改plotArea下的所有图表数据
                        clearChart(plotArea);
                    }
                    if (chartList != null && chartList.size() > 0) {
                        for (ChartTypeData chartTypeData : chartList) {
                            ChartType chartType = chartTypeData.getChartType();
                            // 得到每种类型图表的具体数据
                            if (list != null && list.size() > 0 && chartTypeData.getStartPosition() != null && chartTypeData.getEndPosition() != null) {
                                int size = list.get(0).length;
                                if (chartTypeData.getEndPosition() < size) {
                                    // 需要保留第一个图例的样式
                                    XmlObject chartXml = null;
                                    for (int i = chartTypeData.getStartPosition(); i <= chartTypeData.getEndPosition(); i++) {
                                        CTSerTx serTx = null;
                                        CTAxDataSource catDataSource = null;
                                        CTNumDataSource valDataSource = null;
                                        // 创建一个新的系列,并添加该系列的idx，同时得到对应的变量数据
                                        switch (chartType) {
                                            case BAR:
                                                CTBarChart barChart = plotArea.getBarChartArray(0);
                                                if (chartXml == null) {
                                                    chartXml = barChart.getSerArray(0).copy();
                                                    barChart.setSerArray(null);
                                                }
                                                CTBarSer ctBarSer = barChart.addNewSer();
                                                ctBarSer.set(chartXml);
                                                CTUnsignedInt barIdx = ctBarSer.getIdx();
                                                barIdx.setVal(i - 1);
                                                ctBarSer.setIdx(barIdx);
                                                serTx = ctBarSer.getTx();
                                                catDataSource = ctBarSer.getCat();
                                                valDataSource = ctBarSer.getVal();
                                                break;
                                            case LINE:
                                                CTLineChart lineChart = plotArea.getLineChartArray(0);
                                                if (chartXml == null) {
                                                    chartXml = lineChart.getSerArray(0).copy();
                                                    lineChart.setSerArray(null);
                                                }
                                                CTLineSer ctLineSer = lineChart.addNewSer();
                                                ctLineSer.set(chartXml);
                                                CTUnsignedInt lineIdx = ctLineSer.getIdx();
                                                lineIdx.setVal(i - 1);
                                                ctLineSer.setIdx(lineIdx);
                                                serTx = ctLineSer.getTx();
                                                catDataSource = ctLineSer.getCat();
                                                valDataSource = ctLineSer.getVal();
                                                break;
                                            case PIE:
                                                CTPieChart pieChart = plotArea.getPieChartArray(0);
                                                if (chartXml == null) {
                                                    chartXml = pieChart.getSerArray(0).copy();
                                                    pieChart.setSerArray(null);
                                                }
                                                CTPieSer ctPieSer = pieChart.addNewSer();
                                                ctPieSer.set(chartXml);
                                                CTUnsignedInt pieIdx = ctPieSer.getIdx();
                                                pieIdx.setVal(i - 1);
                                                ctPieSer.setIdx(pieIdx);
                                                serTx = ctPieSer.getTx();
                                                catDataSource = ctPieSer.getCat();
                                                valDataSource = ctPieSer.getVal();
                                                break;
                                            case AREA:
                                                CTAreaChart areaChart = plotArea.getAreaChartArray(0);
                                                if (chartXml == null) {
                                                    chartXml = areaChart.getSerArray(0).copy();
                                                    areaChart.setSerArray(null);
                                                }
                                                CTAreaSer ctAreaSer = areaChart.addNewSer();
                                                ctAreaSer.set(chartXml);
                                                CTUnsignedInt areaIdx = ctAreaSer.getIdx();
                                                areaIdx.setVal(i - 1);
                                                ctAreaSer.setIdx(areaIdx);
                                                serTx = ctAreaSer.getTx();
                                                catDataSource = ctAreaSer.getCat();
                                                valDataSource = ctAreaSer.getVal();
                                                break;
                                            case DOUGHNUT:
                                                CTDoughnutChart doughnutChart = plotArea.getDoughnutChartArray(0);
                                                if (chartXml == null) {
                                                    chartXml = doughnutChart.getSerArray(0).copy();
                                                    doughnutChart.setSerArray(null);
                                                }
                                                CTPieSer ctDoughnutSer = doughnutChart.addNewSer();
                                                ctDoughnutSer.set(chartXml);
                                                CTUnsignedInt doughnutIdx = ctDoughnutSer.getIdx();
                                                doughnutIdx.setVal(i - 1);
                                                ctDoughnutSer.setIdx(doughnutIdx);
                                                serTx = ctDoughnutSer.getTx();
                                                catDataSource = ctDoughnutSer.getCat();
                                                valDataSource = ctDoughnutSer.getVal();
                                                break;
                                        }
                                        // 添加该系列的名称
                                        if (serTx != null) {
                                            CTStrRef txStrRef = serTx.getStrRef();
                                            if (titleArr != null && i < titleArr.length) {
                                                String txRange = new CellRangeAddress(0, 0, i, i).formatAsString(sheetName, true);
                                                txStrRef.setF(txRange);
                                                CTStrData txData = txStrRef.getStrCache();
                                                CTUnsignedInt txPtCount = txData.getPtCount();
                                                txPtCount.setVal(1);
                                                txData.setPtArray(null);
                                                CTStrVal txPt = txData.addNewPt();
                                                txPt.setV(titleArr[i]);
                                                txPt.setIdx(0);
                                            } else {
                                                serTx.setStrRef(null);
                                            }
                                        }

                                        int firstRow = titleArr != null ? 1 : 0;
                                        int lastRow = titleArr != null ? list.size() : list.size() - 1;
                                        // 添加该系列下的种类对应x轴的数据范围
                                        if (catDataSource != null) {
                                            CTStrRef catStrRef = catDataSource.getStrRef();
                                            CTNumRef catNumRef = catDataSource.getNumRef();
                                            if (catNumRef != null) {
                                                catDataSource.set(null);
                                                //catNumRef.set(null);
                                                catStrRef = catDataSource.addNewStrRef();
                                                this.createStrRef(catStrRef, firstRow, lastRow, list, 0);

                                                /*catNumRef = catDataSource.getNumRef();
                                                this.setNumRef(catNumRef,firstRow,lastRow,list,0);*/
                                            } else if (catStrRef != null) {
                                                catStrRef = catDataSource.getStrRef();
                                                this.setStrRef(catStrRef, firstRow, lastRow, list, 0);
                                            }
                                        }

                                        // 添加该系列下的种类对应y轴的数据范围
                                        if (valDataSource != null) {
                                            CTNumRef valNumRef = valDataSource.getNumRef();
                                            this.setNumRef(valNumRef, firstRow, lastRow, list, i);
                                        }
                                    }
                                }
                            } else {
                                clearChart(plotArea);
                            }
                        }
                    }
                }
            } else {
                LOGGER.info("未找到序号为{}的图表，请核对模板中的序号", chartId);
            }
        } catch (Exception e) {
            e.printStackTrace();
            LOGGER.info("填充的数据格式不符合要求，异常信息为{}", e.getMessage());
        }
    }

    /**
     * 生成图表中格式数据为字符串的
     *
     * @param catStrRef
     * @param firstRow
     * @param lastRow
     * @param list
     */
    private void createStrRef(CTStrRef catStrRef, int firstRow, int lastRow, List<Object[]> list, int index) {
        String catRange = new CellRangeAddress(firstRow, lastRow, index, index).formatAsString(sheetName, true);
        catStrRef.setF(catRange);
        CTStrData catStrData = catStrRef.addNewStrCache();
        CTUnsignedInt catPtCount = catStrData.addNewPtCount();
        catPtCount.setVal(list.size());
        // 开始设置该系列下种类对应的名称
        catStrData.setPtArray(null);
        for (int j = 0; j < list.size(); j++) {
            Object[] dataArr = list.get(j);
            CTStrVal ctStrVal = catStrData.addNewPt();
            ctStrVal.setIdx(j);
            ctStrVal.setV(dataArr[index].toString());
        }
    }

    /**
     * 设置图表中格式数据为字符串的
     *
     * @param catStrRef
     * @param firstRow
     * @param lastRow
     * @param list
     */
    private void setStrRef(CTStrRef catStrRef, int firstRow, int lastRow, List<Object[]> list, int index) {
        String catRange = new CellRangeAddress(firstRow, lastRow, index, index).formatAsString(sheetName, true);
        catStrRef.setF(catRange);
        CTStrData catStrData = catStrRef.getStrCache();
        CTUnsignedInt catPtCount = catStrData.getPtCount();
        catPtCount.setVal(list.size());
        // 开始设置该系列下种类对应的名称
        catStrData.setPtArray(null);
        for (int j = 0; j < list.size(); j++) {
            Object[] dataArr = list.get(j);
            CTStrVal ctStrVal = catStrData.addNewPt();
            ctStrVal.setIdx(j);
            ctStrVal.setV(dataArr[index].toString());
        }
    }

    /**
     * 设置图表中格式数据为数值的
     *
     * @param valNumRef
     * @param firstRow
     * @param lastRow
     * @param list
     */
    private void setNumRef(CTNumRef valNumRef, int firstRow, int lastRow, List<Object[]> list, int index) {
        String valRange = new CellRangeAddress(firstRow, lastRow, index, index).formatAsString(sheetName, true);
        valNumRef.setF(valRange);
        CTNumData valNumData = valNumRef.getNumCache();
        CTUnsignedInt valPtCount = valNumData.getPtCount();
        valPtCount.setVal(list.size());
        // 开始设置该系列下种类对应的值
        valNumData.setPtArray(null);
        for (int j = 0; j < list.size(); j++) {
            Object[] dataArr = list.get(j);
            CTNumVal ctNumVal = valNumData.addNewPt();
            ctNumVal.setIdx(j);
            ctNumVal.setV(dataArr[index].toString());
        }
    }

    /**
     * 清空plotArea下的图形数据
     * 由于没有统一的清楚方法，只能一一删除
     *
     * @param plotArea
     */
    private void clearChart(CTPlotArea plotArea) {
        List<CTBarChart> barList = plotArea.getBarChartList();
        if (barList != null && barList.size() > 0) {
            for (CTBarChart chart : barList) {
                chart.setSerArray(null);
            }
        }
        List<CTLineChart> lineList = plotArea.getLineChartList();
        if (lineList != null && lineList.size() > 0) {
            for (CTLineChart chart : lineList) {
                chart.setSerArray(null);
            }
        }
        List<CTPieChart> pieList = plotArea.getPieChartList();
        if (pieList != null && pieList.size() > 0) {
            for (CTPieChart chart : pieList) {
                chart.setSerArray(null);
            }
        }
        List<CTAreaChart> areaList = plotArea.getAreaChartList();
        if (areaList != null && areaList.size() > 0) {
            for (CTAreaChart chart : areaList) {
                chart.setSerArray(null);
            }
        }
        List<CTDoughnutChart> doughuntList = plotArea.getDoughnutChartList();
        if (doughuntList != null && doughuntList.size() > 0) {
            for (CTDoughnutChart chart : doughuntList) {
                chart.setSerArray(null);
            }
        }
    }

    /**
     * 生成excel数据
     *
     * @param list
     * @param title
     * @return XSSFWorkbook
     */
    public static XSSFWorkbook excelList(List<Object[]> list, String[] title) {
        //工作区
        XSSFWorkbook wb = new XSSFWorkbook();
        //创建第一个sheet
        XSSFSheet sheet = wb.createSheet(sheetName);
        int startRow = 0;
        //生成第一行
        XSSFRow row = null;
        if (title != null && title.length > 0) {
            row = sheet.createRow(startRow++);
            for (int i = 0; i < title.length; i++) {
                row.createCell(i).setCellValue(title[i]);
            }
        }
        if (list != null && list.size() > 0) {
            int size = list.get(0).length;
            for (int i = 0; i < list.size(); i++) {
                row = sheet.createRow(startRow++);
                for (int j = 0; j < size; j++) {
                    String val = list.get(i)[j] == null ? "" : list.get(i)[j].toString();
                    // 需要优化，针对一些数字和字符串的区别
                    try {
                        BigDecimal bigVal = new BigDecimal(val);
                        row.createCell(j).setCellValue(bigVal.doubleValue());
                    } catch (Exception e) {
                        row.createCell(j).setCellValue(val);
                    }
                }
            }
        }
        return wb;
    }
}
