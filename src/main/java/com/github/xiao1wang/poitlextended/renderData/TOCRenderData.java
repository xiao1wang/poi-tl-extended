package com.github.xiao1wang.poitlextended.renderData;

import com.deepoove.poi.data.RenderData;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.util.LocaleUtil;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTEmpty;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHpsMeasure;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtBlock;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtContentBlock;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtEndPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSdtPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabs;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabTlc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTheme;

import java.math.BigInteger;
import java.util.List;

/**
 * TODO: 目录数据
 */
public class TOCRenderData implements RenderData {

    // 目录下的标题基本到几层结构
    private int maxLevel = 0;
    // 目录标题
    private String title = null;
    // 域所在的位置块
    private CTSdtBlock block;
    // 当前操作的文档
    private NiceXWPFDocument document = null;
    // 当前操作的位置
    private XWPFRun run = null;

    public TOCRenderData(String title, int maxLevel) {
        this.title = title;
        this.maxLevel = maxLevel;
    }

    public TOCRenderData(int maxLevel, String title) {
        this.maxLevel = maxLevel;
        this.title = title;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public CTSdtBlock getBlock() {
        return block;
    }

    public void setBlock(CTSdtBlock block) {
        this.block = block;
    }

    public int getMaxLevel() {
        return maxLevel;
    }

    public void setMaxLevel(int maxLevel) {
        this.maxLevel = maxLevel;
    }

    public NiceXWPFDocument getDocument() {
        return document;
    }

    public void setDocument(NiceXWPFDocument document) {
        this.document = document;
    }

    public XWPFRun getRun() {
        return run;
    }

    public void setRun(XWPFRun run) {
        this.run = run;
    }

    /**
     * 设置目录的标题
     */
    public void setTOCTitle() {
        CTSdtPr sdtPr = block.addNewSdtPr();
        CTDecimalNumber id = sdtPr.addNewId();
        id.setVal(new BigInteger("4844945"));
        sdtPr.addNewDocPartObj().addNewDocPartGallery().setVal("Table of contents");
        CTSdtEndPr sdtEndPr = block.addNewSdtEndPr();
        CTRPr rPr = sdtEndPr.addNewRPr();
        CTFonts fonts = rPr.addNewRFonts();
        fonts.setAsciiTheme(STTheme.MINOR_H_ANSI);
        fonts.setEastAsiaTheme(STTheme.MINOR_H_ANSI);
        fonts.setHAnsiTheme(STTheme.MINOR_H_ANSI);
        fonts.setCstheme(STTheme.MINOR_BIDI);
        rPr.addNewB().setVal(STOnOff.OFF);
        rPr.addNewBCs().setVal(STOnOff.OFF);
        rPr.addNewColor().setVal("auto");
        rPr.addNewSz().setVal(new BigInteger("24"));
        rPr.addNewSzCs().setVal(new BigInteger("24"));
        CTSdtContentBlock content = block.addNewSdtContent();
        CTP p = content.addNewP();
        p.setRsidR("00EF7E24".getBytes(LocaleUtil.CHARSET_1252));
        p.setRsidRDefault("00EF7E24".getBytes(LocaleUtil.CHARSET_1252));
        p.addNewPPr().addNewPStyle().setVal("TOCHeading");
        p.addNewR().addNewT().setStringValue(title);
        //设置段落对齐方式，即将“目录”二字居中
        CTPPr pr = p.getPPr();
        CTJc jc = pr.isSetJc() ? pr.getJc() : pr.addNewJc();
        STJc.Enum en = STJc.Enum.forInt(ParagraphAlignment.CENTER.getValue());
        jc.setVal(en);
        //"目录"二字的字体
        CTRPr pRpr = p.getRArray(0).addNewRPr();
        fonts = pRpr.isSetRFonts() ? pRpr.getRFonts() : pRpr.addNewRFonts();
        fonts.setAscii("Times New Roman");
        fonts.setEastAsia("华文中宋");
        fonts.setHAnsi("华文中宋");
        //"目录"二字加粗
        CTOnOff bold = pRpr.isSetB() ? pRpr.getB() : pRpr.addNewB();
        bold.setVal(STOnOff.TRUE);
        // 设置“目录”二字字体大小为24号
        CTHpsMeasure sz = pRpr.isSetSz() ? pRpr.getSz() : pRpr.addNewSz();
        sz.setVal(new BigInteger("36"));
    }

    /**
     * 设置目录的标题
     */
    public void setTOCTitleFirst() {
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
        XWPFParagraph paragraph = bodyContainer.insertNewParagraph(run);
        CTP p = paragraph.getCTP();
        p.setRsidR("00EF7E24".getBytes(LocaleUtil.CHARSET_1252));
        p.setRsidRDefault("00EF7E24".getBytes(LocaleUtil.CHARSET_1252));
        p.addNewPPr().addNewPStyle().setVal("TOCHeading");
        p.addNewR().addNewT().setStringValue(title);
        //设置段落对齐方式，即将“目录”二字居中
        CTPPr pr = p.getPPr();
        CTJc jc = pr.isSetJc() ? pr.getJc() : pr.addNewJc();
        STJc.Enum en = STJc.Enum.forInt(ParagraphAlignment.CENTER.getValue());
        jc.setVal(en);
        //"目录"二字的字体
        CTRPr pRpr = p.getRArray(0).addNewRPr();
        //"目录"二字加粗
        CTOnOff bold = pRpr.isSetB() ? pRpr.getB() : pRpr.addNewB();
        bold.setVal(STOnOff.TRUE);
        // 设置“目录”二字字体大小为24号
        CTHpsMeasure sz = pRpr.isSetSz() ? pRpr.getSz() : pRpr.addNewSz();
        sz.setVal(new BigInteger("36"));
    }

    public void setItem2TOC(NiceXWPFDocument doc) {
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
        List<IBodyElement> bodyElementList = doc.getBodyElements();
        if (bodyElementList != null && bodyElementList.size() > 0) {
            int index = 0;

            for (IBodyElement element : bodyElementList) {
                if (element instanceof XWPFParagraph) {
                    XWPFParagraph par = (XWPFParagraph) element;
                    String parStyle = par.getStyle();
                    if (parStyle != null && StringUtils.isNumeric(parStyle)) {
                        List<CTBookmark> bookmarkList = par.getCTP().getBookmarkStartList();
                        try {
                            int level = Integer.parseInt(parStyle);
                            if (level <= maxLevel) {
                                String title = par.getText();
                                String bookmarkRef = null;
                                if (bookmarkList == null || bookmarkList.size() == 0) {
                                    bookmarkRef = "_Toc112723803" + (index++);
                                } else {
                                    bookmarkRef = bookmarkList.get(bookmarkList.size() - 1).getName();
                                }

                                XWPFParagraph paragraph = bodyContainer.insertNewParagraph(run);
                                CTP p = paragraph.getCTP();
                                p.setRsidR("00EF7E24".getBytes(LocaleUtil.CHARSET_1252));
                                p.setRsidRDefault("00EF7E24".getBytes(LocaleUtil.CHARSET_1252));
                                CTPPr pPr = p.addNewPPr();
                                pPr.addNewPStyle().setVal((level * 10) + "");
                                CTTabs tabs = pPr.addNewTabs();
                                CTTabStop tab = tabs.addNewTab();
                                tab.setVal(STTabJc.RIGHT);
                                tab.setLeader(STTabTlc.DOT);
                                tab.setPos(new BigInteger("8290"));
                                pPr.addNewRPr().addNewNoProof();
                                CTR run = p.addNewR();
                                run.addNewRPr().addNewNoProof();
                                run.addNewT().setStringValue(title);
                                run = p.addNewR();
                                run.addNewRPr().addNewNoProof();
                                run.addNewTab();
                                run = p.addNewR();
                                run.addNewRPr().addNewNoProof();
                                run.addNewFldChar().setFldCharType(STFldCharType.BEGIN);
                                run = p.addNewR();
                                run.addNewRPr().addNewNoProof();
                                CTText text = run.addNewInstrText();
                                text.setSpace(SpaceAttribute.Space.PRESERVE);
                                text.setStringValue(" PAGEREF " + bookmarkRef + " \\h ");
                                p.addNewR().addNewRPr().addNewNoProof();
                                run = p.addNewR();
                                run.addNewRPr().addNewNoProof();
                                run.addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
                                run = p.addNewR();
                                run.addNewRPr().addNewNoProof();

                                // 获取当前标题名称，是在文档的第几页出现的
                                int page = pageIndex(par, title, bodyElementList);

                                run.addNewT().setStringValue(Integer.toString(page));
                                run = p.addNewR();
                                run.addNewRPr().addNewNoProof();
                                run.addNewFldChar().setFldCharType(STFldCharType.END);
                            }
                        } catch (NumberFormatException e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
        }
    }

    public int pageIndex(XWPFParagraph currentPar, String headTitle, List<IBodyElement> bodyElementList) {
        int page = 1;
        boolean findFlag = false;
        if (bodyElementList != null && bodyElementList.size() > 0) {
            for (int i = 0; i < bodyElementList.size() && !findFlag; i++) {
                // 分页符的位置可能出现在段落，或者表格下
                IBodyElement element = bodyElementList.get(i);
                if (element instanceof XWPFParagraph) {
                    XWPFParagraph par = (XWPFParagraph) element;
                    String title = par.getText();
                    List<CTR> ctrlist = par.getCTP().getRList();//获取<w:p>标签下的<w:r>list
                    for (int j = 0; j < ctrlist.size(); j++) {  //遍历r
                        CTR r = ctrlist.get(j);
                        List<CTEmpty> breaklist = r.getLastRenderedPageBreakList();//判断是否存在此标签
                        if (breaklist.size() > 0) {
                            page++; //页数添加
                        }
                        if (headTitle.equals(title) && currentPar == par) {
                            findFlag = true;
                            break;
                        }
                    }
                } else if (element instanceof XWPFTable) {
                    XWPFTable table = (XWPFTable) element;
                    List<XWPFTableRow> tableList = table.getRows();
                    if (tableList != null && tableList.size() > 0) {
                        for (XWPFTableRow row : tableList) {
                            List<XWPFTableCell> cellList = row.getTableCells();
                            if (cellList != null && cellList.size() > 0) {
                                for (XWPFTableCell cell : cellList) {
                                    List<XWPFParagraph> paragraphList = cell.getParagraphs();
                                    if (paragraphList != null && paragraphList.size() > 0) {
                                        for (XWPFParagraph par : paragraphList) {
                                            List<CTR> ctrlist = par.getCTP().getRList();//获取<w:p>标签下的<w:r>list
                                            for (int j = 0; j < ctrlist.size(); j++) {  //遍历r
                                                CTR r = ctrlist.get(j);
                                                List<CTEmpty> breaklist = r.getLastRenderedPageBreakList();//判断是否存在此标签
                                                if (breaklist.size() > 0) {
                                                    page++; //页数添加
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
        }
        if (!findFlag) {
            page = 1;
        }
        return page;
    }
}
