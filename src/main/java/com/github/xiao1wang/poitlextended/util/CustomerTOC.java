package com.github.xiao1wang.poitlextended.util;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.util.LocaleUtil;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTEmpty;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTParaRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabStop;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTabs;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTabTlc;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

/**
 * 重写目录生成
 *
 * @Author : wangyahui
 * @Date: 2020-03-27 16:48
 */
public class CustomerTOC {

	private static final int titleRange = 10;

	/**
	 * 自动产生目录，缺陷是，对应的页数有时候有误，产生有误的原因是传递的word中分页符不规范
	 *
	 * @param maxLevel
	 * @param contentFlag
	 * @param document
	 */
	public static void automaticGenerateTOC(int maxLevel, String contentFlag, XWPFDocument document) {
		XWPFParagraph p = null;
		boolean flag = false;
		for (int i = 0; i < document.getParagraphs().size(); i++) {
			p = document.getParagraphs().get(i);
			String text = p.getParagraphText();
			if (text != null && text.toLowerCase().contains(contentFlag.toLowerCase())) {
				flag = true;
				break;
			}
		}
		if (flag) {
			List<XWPFParagraph> list = new ArrayList<>();
			Iterator var3 = document.getParagraphs().iterator();
			List<IBodyElement> bodyElementList = document.getBodyElements();
			while (var3.hasNext()) {
				XWPFParagraph paragraph = (XWPFParagraph) var3.next();
				String parStyle = paragraph.getStyle();
				if (parStyle != null && StringUtils.isNumeric(parStyle)) {
					int level = Integer.parseInt(parStyle);
					if (level <= maxLevel) {
						list.add(paragraph);
					}
				}
			}
			if (list.size() > 0) {
				XWPFParagraph currentP = p;
				int[] chapterArr = new int[maxLevel];
				Arrays.fill(chapterArr, 0);
				for (int j = 0; j < list.size(); j++) {
					XWPFParagraph paragraph = list.get(j);
					String parStyle = paragraph.getStyle();
					if (parStyle != null && StringUtils.isNumeric(parStyle)) {
						int level = Integer.parseInt(parStyle);
						for (int k = level; k < maxLevel; k++) {
							chapterArr[k] = 0;
						}
						chapterArr[level - 1] = chapterArr[level - 1] + 1;
						// 得到数组对应的字符串值，同时去掉有.0的内容
						String s = Arrays.toString(chapterArr);
						String chapterArrStr = s.substring(1, s.length() - 1).replaceAll(" ", "").replaceAll(",", ".");
						while (chapterArrStr.endsWith(".0")) {
							chapterArrStr = chapterArrStr.substring(0, chapterArrStr.lastIndexOf(".0"));
						}

						List<CTBookmark> bookmarkList = paragraph.getCTP().getBookmarkStartList();
						String title = paragraph.getText();
						int pageBreakNum = pageBreak(paragraph, title, bodyElementList);
						// 前面多少页不算
						pageBreakNum = pageBreakNum - 2;
						XmlCursor cursor = currentP.getCTP().newCursor();
						XWPFParagraph newPara = document.insertNewParagraph(cursor);
						addRow(maxLevel, chapterArrStr, j, newPara, level, paragraph.getText(), pageBreakNum, bookmarkList.get(0).getName());
					}
				}
			}
			// 清空当前段落的内容
			List<XWPFRun> runList = p.getRuns();
			for (int j = 0; j < runList.size(); j++) {
				XWPFRun r = runList.get(j);
				r.setText("", 0);
			}
			//p.getCTP().set(null);
			/*int index = document.getPosOfParagraph(p);
			document.removeBodyElement(index);*/
		}
	}

	public static void addRow(int maxLevel, String chapterNum, int index, XWPFParagraph newPara, int level, String title, int page, String bookmarkRef) {
		CTP p = newPara.getCTP();
		// 设置标题tab等级
		p.setRsidRDefault("00EF7E24".getBytes(LocaleUtil.CHARSET_1252));
		CTPPr pPr = p.addNewPPr();
		pPr.addNewPStyle().setVal((titleRange * level) + "");
		CTTabs tabs = pPr.addNewTabs();
		CTTabStop tab = tabs.addNewTab();
		tab.setVal(STTabJc.RIGHT);
		tab.setLeader(STTabTlc.DOT);
		tab.setPos(new BigInteger("8296"));
		CTParaRPr paraRPr = pPr.addNewRPr();
		paraRPr.addNewNoProof();

		// 如果是第一项，需要设置目录结构，有时候不起作用
		/*if(index == 0) {
			p.addNewR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);
			CTR run = p.addNewR();
			run.addNewRPr().addNewNoProof();
			CTText text = run.addNewInstrText();
			text.setSpace(SpaceAttribute.Space.PRESERVE);
			text.setStringValue("TOC \\o \"1-"+maxLevel+"\" \\h \\z \\u");
			run = p.addNewR();
			run.addNewRPr().addNewNoProof();
			run.addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
		}*/

		// 添加链接导航
		CTHyperlink hyperlink = p.addNewHyperlink();
		hyperlink.setAnchor(bookmarkRef);
		hyperlink.setHistory(STOnOff.X_1);

		// 添加标题对应的序号
		/*CTR run = hyperlink.addNewR();
		CTRPr rPr = run.addNewRPr();
		rPr.addNewRStyle().setVal("ad");
		rPr.addNewNoProof();
		run.addNewT().setStringValue(chapterNum.toString());
		run = hyperlink.addNewR();
		rPr = run.addNewRPr();
		rPr.addNewSz().setVal(new BigInteger("21"));
		rPr.addNewSzCs().setVal(new BigInteger("22"));
		rPr.addNewNoProof();
		run.addNewTab();*/

		CTR run = hyperlink.addNewR();
		run.addNewRPr().addNewNoProof();
		run.addNewT().setStringValue(chapterNum + "  " + title);
		run = hyperlink.addNewR();
		CTRPr rPr = run.addNewRPr();
		rPr.addNewNoProof();
		rPr.addNewWebHidden();
		run.addNewTab();
		run = hyperlink.addNewR();
		rPr = run.addNewRPr();
		rPr.addNewNoProof();
		rPr.addNewWebHidden();
		run.addNewFldChar().setFldCharType(STFldCharType.BEGIN);
		// pageref run
		run = hyperlink.addNewR();
		rPr = run.addNewRPr();
		rPr.addNewNoProof();
		rPr.addNewWebHidden();
		CTText text = run.addNewInstrText();
		text.setSpace(SpaceAttribute.Space.PRESERVE);
		// bookmark reference
		text.setStringValue(" PAGEREF " + bookmarkRef + " \\h ");
		run = hyperlink.addNewR();
		rPr = run.addNewRPr();
		rPr.addNewNoProof();
		rPr.addNewWebHidden();
		run = hyperlink.addNewR();
		rPr = run.addNewRPr();
		rPr.addNewNoProof();
		rPr.addNewWebHidden();
		run.addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
		// page number run
		run = hyperlink.addNewR();
		rPr = run.addNewRPr();
		rPr.addNewNoProof();
		rPr.addNewWebHidden();
		run.addNewT().setStringValue(Integer.toString(page));
		run = hyperlink.addNewR();
		rPr = run.addNewRPr();
		rPr.addNewNoProof();
		rPr.addNewWebHidden();
		run.addNewFldChar().setFldCharType(STFldCharType.END);
	}

	public static int pageBreak(XWPFParagraph currentPar, String headTitle, List<IBodyElement> bodyElementList) {
		int pageBreak = 1;
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
							pageBreak++; //页数添加
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
						// 表格中最新一次数据所在的行号
						int lastBreakRow = 0;
						// 当前数据行号
						int rowindex = -1;
						for (XWPFTableRow row : tableList) {
							rowindex++;
							// 同一表格下，可能同一行有多个lastRenderedPageBreak
							boolean findBreakFlag = false;
							List<XWPFTableCell> cellList = row.getTableCells();
							if (cellList != null && cellList.size() > 0) {
								for (int m = 0; m < cellList.size() && !findBreakFlag; m++) {
									XWPFTableCell cell = cellList.get(m);
									List<XWPFParagraph> paragraphList = cell.getParagraphs();
									if (paragraphList != null && paragraphList.size() > 0) {
										for (int n = 0; n < paragraphList.size() && !findBreakFlag; n++) {
											XWPFParagraph par = paragraphList.get(n);
											List<CTR> ctrlist = par.getCTP().getRList();//获取<w:p>标签下的<w:r>list
											for (int j = 0; j < ctrlist.size() && !findBreakFlag; j++) {  //遍历r
												CTR r = ctrlist.get(j);
												List<CTEmpty> breaklist = r.getLastRenderedPageBreakList();//判断是否存在此标签
												if (breaklist.size() > 0) {
													findBreakFlag = true;
													// 处理分隔符在表格前后两行都存在的问题
													if (lastBreakRow == 0 || rowindex - lastBreakRow > 1) {
														lastBreakRow = rowindex;
														pageBreak++; //页数添加
														break;
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
		}
		if (!findFlag) {
			pageBreak = 1;
		}
		return pageBreak;
	}

	/**
	 * 用户手动更新目录
	 *
	 * @param document
	 * @param placeholder
	 */
	public static void handGenerateTOC(XWPFDocument document, String placeholder) {
		boolean findFlag = false;
		List<XWPFParagraph> list = document.getParagraphs();
		for (int i = 0; i < list.size() && !findFlag; i++) {
			XWPFParagraph p = list.get(i);
			String text = p.getParagraphText();
			if (text != null && text.toLowerCase().contains(placeholder.toLowerCase())) {
				// 清空当前段落的内容
				List<XWPFRun> runList = p.getRuns();
				for (int j = 0; j < runList.size(); j++) {
					XWPFRun r = runList.get(j);
					r.setText("", 0);
				}
				CTSimpleField ctSimpleField = p.getCTP().addNewFldSimple();
				ctSimpleField.setInstr("TOC \\o \"1-3\" \\h \\z \\u");
				// 强制在用户打开文档时，手动让用户更新目录结构
				ctSimpleField.setDirty(STOnOff.TRUE);
				break;
			}
		}
	}
}
