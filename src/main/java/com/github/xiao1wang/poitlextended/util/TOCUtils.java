package com.github.xiao1wang.poitlextended.util;

import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.util.POILogFactory;
import org.apache.poi.util.POILogger;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTEmpty;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSimpleField;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTTextImpl;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * TODO : 目录更新工具类
 */
public class TOCUtils {

	private static final POILogger LOG = POILogFactory.getLogger(XWPFDocument.class);

	/**
	 * 更新目录
	 *
	 * @param doc        当前文档对象
	 * @param maxLevel   需要展示标题的等级数
	 * @param fromPageNo 记录文章页数的开始位置，由于目录的特殊性，没有分页符，只能通过这个参数调整了
	 */
	public static void updateItem2TOC(NiceXWPFDocument doc, Integer maxLevel, int fromPageNo) {
		List<IBodyElement> bodyElementList = doc.getBodyElements();
        if (bodyElementList != null && bodyElementList.size() > 0) {
            // 存放word文档，每页对应的标题名称
            Map<Integer, Set<String>> pageMap = new TreeMap<>();
            List<XWPFParagraph> paragraphList = new ArrayList<>();
            for (IBodyElement element : bodyElementList) {
                if (element instanceof XWPFParagraph) {
                    XWPFParagraph par = (XWPFParagraph) element;
                    String parStyle = par.getStyle();
                    if (parStyle != null && StringUtils.isNumeric(parStyle)) {
                        int level = Integer.parseInt(parStyle);
                        if (level <= maxLevel) {
                            // 该执行块用于获取属于标题的段落
                            String title = par.getText();
                            // 获取当前标题名称，是在文档的第几页出现的
                            int page = pageIndex(par, title, bodyElementList, fromPageNo);
                            if (pageMap.containsKey(page)) {
                                Set<String> list = pageMap.get(page);
                                list.add(title);
                            } else {
                                Set<String> list = new HashSet<>();
                                list.add(title);
                                pageMap.put(page, list);
                            }
                        }
                    }
                    // 用于获取目录的段落
                    List<CTHyperlink> linkList = par.getCTP().getHyperlinkList();
                    if (linkList.size() > 0) {
                        paragraphList.add(par);
                    }
                }
            }
            Pattern pattern = Pattern.compile("\\d+");
            // 更新页码
            for (XWPFParagraph par : paragraphList) {
	            String title = par.getParagraphText();
	            CTHyperlink ctHyperlink = par.getCTP().getHyperlinkList().get(0);
	            List<CTR> rList = null;
	            // 判断文字后面是否还有数字，如果没有数据，说明是另外一种结构
	            // 得到当前文本最后一个数字
	            Matcher matcher = pattern.matcher(title);
	            int size = -1;
	            while (matcher.find()) {
		            size = Integer.parseInt(matcher.group());
	            }
	            if (size == -1) {
		            List<CTSimpleField> list = ctHyperlink.getFldSimpleList();
		            if (list != null && list.size() > 0) {
			            rList = list.get(0).getRList();
		            }
	            } else {
		            rList = ctHyperlink.getRList();
	            }
	            if (rList != null && rList.size() > 0) {
		            boolean find = false;
		            int pageNo = -1;
		            Iterator<Integer> pageIte = pageMap.keySet().iterator();
		            while (pageIte.hasNext() && !find) {
			            Integer page = pageIte.next();
			            Set<String> titleList = pageMap.get(page);
			            if (titleList != null && titleList.size() > 0) {
                            Iterator<String> titleIte = titleList.iterator();
                            while (titleIte.hasNext() && !find) {
                                String content = titleIte.next();
                                if (title.indexOf(content) != -1) {
                                    if (rList != null && rList.size() > 0) {
                                        // 得到最后一个r中只包含数据的位置
                                        int rIndex = -1;
                                        for (int i = 0; i < rList.size(); i++) {
                                            CTR r = rList.get(i);
                                            List<CTText> tList = r.getTList();
                                            if (tList != null && tList.size() > 0) {
                                                String val = ((CTTextImpl) tList.get(0)).getStringValue();
                                                if (val != null && StringUtils.isNumeric(val)) {
                                                    rIndex = i;
                                                }
                                            }
                                        }
                                        if (rIndex != -1) {
                                            find = true;
                                            pageNo = page;
                                            List<CTText> tList = rList.get(rIndex).getTList();
                                            if (tList != null && tList.size() > 0) {
                                                tList.get(0).setStringValue(page + "");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    if (find) {
                        // 为了防止有相同标题的文字
                        Set<String> list = pageMap.get(pageNo);
                        Set<String> newSet = new HashSet<>();
                        Iterator<String> titleIte = list.iterator();
                        while (titleIte.hasNext()) {
                            String currentTitle = titleIte.next();
                            if (title.indexOf(currentTitle) == -1) {
                                newSet.add(currentTitle);
                            }
                        }
                        pageMap.put(pageNo, newSet);
                    }
                }
            }
        }
    }

    public static int pageIndex(XWPFParagraph currentPar, String headTitle, List<IBodyElement> bodyElementList, int fromPageNo) {
        int page = fromPageNo;
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
	                        // 同一表格下，可能有多个lastRenderedPageBreak
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
							                        page++; //页数添加
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
        if (!findFlag) {
            page = 1;
        }
        return page;
    }

}
