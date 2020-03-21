/*
 * Copyright 2014-2020 Sayi
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package com.deepoove.poi.xwpf;

import com.deepoove.poi.util.ParagraphUtils;
import com.deepoove.poi.util.ReflectionUtils;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;

import java.util.List;

public class CellBodyContainer implements BodyContainer {

    XWPFTableCell cell;

    public CellBodyContainer(XWPFTableCell cell) {
        this.cell = cell;
    }

    @Override
    public int getPosOfParagraphCTP(CTP startCtp) {
        IBodyElement current;
        List<IBodyElement> bodyElements = cell.getBodyElements();
        for (int i = 0; i < bodyElements.size(); i++) {
            current = bodyElements.get(i);
            if (current.getElementType() == BodyElementType.PARAGRAPH) {
                if (((XWPFParagraph) current).getCTP().equals(startCtp)) {
                    return i;
                }
            }
        }
        return -1;
    }

    @SuppressWarnings("unchecked")
    @Override
    public void removeBodyElement(int pos) {
        // TODO remove Nest Table
        List<IBodyElement> bodyElements = getBodyElements();
        if (pos >= 0 && pos < bodyElements.size()) {
            BodyElementType type = bodyElements.get(pos).getElementType();
            if (type == BodyElementType.TABLE) {
                int indexOf = cell.getTables().indexOf(bodyElements.get(pos));
                // remove cell's table
                List<XWPFTable> tables = (List<XWPFTable>) ReflectionUtils.getValue("tables", cell);
                tables.remove(indexOf);
                cell.getCTTc().removeTbl(indexOf);
            }
            if (type == BodyElementType.PARAGRAPH) {
                int indexOf = cell.getParagraphs().indexOf(bodyElements.get(pos));
                // cell.removeParagraph may has bug
                // cell.removeParagraph(indexOf);
                List<XWPFParagraph> paragraphs = (List<XWPFParagraph>) ReflectionUtils.getValue("paragraphs", cell);
                paragraphs.remove(indexOf);
                cell.getCTTc().removeP(indexOf);
            }
            bodyElements.remove(pos);
        }

    }

    @Override
    public int getPosOfParagraph(XWPFParagraph startParagraph) {
        return getPosOfParagraphCTP(startParagraph.getCTP());
    }

    @SuppressWarnings("unchecked")
    @Override
    public List<IBodyElement> getBodyElements() {
        return (List<IBodyElement>) ReflectionUtils.getValue("bodyElements", cell);
    }

    @Override
    public XWPFParagraph insertNewParagraph(XmlCursor insertPostionCursor) {
        return cell.insertNewParagraph(insertPostionCursor);
    }

    @Override
    public int getParaPos(XWPFParagraph insertNewParagraph) {
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        for (int i = 0; i < paragraphs.size(); i++) {
            if (paragraphs.get(i) == insertNewParagraph) {
                return i;
            }
        }
        return -1;
    }

    @SuppressWarnings("unchecked")
    @Override
    public void setParagraph(XWPFParagraph p, int paraPos) {
        List<XWPFParagraph> paragraphs = (List<XWPFParagraph>) ReflectionUtils.getValue("paragraphs", cell);
        paragraphs.set(paraPos, p);
        CTTc ctTc = cell.getCTTc();
        ctTc.setPArray(paraPos, p.getCTP());

    }

    @Override
    public IBody getTarget() {
        return cell;
    }

    @Override
    public void updateBodyElements(IBodyElement insertNewParagraph, IBodyElement copy) {
        int pos = -1;
        List<IBodyElement> bodyElements = getBodyElements();
        for (int i = 0; i < bodyElements.size(); i++) {
            if (bodyElements.get(i) == insertNewParagraph) {
                pos = i;
            }
        }
        if (-1 != pos) bodyElements.set(pos, copy);

    }

    @Override
    public XWPFTable insertNewTbl(XmlCursor insertPostionCursor) {
        return cell.insertNewTbl(insertPostionCursor);
    }

    @Override
    public int getTablePos(XWPFTable insertNewTbl) {
        List<XWPFTable> tables = cell.getTables();
        for (int i = 0; i < tables.size(); i++) {
            if (tables.get(i) == insertNewTbl) {
                return i;
            }
        }
        return -1;
    }

    @SuppressWarnings("unchecked")
    @Override
    public void setTable(int pos, XWPFTable table) {
        // cell.getTables().set(pos, table);
        List<XWPFTable> tables = (List<XWPFTable>) ReflectionUtils.getValue("tables", cell);
        tables.set(pos, table);
        cell.getCTTc().setTblArray(pos, table.getCTTbl());

    }

    @Override
    public XWPFTable insertNewTable(XWPFRun run, int row, int col) {
        XmlCursor cursor = ((XWPFParagraph) run.getParent()).getCTP().newCursor();
        XWPFTable table = insertNewTbl(cursor);

        // hack for cursor.removeXmlContents(); in XWPFTableCell
        List<XWPFTableRow> rows = table.getRows();
        for (int i = 0; i < rows.size(); i++) {
            table.removeRow(i);
        }
        for (int i = 0; i < row; i++) {
            XWPFTableRow tabRow = (table.getRow(i) == null) ? table.createRow() : table.getRow(i);
            for (int k = 0; k < col; k++) {
                if (tabRow.getCell(k) == null) {
                    tabRow.createCell();
                }
            }
        }
        return table;
    }

    @Override
    public void clearPlaceholder(XWPFRun run) {
        IRunBody parent = run.getParent();
        run.setText("", 0);
        // 遇到不明确的单元格匹配问题, 可能丢失段落元素，<p>元素必须位于</tc>元素之前
        if (parent instanceof XWPFParagraph) {
            String paragraphText = ParagraphUtils.trimLine((XWPFParagraph) parent);
            if ("".equals(paragraphText)) {
                int pos = getPosOfParagraph((XWPFParagraph) parent);
                int lastPos = cell.getBodyElements().size() - 1;
                if (canRemoveParagraph(pos, lastPos)) {
                    removeBodyElement(pos);
                }
            }
        }
    }

    private boolean canRemoveParagraph(int pos, int lastPos) {
        return pos < lastPos
                || (pos > 0 && cell.getBodyElements().get(pos - 1).getElementType() == BodyElementType.PARAGRAPH);
    }

}
