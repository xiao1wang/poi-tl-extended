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
package com.deepoove.poi.policy;

import com.deepoove.poi.data.NumbericRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.RenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.util.StyleUtils;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.math.BigInteger;
import java.util.List;

/**
 * @author Sayi
 * @version 0.0.5
 */
public class NumbericRenderPolicy extends AbstractRenderPolicy<NumbericRenderData> {

    @Override
    protected boolean validate(NumbericRenderData data) {
        return (null != data && CollectionUtils.isNotEmpty(data.getNumbers()));
    }

    @Override
    public void doRender(RenderContext<NumbericRenderData> context) throws Exception {
        Helper.renderNumberic(context.getRun(), context.getData());
    }

    @Override
    protected void afterRender(RenderContext<NumbericRenderData> context) {
        clearPlaceholder(context, true);
    }

    public static class Helper {

        public static void renderNumberic(XWPFRun run, NumbericRenderData numbericData)
                throws Exception {
            NiceXWPFDocument doc = (NiceXWPFDocument) run.getParent().getDocument();
            List<? extends RenderData> datas = numbericData.getNumbers();
            Style style = numbericData.getFmtStyle();

            BigInteger numID = doc.addNewNumbericId(numbericData.getNumFmt());

            BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(run);
            XWPFRun newRun;
            for (RenderData line : datas) {
                newRun = createRunLine(run, bodyContainer, style, numID);
                if (line instanceof PictureRenderData) {
                    PictureRenderPolicy.Helper.renderPicture(newRun, (PictureRenderData) line);
                } else if (line instanceof TextRenderData) {
                    TextRenderPolicy.Helper.renderTextRun(newRun, line);
                } else {
                    throw new RenderException(
                            "NumbericRender only support PictureRenderData and TextRenderData");
                }
            }
        }

        private static XWPFRun createRunLine(XWPFRun run, BodyContainer bodyContainer, Style style,
                                             BigInteger numID) {
            XWPFParagraph paragraph = bodyContainer.insertNewParagraph(run);
            StyleUtils.styleParagraph(paragraph, style);
            paragraph.setNumID(numID);
            return paragraph.createRun();
        }
    }
}
