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
package com.deepoove.poi.policy.ref;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.exception.ReflectionException;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.SimpleValue;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.impl.CTNonVisualDrawingPropsImpl;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import javax.xml.namespace.QName;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * 可选文字匹配XWPFPicture
 *
 * @author Sayi
 * @version 1.6.0
 */
public abstract class OptionalTextPictureRefRenderPolicy extends ReferenceRenderPolicy<List<XWPFPicture>>
        implements OptionalText {

    protected static final QName TITLE = new QName("", "title");

    @Override
    protected List<XWPFPicture> locate(XWPFTemplate template) {
        logger.info("Try locate the XWPFPicture object which mathing optional text [{}]...", optionalText());
        List<XWPFPicture> optionalPictures = new ArrayList<>();
        NiceXWPFDocument document = template.getXWPFDocument();
        List<XWPFPicture> pictures = document.getAllEmbeddedPictures();
        for (XWPFPicture pic : pictures) {
            // it's array, to do in the future
            CTDrawing ctDrawing = getCTDrawing(pic);
            if (null == ctDrawing) continue;

            CTNonVisualDrawingProps docPr = null;
            if (ctDrawing.sizeOfAnchorArray() > 0) {
                CTAnchor anchorArray = ctDrawing.getAnchorArray(0);
                docPr = anchorArray.getDocPr();

            } else if (ctDrawing.sizeOfInlineArray() > 0) {
                CTInline inline = ctDrawing.getInlineArray(0);
                docPr = inline.getDocPr();
            }

            if (null != docPr) {
                String title = getTitle(docPr);
                if (Objects.equals(optionalText(), title)) {
                    optionalPictures.add(pic);
                    continue;
                }
                String descr = docPr.getDescr();
                if (Objects.equals(optionalText(), descr)) {
                    optionalPictures.add(pic);
                    continue;
                }
            }
        }
        return optionalPictures;

    }

    private String getTitle(CTNonVisualDrawingProps docPr) {
        CTNonVisualDrawingPropsImpl docPrImpl = (CTNonVisualDrawingPropsImpl) docPr;
        synchronized (docPrImpl.monitor()) {
            // check_orphaned();
            SimpleValue localSimpleValue = null;
            localSimpleValue = (SimpleValue) docPrImpl.get_store().find_attribute_user(TITLE);
            if (localSimpleValue == null) {
                return null;
            }
            return localSimpleValue.getStringValue();
        }
    }

    public CTDrawing getCTDrawing(XWPFPicture pic) throws RuntimeException {
        XWPFRun run;
        try {
            Field field = XWPFPicture.class.getDeclaredField("run");
            field.setAccessible(true);
            run = (XWPFRun) field.get(pic);
        } catch (Exception e) {
            throw new ReflectionException("run", XWPFPicture.class, e);
        }
        CTR ctr = run.getCTR();
        return ctr.getDrawingList() != null ? ctr.getDrawingArray(0) : null;
    }

}
