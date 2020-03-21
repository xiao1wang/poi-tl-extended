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

import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlObject;

/**
 * Paragraph is parent
 */
public interface ParagraphContext extends ParentContext {

    XWPFParagraph getParagraph();

    XWPFRun insertNewRun(XWPFRun xwpfRun, int insertPostionCursor);

    void setAndUpdateRun(XWPFRun xwpfRun2, XWPFRun insertNewRun, int insertPostionCursor);

    XWPFRun createRun(XmlObject object, IRunBody paragraph);

    XWPFRun createRun(XWPFRun xwpfRun, IRunBody paragraph);

    void removeRun(int pos);

}
