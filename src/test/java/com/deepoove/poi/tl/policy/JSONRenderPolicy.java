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
package com.deepoove.poi.tl.policy;

import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.policy.TextRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.util.StyleUtils;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.JsonNodeType;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map.Entry;

/**
 * highlighting json
 */
public class JSONRenderPolicy extends AbstractRenderPolicy<String> {

    ObjectMapper objectMapper = new ObjectMapper();

    private String defaultColor;

    public JSONRenderPolicy() {
        this("000000");
    }

    public JSONRenderPolicy(String defaultColor) {
        this.defaultColor = defaultColor;
    }

    @Override
    protected boolean validate(String data) {
        return null != data;
    }

    @Override
    protected void afterRender(RenderContext<String> context) {
        clearPlaceholder(context, false);
    }

    @Override
    public void doRender(RenderContext<String> context) throws Exception {
        XWPFRun run = context.getRun();
        String data = context.getData();
        JsonNode jsonNode = objectMapper.readTree(data);
        List<TextRenderData> codes = convert(jsonNode, 1);

        XWPFParagraph paragraph = (XWPFParagraph) run.getParent();
        codes.forEach(code -> {
            XWPFRun span = paragraph.createRun();
            StyleUtils.styleRun(span, run);
            TextRenderPolicy.Helper.renderTextRun(span, code);
        });

    }

    public List<TextRenderData> convert(JsonNode jsonNode, int level) {
        String indent = "";
        String cindent = "";
        for (int i = 0; i < level; i++) {
            indent += "    ";
            if (i != level - 1) cindent += "    ";
        }
        List<TextRenderData> result = new ArrayList<>();
        if (jsonNode.isValueNode()) {
            JsonNodeType nodeType = jsonNode.getNodeType();
            switch (nodeType) {
                case BOOLEAN:
                    // orange
                    result.add(new TextRenderData("FFB90F", jsonNode.booleanValue() + ""));
                    break;
                case NUMBER:
                    // red
                    result.add(new TextRenderData("FF6A6A", jsonNode.numberValue() + ""));
                    break;
                case NULL:
                    // red
                    result.add(new TextRenderData("FF6A6A", "null"));
                    break;
                case STRING:
                    // green
                    result.add(new TextRenderData("7CCD7C", "\"" + jsonNode.asText() + "\""));
                    break;
                default:
                    result.add(new TextRenderData(defaultColor, "\"" + jsonNode.asText() + "\""));
                    break;
            }
        } else if (jsonNode.isArray()) {
            result.add(new TextRenderData(defaultColor, "["));
            result.add(new TextRenderData("\n"));
            int size = jsonNode.size();
            for (int k = 0; k < size; k++) {
                JsonNode arrayItem = jsonNode.get(k);
                result.add(new TextRenderData(indent));
                result.addAll(convert(arrayItem, level + 1));
                if (k != size - 1) {
                    result.add(new TextRenderData(defaultColor, ","));
                }
                result.add(new TextRenderData("\n"));
            }
            result.add(new TextRenderData(defaultColor, cindent + "]"));
        } else if (jsonNode.isObject()) {
            Iterator<Entry<String, JsonNode>> fields = jsonNode.fields();
            result.add(new TextRenderData(defaultColor, "{"));
            result.add(new TextRenderData("\n"));
            boolean hasNext = fields.hasNext();
            while (true) {
                if (!hasNext) break;
                Entry<String, JsonNode> entry = fields.next();
                result.add(new TextRenderData(defaultColor, indent + "\"" + entry.getKey() + "\": "));
                JsonNode value = entry.getValue();
                result.addAll(convert(value, level + 1));
                hasNext = fields.hasNext();
                if (hasNext) {
                    result.add(new TextRenderData(defaultColor, ","));//
                }
                result.add(new TextRenderData("\n"));
            }
            result.add(new TextRenderData(defaultColor, cindent + "}"));
        }

        return result;
    }

}
