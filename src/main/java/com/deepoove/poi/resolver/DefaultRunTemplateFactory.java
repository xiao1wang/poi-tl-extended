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

package com.deepoove.poi.resolver;

import com.deepoove.poi.config.Configure;
import com.deepoove.poi.template.run.RunTemplate;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.Set;

/**
 * @author Sayi
 */
public class DefaultRunTemplateFactory implements RunTemplateFactory<RunTemplate> {

    public static final char EMPTY_CHAR = '\0';
    private final Configure config;

    public DefaultRunTemplateFactory(Configure config) {
        this.config = config;
    }

    @Override
    public RunTemplate createRunTemplate(String tag, XWPFRun run) {
        RunTemplate template = new RunTemplate();
        Set<Character> gramerChars = config.getGramerChars();
        Character symbol = Character.valueOf(EMPTY_CHAR);
        if (!"".equals(tag)) {
            char fisrtChar = tag.charAt(0);
            for (Character chara : gramerChars) {
                if (chara.equals(fisrtChar)) {
                    symbol = Character.valueOf(fisrtChar);
                    break;
                }
            }
        }
        template.setSource(config.getGramerPrefix() + tag + config.getGramerSuffix());
        template.setTagName(symbol.equals(Character.valueOf(EMPTY_CHAR)) ? tag : tag.substring(1));
        template.setSign(symbol);
        template.setRun(run);
        return template;
    }

}
