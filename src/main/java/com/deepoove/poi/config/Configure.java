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
package com.deepoove.poi.config;

import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.policy.DocxRenderPolicy;
import com.deepoove.poi.policy.MiniTableRenderPolicy;
import com.deepoove.poi.policy.NumbericRenderPolicy;
import com.deepoove.poi.policy.PictureRenderPolicy;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.policy.TextRenderPolicy;
import com.deepoove.poi.policy.ref.ReferenceRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.render.compute.DefaultRenderDataComputeFactory;
import com.deepoove.poi.render.compute.RenderDataComputeFactory;
import com.deepoove.poi.resolver.DefaultRunTemplateFactory;
import com.deepoove.poi.resolver.RunTemplateFactory;
import com.deepoove.poi.util.RegexUtils;
import com.github.xiao1wang.poitlextended.renderpolicy.ChartRenderPolicy;
import com.github.xiao1wang.poitlextended.renderpolicy.ListRenderPolicy;
import com.github.xiao1wang.poitlextended.renderpolicy.TableRenderPolicy;
import org.apache.commons.lang3.tuple.Pair;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * The config of template
 *
 * @author Sayi
 * @version 1.0.0
 */
public class Configure implements Cloneable {

    /**
     * regular expression: Chinese, letters, numbers and underscores
     */
    public static final String DEFAULT_GRAMER_REGEX = "((#)?[\\w\\u4e00-\\u9fa5]+(\\.[\\w\\u4e00-\\u9fa5]+)*)?";

    /**
     * template by bind: Highest priority
     */
    private final Map<String, RenderPolicy> CUSTOM_POLICYS = new HashMap<String, RenderPolicy>();

    /**
     * template by plugin: Low priority
     */
    private final Map<Character, RenderPolicy> DEFAULT_POLICYS = new HashMap<Character, RenderPolicy>();

    /**
     * template by reference
     */
    private final List<ReferenceRenderPolicy<?>> REFERENCE_POLICIES = new ArrayList<>();

    /**
     * if & for each
     * <p>
     * eg. {{?user}} Hello, World {{/user}}
     * </p>
     */
    Pair<Character, Character> iterable = Pair.of(GramerSymbol.ITERABLE_START.getSymbol(),
            GramerSymbol.BLOCK_END.getSymbol());

    /**
     * tag prefix
     */
    String gramerPrefix = "{{";

    /**
     * tag suffix
     */
    String gramerSuffix = "}}";

    /**
     * tag regular expression
     */
    String grammerRegex = DEFAULT_GRAMER_REGEX;

    /**
     * the mode of compute tag
     */
    ELMode elMode = ELMode.POI_TL_STANDARD_MODE;

    /**
     * the factory of render data compute
     */
    RenderDataComputeFactory renderDataComputeFactory = new DefaultRenderDataComputeFactory(this);

    /**
     * the factory of resovler run template
     */
    RunTemplateFactory<?> runTemplateFactory = new DefaultRunTemplateFactory(this);

    /**
     * the policy of process tag for valid render data error(null or illegal)
     */
    ValidErrorHandler handler = new ClearHandler();

    Configure() {
        plugin(GramerSymbol.TEXT, new TextRenderPolicy());
        plugin(GramerSymbol.TEXT_ALIAS, new TextRenderPolicy());
        plugin(GramerSymbol.IMAGE, new PictureRenderPolicy());
        plugin(GramerSymbol.TABLE, new MiniTableRenderPolicy());
        plugin(GramerSymbol.NUMBERIC, new NumbericRenderPolicy());
        plugin(GramerSymbol.DOCX_TEMPLATE, new DocxRenderPolicy());
        plugin(GramerSymbol.UCHART, new ChartRenderPolicy());
        plugin(GramerSymbol.ULIST, new ListRenderPolicy());
        plugin(GramerSymbol.UTABLE, new TableRenderPolicy());
    }

    /**
     * 创建默认配置
     *
     * @return
     */
    public static Configure createDefault() {
        return newBuilder().build();
    }

    /**
     * 构建器
     *
     * @return
     */
    public static ConfigureBuilder newBuilder() {
        return new ConfigureBuilder();
    }

    /**
     * 新增或变更语法插件
     *
     * @param c      语法
     * @param policy 策略
     */
    public Configure plugin(char c, RenderPolicy policy) {
        DEFAULT_POLICYS.put(Character.valueOf(c), policy);
        return this;
    }

    /**
     * 新增或变更语法插件
     *
     * @param symbol 语法
     * @param policy 策略
     * @return
     */
    Configure plugin(GramerSymbol symbol, RenderPolicy policy) {
        DEFAULT_POLICYS.put(symbol.getSymbol(), policy);
        return this;
    }

    /**
     * 自定义模板和策略
     *
     * @param tagName 模板名称
     * @param policy  策略
     */
    public void customPolicy(String tagName, RenderPolicy policy) {
        CUSTOM_POLICYS.put(tagName, policy);
    }

    /**
     * 新增引用渲染策略
     *
     * @param policy
     */
    public void referencePolicy(ReferenceRenderPolicy<?> policy) {
        REFERENCE_POLICIES.add(policy);
    }

    /**
     * 获取标签策略
     *
     * @param tagName 模板名称
     * @param sign    语法
     */
    public RenderPolicy getPolicy(String tagName, Character sign) {
        RenderPolicy policy = getCustomPolicy(tagName);
        return null == policy ? getDefaultPolicy(sign) : policy;
    }

    public List<ReferenceRenderPolicy<?>> getReferencePolicies() {
        return REFERENCE_POLICIES;
    }

    private RenderPolicy getCustomPolicy(String tagName) {
        return CUSTOM_POLICYS.get(tagName);
    }

    private RenderPolicy getDefaultPolicy(Character sign) {
        return DEFAULT_POLICYS.get(sign);
    }

    public Map<Character, RenderPolicy> getDefaultPolicys() {
        return DEFAULT_POLICYS;
    }

    public Map<String, RenderPolicy> getCustomPolicys() {
        return CUSTOM_POLICYS;
    }

    public Set<Character> getGramerChars() {
        Set<Character> ret = new HashSet<Character>(DEFAULT_POLICYS.keySet());
        // ? /
        ret.add(iterable.getKey());
        ret.add(iterable.getValue());
        return ret;
    }

    public String getGramerPrefix() {
        return gramerPrefix;
    }

    public String getGramerSuffix() {
        return gramerSuffix;
    }

    public String getGrammerRegex() {
        return grammerRegex;
    }

    public ELMode getElMode() {
        return elMode;
    }

    public ValidErrorHandler getValidErrorHandler() {
        return handler;
    }

    public RenderDataComputeFactory getRenderDataComputeFactory() {
        return renderDataComputeFactory;
    }

    public RunTemplateFactory<?> getRunTemplateFactory() {
        return runTemplateFactory;
    }

    public Pair<Character, Character> getIterable() {
        return iterable;
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append("Configure Info").append(":\n");
        sb.append("  Basic gramer: ").append(gramerPrefix).append(gramerSuffix).append("\n");
        sb.append("  If and foreach gramer: ").append(gramerPrefix).append(iterable.getLeft()).append(gramerSuffix);
        sb.append(gramerPrefix).append(iterable.getRight()).append(gramerSuffix).append("\n");
        sb.append("  EL Mode: ").append(elMode).append("\n");
        sb.append("  Regex:").append(grammerRegex).append("\n");
        sb.append("  Valid Error Handler: ").append(handler.getClass().getSimpleName()).append("\n");
        sb.append("  Default Plugin: ").append("\n");
        DEFAULT_POLICYS.forEach((chara, policy) -> {
            sb.append("    ").append(gramerPrefix).append(chara.charValue()).append(gramerSuffix);
            sb.append("->").append(policy.getClass().getSimpleName()).append("\n");
        });
        sb.append("  Bind Plugin: ").append("\n");
        CUSTOM_POLICYS.forEach((str, policy) -> {
            sb.append("    ").append(gramerPrefix).append(str).append(gramerSuffix);
            sb.append("->").append(policy.getClass().getSimpleName()).append("\n");
        });
        sb.append("  Reference Plugin: ").append("\n");
        REFERENCE_POLICIES.forEach(policy -> {
            sb.append("    ").append(policy.getClass().getSimpleName()).append("\n");
        });

        return sb.toString();
    }

    @Override
    protected Configure clone() throws CloneNotSupportedException {
        // shallow clone
        return (Configure) super.clone();
    }

    public Configure clone(String prefix, String suffix) throws CloneNotSupportedException {
        Configure clone = clone();
        clone.gramerPrefix = prefix;
        clone.gramerSuffix = suffix;
        if (clone.elMode == ELMode.SPEL_MODE || clone.elMode == ELMode.SIMPLE_SPEL_MODE) {
            clone.grammerRegex = RegexUtils.createGeneral(clone.gramerPrefix, clone.gramerSuffix);
        }
        return clone;
    }

    public enum ELMode {

        /**
         * 标准模式：无法计算表达式时，RenderData默认为null值
         */
        POI_TL_STANDARD_MODE,
        /**
         * 严格模式：无法计算表达式直接抛出异常
         */
        POI_TL_STICT_MODE,
        /**
         * 简单Spring EL模式(不抛出异常)
         */
        SIMPLE_SPEL_MODE,
        /**
         * Spring EL模式
         */
        SPEL_MODE;

    }

    public interface ValidErrorHandler {
        void handler(RenderContext<?> context);
    }

    public static class DiscardHandler implements ValidErrorHandler {
        @Override
        public void handler(RenderContext<?> context) {
            // no-op
        }
    }

    public static class ClearHandler implements ValidErrorHandler {
        @Override
        public void handler(RenderContext<?> context) {
            context.getRun().setText("", 0);
        }
    }

    public static class AbortHandler implements ValidErrorHandler {
        @Override
        public void handler(RenderContext<?> context) {
            throw new RenderException("Non-existent variable and a variable with illegal value for "
                    + context.getTagSource() + ", data: " + context.getData());
        }
    }

}
