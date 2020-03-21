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
package com.deepoove.poi;

import com.deepoove.poi.config.Configure;
import com.deepoove.poi.exception.ResolverException;
import com.deepoove.poi.policy.RenderPolicy;
import com.deepoove.poi.policy.ref.ReferenceRenderPolicy;
import com.deepoove.poi.render.DefaultRender;
import com.deepoove.poi.render.Render;
import com.deepoove.poi.resolver.Resolver;
import com.deepoove.poi.resolver.TemplateResolver;
import com.deepoove.poi.template.MetaTemplate;
import com.deepoove.poi.util.PoitlIOUtils;
import com.deepoove.poi.util.Preconditions;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.Version;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * The template of word(docx)
 *
 * @author Sayi
 * @version 0.0.1
 */
public class XWPFTemplate implements Closeable {
    private static Logger logger = LoggerFactory.getLogger(XWPFTemplate.class);
    private static final String SUPPORT_MINIMUM_VERSION = "4.0.0";

    private NiceXWPFDocument doc;

    private Configure config;
    private Resolver resolver;
    private Render renderer;

    private List<MetaTemplate> eleTemplates;

    static {
        Preconditions.checkMinimumVersion(Version.getVersion(), SUPPORT_MINIMUM_VERSION,
                (cur, min) -> "Require Apach POI version at least " + min + ", but now is " + cur
                        + ", please check the dependency of project.");
    }

    private XWPFTemplate() {
    }

    public static XWPFTemplate compile(String filePath) {
        return compile(new File(filePath));
    }

    public static XWPFTemplate compile(File file) {
        return compile(file, Configure.createDefault());
    }

    public static XWPFTemplate compile(InputStream inputStream) {
        return compile(inputStream, Configure.createDefault());
    }

    public static XWPFTemplate compile(String filePath, Configure config) {
        return compile(new File(filePath), config);
    }

    public static XWPFTemplate compile(File file, Configure config) {
        try {
            return compile(new FileInputStream(file), config);
        } catch (FileNotFoundException e) {
            throw new ResolverException("Cannot find the file [" + file.getPath() + "]", e);
        }
    }

    /**
     * template file as InputStream
     *
     * @param inputStream
     * @param config
     * @return
     * @version 1.2.0
     */
    public static XWPFTemplate compile(InputStream inputStream, Configure config) {
        try {
            XWPFTemplate template = new XWPFTemplate();
            template.config = config;
            template.doc = new NiceXWPFDocument(inputStream);
            template.resolver = new TemplateResolver(template.config);
            template.renderer = new DefaultRender();
            template.eleTemplates = template.resolver.resolveDocument(template.doc);
            return template;
        } catch (OLE2NotOfficeXmlFileException e) {
            logger.error("Poi-tl currently only supports .docx format");
            throw new ResolverException("Compile template failed", e);
        } catch (IOException e) {
            throw new ResolverException("Compile template failed", e);
        }
    }

    /**
     * Render the template by data model
     *
     * @param model
     * @return
     */
    public XWPFTemplate render(Object model) {
        this.renderer.render(this, model);
        return this;
    }

    /**
     * Render the template by data model and write to OutputStream, do'not
     * forget invoke #{@link XWPFTemplate#close()},
     * #{@link OutputStream#close()}
     *
     * @param model
     * @param out
     * @return
     * @throws IOException
     */
    public XWPFTemplate render(Object model, OutputStream out) throws IOException {
        this.render(model);
        this.write(out);
        return this;
    }

    /**
     * bind reference render policy
     *
     * @param refPolicy
     * @return
     */
    public XWPFTemplate bindRefPolicy(ReferenceRenderPolicy<?> refPolicy) {
        this.config.referencePolicy(refPolicy);
        return this;
    }

    /**
     * bind render policy
     *
     * @param tagName
     * @param policy
     * @return
     */
    public XWPFTemplate bind(String tagName, RenderPolicy policy) {
        this.config.customPolicy(tagName, policy);
        return this;
    }

    /**
     * write to output stream, do'not forget invoke
     * #{@link XWPFTemplate#close()}, #{@link OutputStream#close()} finally
     *
     * @param out eg.ServletOutputStream
     * @throws IOException
     */
    public void write(OutputStream out) throws IOException {
        this.doc.write(out);
    }

    /**
     * write to file
     *
     * @param path
     * @throws IOException
     */
    public void writeToFile(String path) throws IOException {
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(path);
            this.write(out);
            out.flush();
        } finally {
            PoitlIOUtils.closeQuietlyMulti(this.doc, out);
        }
    }

    /**
     * reload the template
     *
     * @param doc
     */
    public void reload(NiceXWPFDocument doc) {
        PoitlIOUtils.closeLoggerQuietly(this.doc);
        this.doc = doc;
        this.eleTemplates = this.resolver.resolveDocument(doc);
    }

    /**
     * close the document
     *
     * @throws IOException
     */
    @Override
    public void close() throws IOException {
        this.doc.close();
    }

    public List<MetaTemplate> getElementTemplates() {
        return eleTemplates;
    }

    public NiceXWPFDocument getXWPFDocument() {
        return this.doc;
    }

    public Configure getConfig() {
        return config;
    }

    public Resolver getResolver() {
        return resolver;
    }

}
